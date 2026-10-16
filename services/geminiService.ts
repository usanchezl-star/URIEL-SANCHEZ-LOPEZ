
import { GoogleGenAI, Type } from "@google/genai";
import { ERFDocument } from "../types";

const erfSchema = {
  type: Type.OBJECT,
  properties: {
    metadata: {
      type: Type.OBJECT,
      properties: {
        nombre: { type: Type.STRING },
        codigo: { type: Type.STRING }
      },
      required: ["nombre", "codigo"]
    },
    identificacion: {
      type: Type.OBJECT,
      properties: {
        sistema: { type: Type.STRING },
        modulo: { type: Type.STRING },
        proceso: { type: Type.STRING },
        descripcion: { type: Type.STRING },
        requeridoPor: { type: Type.STRING }
      },
      required: ["sistema", "proceso"]
    },
    antecedentes: { type: Type.STRING },
    requerimiento: {
      type: Type.OBJECT,
      properties: {
        objetivo: { type: Type.STRING },
        necesidad: { type: Type.STRING }
      },
      required: ["objetivo", "necesidad"]
    },
    reglasNegocio: {
      type: Type.ARRAY,
      items: { type: Type.STRING }
    },
    criteriosAceptacion: {
      type: Type.ARRAY,
      items: { type: Type.STRING },
      description: "Minimum essential criteria that the process must meet to be considered successful/accepted."
    },
    pasosProceso: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          numero_paso: { type: Type.NUMBER, description: "El número correlativo." },
          instruccion: { type: Type.STRING, description: "El texto del paso." },
          timestamp: { type: Type.STRING, description: "The exact minute and second (MM:SS) where the action described in the audio VISUALLY appears on the screen. Wait for the visual match." },
          resultadoEsperado: { type: Type.STRING, description: "The immediate and literal response of the system after the user action." },
          id_imagen: { type: Type.STRING, description: "El número de orden de la imagen que ilustra este paso (ej: 'imagen_1' si es la primera que aparece en el documento)." }
        },
        required: ["numero_paso", "instruccion"]
      }
    },
    anexos: {
      type: Type.ARRAY,
      items: { type: Type.STRING },
      description: "List of filenames provided in the source that serve as technical support."
    },
    hipervinculos: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          etiqueta: { type: Type.STRING },
          url: { type: Type.STRING }
        },
        required: ["etiqueta", "url"]
      },
      description: "List of external links (Google Drive, etc.) mentioned in the context."
    },
    transcripcion: { 
      type: Type.STRING, 
      description: "LITERAL and COMPLETE transcription of everything said in the video audio. Do not summarize, do not omit anything." 
    },
    missingFields: {
      type: Type.ARRAY,
      items: { type: Type.STRING },
      description: "List of information pieces missing from the source that are required for a complete ERF"
    }
  },
  required: ["identificacion", "requerimiento", "pasosProceso", "criteriosAceptacion", "anexos"]
};

export const extractFrameFromVideo = async (videoFile: File, timestamp: string, quality: number = 0.7): Promise<string> => {
  return new Promise((resolve) => {
    const video = document.createElement('video');
    const videoUrl = URL.createObjectURL(videoFile);
    video.src = videoUrl;
    video.crossOrigin = 'anonymous';
    video.muted = true;
    video.playsInline = true;
    
    const parts = timestamp.split(':');
    if (parts.length !== 2) {
      URL.revokeObjectURL(videoUrl);
      resolve('');
      return;
    }
    
    let seconds = parseInt(parts[0]) * 60 + parseInt(parts[1]);
    let retries = 0;
    const MAX_RETRIES = 10;
    const RETRY_OFFSET = 0.1;

    const cleanup = () => {
      URL.revokeObjectURL(videoUrl);
      video.pause();
      video.src = "";
      video.load();
      video.remove();
    };

    const checkBrightness = (ctx: CanvasRenderingContext2D, width: number, height: number): boolean => {
      try {
        const imageData = ctx.getImageData(0, 0, width, height);
        const data = imageData.data;
        let totalBrightness = 0;
        const step = 40; // More aggressive sampling
        let count = 0;
        
        for (let i = 0; i < data.length; i += 4 * step) {
          totalBrightness += (0.299 * data[i] + 0.587 * data[i+1] + 0.114 * data[i+2]);
          count++;
        }
        return (totalBrightness / count) > 8;
      } catch (e) { return true; }
    };

    video.onloadedmetadata = () => {
      video.currentTime = Math.min(seconds, video.duration);
    };
    
    video.onseeked = () => {
      const canvas = document.createElement('canvas');
      // Reduced resolution for preview optimization
      const scale = 0.7; 
      canvas.width = video.videoWidth * scale;
      canvas.height = video.videoHeight * scale;
      const ctx = canvas.getContext('2d', { alpha: false });
      
      if (ctx) {
        ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
        if (checkBrightness(ctx, canvas.width, canvas.height) || retries >= MAX_RETRIES) {
          const dataUrl = canvas.toDataURL('image/jpeg', quality);
          cleanup();
          resolve(dataUrl);
        } else {
          retries++;
          seconds += RETRY_OFFSET;
          video.currentTime = Math.min(seconds, video.duration);
        }
      } else {
        cleanup();
        resolve('');
      }
    };
    
    video.onerror = () => { cleanup(); resolve(''); };
    setTimeout(() => { cleanup(); resolve(''); }, 15000);
  });
};

export interface MultimodalPart {
  inlineData: {
    data: string;
    mimeType: string;
  };
}

export const generateTemporalIndex = async (
  videoFile: File,
  duration: number
): Promise<{ timestamp: string; label: string }[]> => {
  const apiKey = process.env.GEMINI_API_KEY || process.env.API_KEY || '';
  const ai = new GoogleGenAI({
    apiKey,
    httpOptions: {
      headers: {
        'User-Agent': 'aistudio-build',
      },
    },
  });
  const modelName = 'gemini-3.8-flash';

  // Extract frames for indexing. Adjust interval based on duration to keep frame count reasonable (~40-60 frames)
  const interval = Math.max(30, Math.floor(duration / 50));
  const framePromises: Promise<string>[] = [];
  for (let t = 0; t < duration; t += interval) {
    const mm = Math.floor(t / 60).toString().padStart(2, '0');
    const ss = (t % 60).toString().padStart(2, '0');
    framePromises.push(extractFrameFromVideo(videoFile, `${mm}:${ss}`, 0.5));
  }

  const frames = await Promise.all(framePromises);
  const validFrames = frames.filter(f => f !== '');

  const mediaParts = validFrames.map((f, i) => ({
    inlineData: {
      data: f.split(',')[1],
      mimeType: 'image/jpeg'
    }
  }));

  const systemPrompt = `
    Eres un experto en indexación de video. Tu tarea es analizar una secuencia de capturas de un video de proceso 
    y generar un ÍNDICE TEMPORAL de los momentos clave.
    
    Debes identificar cambios significativos en la pantalla (cambio de pestaña, apertura de modal, carga de datos, etc.).
    
    Responde estrictamente en formato JSON:
    {
      "index": [
        { "timestamp": "MM:SS", "label": "Descripción breve del cambio o hito" }
      ]
    }
  `;

  const contents = {
    parts: [
      ...mediaParts,
      { text: "Analiza estas capturas secuenciales y genera el índice temporal del proceso." }
    ]
  };

  try {
    const response = await ai.models.generateContent({
      model: modelName,
      contents: contents,
      config: {
        systemInstruction: systemPrompt,
        responseMimeType: "application/json"
      }
    });

    if (!response || !response.text) return [];
    const data = JSON.parse(response.text.trim());
    return data.index || [];
  } catch (error) {
    console.error("Error generating temporal index:", error);
    return [];
  }
};

export const analyzeProcess = async (
  textContext: string,
  mediaParts: MultimodalPart[] = [],
  fileList: string[] = [],
  manualTimestamps: string = "",
  chunkInfo?: { index: number; total: number; startTime: number; endTime: number; videoDuration: number },
  previousContext?: { lastStepNumber: number; summary: string },
  isAudioOnly: boolean = false,
  hasVideo: boolean = true
): Promise<Partial<ERFDocument>> => {
  const apiKey = process.env.GEMINI_API_KEY || process.env.API_KEY || '';
  const ai = new GoogleGenAI({
    apiKey,
    httpOptions: {
      headers: {
        'User-Agent': 'aistudio-build',
      },
    },
  });
  const modelName = 'gemini-3.8-flash';
  
  const modeInstruction = isAudioOnly 
    ? "\n\nMODO: ANÁLISIS DE AUDIO Y CONTEXTO. Concéntrate en la transcripción literal y en identificar los pasos del proceso basados en el audio y los documentos proporcionados. Ignora los detalles visuales minuciosos por ahora para optimizar el procesamiento."
    : "";

  const chunkInstruction = chunkInfo 
    ? `\n\nESTÁS PROCESANDO EL BLOQUE ${chunkInfo.index + 1} DE ${chunkInfo.total} del video.
       El tiempo de inicio de este bloque es ${Math.floor(chunkInfo.startTime / 60)}:${(chunkInfo.startTime % 60).toString().padStart(2, '0')}.
       El tiempo final de este bloque es ${Math.floor(chunkInfo.endTime / 60)}:${(chunkInfo.endTime % 60).toString().padStart(2, '0')}.
       La duración TOTAL del video es ${Math.floor(chunkInfo.videoDuration / 60)}:${(chunkInfo.videoDuration % 60).toString().padStart(2, '0')}.
       ${previousContext ? `Continuidad: El último paso del bloque anterior fue el #${previousContext.lastStepNumber}. Resumen previo: ${previousContext.summary}.` : ''}
       IMPORTANTE: 
       1. Continúa la numeración de los pasos desde el #${(previousContext?.lastStepNumber || 0) + 1}.
       2. Los timestamps que generes deben ser RELATIVOS AL INICIO DEL VIDEO ORIGINAL (usa el tiempo de inicio proporcionado como base).
       3. Asegúrate de que el primer paso de este bloque conecte lógicamente con el final del anterior.
       4. PROHIBIDO ALUCINAR TIEMPOS: NINGÚN timestamp generado en este bloque puede ser mayor a ${Math.floor(chunkInfo.endTime / 60)}:${(chunkInfo.endTime % 60).toString().padStart(2, '0')} ni mayor a la duración total del video.`
    : "";

  const manualTimestampsInstruction = manualTimestamps 
    ? `\n\nATENCIÓN: El usuario ha proporcionado una lista de TIMESTAMPS MANUALES: ${manualTimestamps}. 
       IGNORA la detección automática de escenas para los pasos del proceso y utiliza ÚNICAMENTE estos tiempos exactos para generar los pasos. 
       Cada tiempo proporcionado DEBE ser un paso en el JSON de respuesta.`
    : "";

  const videoRules = hasVideo ? `
    REGLAS DE ORO DE PRECISIÓN:
    1. PRIORIDAD AL AUDIO (CRÍTICO): Analiza PRIMERO la transcripción del audio para definir qué está sucediendo. Lo que se dice en el audio es la fuente de verdad principal para la secuencia de pasos.
    2. SINCRONIZACIÓN AUDIO-VISUAL: Para cada paso identificado en el audio, verifica que la pantalla muestre EXACTAMENTE lo que se describe.
       - Si el audio describe una acción pero la pantalla aún no la muestra, ESPERA y busca el momento exacto en el video donde la pantalla COINCIDE con la descripción del audio.
       - El timestamp (MM:SS) debe corresponder al momento VISUAL que coincide con la descripción del AUDIO.
    3. EVITA REPETICIONES (CRÍTICO): No generes pasos duplicados. Si una acción ya fue descrita, no la vuelvas a incluir. Cada paso debe representar una acción nueva y única. Revisa el resumen previo para no repetir el último paso del bloque anterior.
    4. NO OMITAS NADA: Cada clic, campo o mensaje del sistema mencionado y mostrado DEBE aparecer como un paso.
    5. TIMESTAMPS OBLIGATORIOS: Cada paso debe tener el tiempo exacto (MM:SS) donde la imagen coincide con el audio.
    6. FIDELIDAD TOTAL: No inventes lógica. Documenta lo observado y escuchado.
    7. DESCRIPCIÓN TÉCNICA: Usa nombres literales de la interfaz mencionados en el audio.
    8. PASO A PASO ATÓMICO: Cada acción es un paso independiente.
    9. TRANSCRIPCIÓN LITERAL Y TOTAL (CRÍTICO): Genera una transcripción palabra por palabra de TODO el audio del video en este bloque, desde el inicio hasta el final, en el campo 'transcripcion'. 
       - PROHIBIDO RESUMIR: No sintetices ni omitas frases. 
       - INTEGRIDAD: La transcripción debe reflejar todo el audio escuchado en este fragmento.
    10. INFORMACIÓN FALTANTE: Si no encuentras información suficiente para completar una sección específica, deja el campo estrictamente vacío ("").${manualTimestampsInstruction}

    INSTRUCCIÓN ESPECÍFICA PARA DOCUMENTOS PREVIOS (WORD/PDF) + VIDEO:
    - Si se proporcionan documentos de texto junto con el video, UTILIZA EL DOCUMENTO COMO BASE para homologar el formato, antecedentes, reglas de negocio y estructura.
    - COMPLEMENTA el flujo de proceso del documento base con las NUEVAS ACCIONES o modificaciones detectadas en el video.
    - Extrae capturas frescas del video para los pasos del proceso asignando el timestamp correspondiente.
    - Analiza el archivo (PDF/Word) adjunto. Debes extraer cada paso del procedimiento y asociarlo correctamente con la captura de pantalla que le corresponde según el contexto del texto o su posición original.
    - Guidelines para la extracción:
      * Identificación de Pasos: Extrae el texto instructivo de forma literal pero organizada.
      * Asociación de Imágenes: Identifica la imagen más cercana a cada instrucción. Si el texto dice "Haga clic en el botón X" y hay una imagen de ese botón justo debajo o encima, esa es su captura asociada.
      * En el documento de salida, no omitas ninguna imagen para ninguno de los 3 formatos (.mp4, .pdf ni word).
    - Para cada paso detectado, genera un objeto JSON que contenga:
      numero_paso: El número correlativo.
      instruccion: El texto del paso.
      id_imagen: El número de orden de la imagen que ilustra este paso (ej: 'imagen_1' si es la primera que aparece en el documento).
  ` : `
    REGLAS DE ORO DE HOMOLOGACIÓN (SOLO DOCUMENTO):
    1. HOMOLOGACIÓN ESTRICTA: Tu tarea principal es leer el documento proporcionado y adaptar su contenido a la estructura JSON solicitada.
    2. SIN TIMESTAMPS: Como NO hay video, NO generes ni inventes timestamps (MM:SS). Deja el campo 'timestamp' completamente vacío ("") o no lo incluyas en los pasos del proceso.
    3. EXTRACCIÓN DE PASOS E IMÁGENES: Analiza el archivo (PDF/Word) adjunto. Debes extraer cada paso del procedimiento y asociarlo correctamente con la captura de pantalla que le corresponde según el contexto del texto o su posición original.
       - Guidelines para la extracción:
         * Identificación de Pasos: Extrae el texto instructivo de forma literal pero organizada.
         * Asociación de Imágenes: Identifica la imagen más cercana a cada instrucción. Si el texto dice "Haga clic en el botón X" y hay una imagen de ese botón justo debajo o encima, esa es su captura asociada.
         * En el documento de salida, no omitas ninguna imagen para ninguno de los 3 formatos (.mp4, .pdf ni word).
       - Para cada paso detectado, genera un objeto JSON que contenga:
         numero_paso: El número correlativo.
         instruccion: El texto del paso.
         id_imagen: El número de orden de la imagen que ilustra este paso (ej: 'imagen_1' si es la primera que aparece en el documento).
    4. FIDELIDAD AL TEXTO: Mantén la redacción original lo más fiel posible, adaptándola solo para que encaje en el formato de pasos atómicos.
    5. INFORMACIÓN FALTANTE: Si el documento no contiene información para ciertas secciones (ej. reglas de negocio, criterios de aceptación), deja esos campos vacíos ("").
  `;

  const systemPrompt = `
    Eres un Analista de Sistemas Senior experto en Documentación Técnica para Liverpool. 
    Tu tarea es generar un Flujo de Proceso EXTREMADAMENTE PRECISO y GRANULAR basado UNICAMENTE en el material proporcionado.${modeInstruction}${chunkInstruction}

    ${videoRules}

    INSTRUCCIÓN ESPECÍFICA PARA ANEXOS (PDFs/Documentos):
    - Identifica qué archivos sirven como soporte (Diagramas, Manuales, etc.).
    - Coloca ÚNICAMENTE el NOMBRE DEL ARCHIVO original en el array de 'anexos'.

    Archivos disponibles para mapear: ${fileList.join(', ')}

    Responde estrictamente en formato JSON siguiendo el esquema proporcionado.
  `;

  const contents = {
    parts: [
      ...mediaParts,
      { text: `CONTEXTO ADICIONAL DEL USUARIO:\n${textContext}\n\nAnaliza el material y genera el Flujo de Proceso. Si hay PDFs, asígnalos como anexos usando su nombre de archivo exacto.` }
    ]
  };

  try {
    const response = await ai.models.generateContent({
      model: modelName,
      contents: contents,
      config: {
        systemInstruction: systemPrompt,
        responseMimeType: "application/json",
        responseSchema: erfSchema
      }
    });

    if (!response || !response.text) {
      return {};
    }

    const text = response.text.trim();
    const parsed = JSON.parse(text);
    
    // Mapear los campos del prompt de vuelta al formato ERFDocument
    if (parsed.pasosProceso && Array.isArray(parsed.pasosProceso)) {
      parsed.pasosProceso = parsed.pasosProceso.map((p: any) => ({
        paso: p.numero_paso || p.paso,
        descripcion: p.instruccion || p.descripcion,
        timestamp: p.timestamp,
        resultadoEsperado: p.resultadoEsperado,
        referenciaImagen: p.id_imagen || p.referenciaImagen
      }));
    }
    
    return parsed;
  } catch (error) {
    console.error("Gemini Error:", error);
    throw error;
  }
};
