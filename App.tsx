
import React, { useState, useRef } from 'react';
import { ERFDocument, createNewERF } from './types';
import Layout from './components/Layout';
import { analyzeProcess, MultimodalPart, extractFrameFromVideo, generateTemporalIndex } from './services/geminiService';
import { saveToHistory, getHistory, HistoryItem, removeFromHistory } from './services/historyService';
import { downloadERFWord } from './services/wordService';
import { exportToPDF } from './services/pdfService';
import * as mammoth from 'mammoth/mammoth.browser';
import * as pdfjsLib from 'pdfjs-dist';

// Configure pdfjs worker using a stable version (3.x) to avoid ES module worker issues
const pdfjs = pdfjsLib && (pdfjsLib as any).default ? (pdfjsLib as any).default : pdfjsLib;
if (pdfjs && pdfjs.GlobalWorkerOptions) {
  pdfjs.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js`;
} else {
  console.error("pdfjsLib not loaded correctly", pdfjsLib);
}

const App: React.FC = () => {
  const [doc, setDoc] = useState<ERFDocument>(createNewERF());
  const [activeSection, setActiveSection] = useState('input');
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [isIndexing, setIsIndexing] = useState(false);
  const [isExportingPDF, setIsExportingPDF] = useState(false);
  const [inputText, setInputText] = useState('');
  const [manualTimestamps, setManualTimestamps] = useState<{ step: number; time: string }[]>([]);
  const [temporalIndex, setTemporalIndex] = useState<{ timestamp: string; label: string }[]>([]);
  const [uploadedFiles, setUploadedFiles] = useState<File[]>([]);
  const [stepsConfirmed, setStepsConfirmed] = useState(false);
  const [refreshingStep, setRefreshingStep] = useState<number | null>(null);
  const [analysisProgress, setAnalysisProgress] = useState({ current: 0, total: 0 });
  const [isProcessingAnnex, setIsProcessingAnnex] = useState(false);
  const [history, setHistory] = useState<HistoryItem[]>(getHistory());
  const [showHistoryPrompt, setShowHistoryPrompt] = useState(false);
  const [hideHistoryBox, setHideHistoryBox] = useState(false);
  const [commandText, setCommandText] = useState('');
  const [deleteConfirmation, setDeleteConfirmation] = useState<{ isOpen: boolean; itemId: string | null }>({ isOpen: false, itemId: null });
  
  const previewRef = useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    const storedHistory = getHistory();
    if (storedHistory.length > 0) {
      setShowHistoryPrompt(true);
    }
  }, []);

  const extractImagesFromPdf = async (file: File): Promise<{nombre: string, contenido: string}[]> => {
    const arrayBuffer = await file.arrayBuffer();
    const pdfjs = pdfjsLib && (pdfjsLib as any).default ? (pdfjsLib as any).default : pdfjsLib;
    const pdf = await pdfjs.getDocument({ data: arrayBuffer }).promise;
    const extractedImages: {nombre: string, contenido: string}[] = [];
    let imgCount = 1;
    
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const ops = await page.getOperatorList();
      
      for (let j = 0; j < ops.fnArray.length; j++) {
        const fn = ops.fnArray[j];
        if (fn === pdfjs.OPS.paintImageXObject || fn === pdfjs.OPS.paintInlineImageXObject) {
          const imgId = ops.argsArray[j][0];
          try {
            // En pdfjs-dist, getOperatorList no siempre resuelve los objetos síncronamente
            const img = await new Promise<any>((resolve) => {
              try {
                page.objs.get(imgId, (image: any) => resolve(image));
              } catch (err) {
                resolve(null);
              }
            });
            
            if (img && img.data && img.width && img.height) {
              const canvas = document.createElement('canvas');
              canvas.width = img.width;
              canvas.height = img.height;
              const ctx = canvas.getContext('2d');
              if (ctx) {
                let imageData: ImageData | null = null;
                const data = img.data;
                const width = img.width;
                const height = img.height;
                
                // Manejar diferentes espacios de color
                if (data.length === width * height * 4) {
                  // RGBA
                  imageData = new ImageData(new Uint8ClampedArray(data), width, height);
                } else if (data.length === width * height * 3) {
                  // RGB -> RGBA
                  const rgba = new Uint8ClampedArray(width * height * 4);
                  for (let k = 0, l = 0; k < data.length; k += 3, l += 4) {
                    rgba[l] = data[k];
                    rgba[l+1] = data[k+1];
                    rgba[l+2] = data[k+2];
                    rgba[l+3] = 255;
                  }
                  imageData = new ImageData(rgba, width, height);
                } else if (data.length === width * height) {
                  // Escala de grises -> RGBA
                  const rgba = new Uint8ClampedArray(width * height * 4);
                  for (let k = 0, l = 0; k < data.length; k++, l += 4) {
                    const val = data[k];
                    rgba[l] = val;
                    rgba[l+1] = val;
                    rgba[l+2] = val;
                    rgba[l+3] = 255;
                  }
                  imageData = new ImageData(rgba, width, height);
                }
                
                if (imageData) {
                  ctx.putImageData(imageData, 0, 0);
                  extractedImages.push({
                    nombre: `${file.name.replace('.pdf', '')} - Pág ${i} - Img ${imgCount++}`,
                    contenido: canvas.toDataURL('image/jpeg', 0.9)
                  });
                }
              }
            }
          } catch (e) {
            console.warn("No se pudo extraer una imagen del PDF", e);
          }
        }
      }
    }
    return extractedImages;
  };

  const extractImagesFromDocx = async (file: File): Promise<{nombre: string, contenido: string}[]> => {
    const arrayBuffer = await file.arrayBuffer();
    const extractedImages: {nombre: string, contenido: string}[] = [];
    let imgCount = 1;
    
    const mammothObj = mammoth && (mammoth as any).default ? (mammoth as any).default : mammoth || (window as any).mammoth;
    if (!mammothObj || !mammothObj.images) {
      console.error("Mammoth library not loaded correctly", mammoth);
      return [];
    }
    
    const options = {
      convertImage: mammothObj.images.imgElement(function(image: any) {
        return image.read("base64").then(function(imageBuffer: string) {
          const dataUrl = "data:" + image.contentType + ";base64," + imageBuffer;
          extractedImages.push({
            nombre: `${file.name.replace('.docx', '')} - Img ${imgCount++}`,
            contenido: dataUrl
          });
          return { src: dataUrl };
        });
      })
    };
    
    await mammothObj.convertToHtml({ arrayBuffer }, options);
    return extractedImages;
  };

  const handleAnnexUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!e.target.files) return;
    setIsProcessingAnnex(true);
    try {
      const files = Array.from(e.target.files) as File[];
      const newAnnexes = [...doc.anexos];
      
      for (const file of files) {
        if (file.type.includes('image')) {
          const base64 = await new Promise<string>((resolve) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result as string);
            reader.readAsDataURL(file);
          });
          newAnnexes.push({ nombre: file.name, contenido: base64 });
        } else if (file.type === 'application/pdf') {
          const images = await extractImagesFromPdf(file);
          images.forEach((img) => {
            newAnnexes.push({ 
              nombre: img.nombre, 
              contenido: img.contenido 
            });
          });
        }
      }
      setDoc({ ...doc, anexos: newAnnexes });
    } catch (err) {
      console.error("Error processing annex:", err);
      alert("Error al procesar el anexo.");
    } finally {
      setIsProcessingAnnex(false);
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      setUploadedFiles(prev => [...prev, ...Array.from(e.target.files!)]);
    }
  };

  const removeFile = (index: number) => {
    setUploadedFiles(prev => prev.filter((_, i) => i !== index));
  };

  const updateDoc = <K extends keyof ERFDocument>(key: K, value: Partial<ERFDocument[K]>) => {
    setDoc(prev => ({ 
      ...prev, 
      [key]: typeof value === 'object' && !Array.isArray(value) ? { ...(prev[key] as object), ...value } : value 
    }));
  };

  const getVideoDuration = (file: File): Promise<number> => {
    return new Promise((resolve) => {
      const video = document.createElement('video');
      video.preload = 'metadata';
      video.onloadedmetadata = () => {
        const duration = video.duration;
        URL.revokeObjectURL(video.src);
        resolve(duration);
      };
      video.onerror = () => resolve(0);
      video.src = URL.createObjectURL(file);
    });
  };

  const processWithAI = async () => {
    if (uploadedFiles.length === 0 && !inputText) {
      alert("Por favor carga al menos un archivo o ingresa contexto.");
      return;
    }
    
    const videoFile = uploadedFiles.find(f => f.type.includes('video'));
    const isLargeOrLong = videoFile && (videoFile.size > 200 * 1024 * 1024 || (await getVideoDuration(videoFile)) > 1800);

    // 1. Prioritize Temporal Index for large/long videos
    if (isLargeOrLong && temporalIndex.length === 0) {
      setIsIndexing(true);
      try {
        const duration = await getVideoDuration(videoFile);
        const index = await generateTemporalIndex(videoFile, duration);
        setTemporalIndex(index);
        // We don't stop here, but we've prioritized it. 
        // We could show it to the user, but the prompt says "prioritize creation... before deep analysis"
        // so we'll continue with the deep analysis using this index.
      } catch (err) {
        console.error("Indexing error:", err);
      } finally {
        setIsIndexing(false);
      }
    }

    setIsAnalyzing(true);
    try {
      let consolidatedText = inputText;
      const mediaParts: MultimodalPart[] = [];
      const fileNames = uploadedFiles.map(f => f.name);
      const extractedDocumentImages: {nombre: string, contenido: string}[] = [];
      
      // 1. Process non-video files
      for (const file of uploadedFiles) {
        if (file.type.includes('video')) continue;
        const extension = file.name.split('.').pop()?.toLowerCase();
        if (extension === 'docx') {
          const arrayBuffer = await file.arrayBuffer();
          const mammothObj = mammoth && (mammoth as any).default ? (mammoth as any).default : mammoth || (window as any).mammoth;
          if (mammothObj && mammothObj.extractRawText) {
            const result = await mammothObj.extractRawText({ arrayBuffer });
            consolidatedText += `\n\n[Doc: ${file.name}]:\n${result.value}`;
          } else {
            console.error("Mammoth library not loaded correctly for text extraction", mammoth);
          }
          
          // Extraer imágenes del DOCX
          try {
            const docxImages = await extractImagesFromDocx(file);
            extractedDocumentImages.push(...docxImages);
          } catch (imgErr) {
            console.error(`Error extracting images from ${file.name}:`, imgErr);
          }
        } else {
          const base64 = await new Promise<string>((resolve) => {
            const reader = new FileReader();
            reader.onload = () => resolve((reader.result as string).split(',')[1]);
            reader.readAsDataURL(file);
          });
          mediaParts.push({ inlineData: { data: base64, mimeType: file.type || 'application/pdf' } });
          
          // Extraer imágenes del PDF
          if (file.type === 'application/pdf') {
            try {
              const pdfImages = await extractImagesFromPdf(file);
              pdfImages.forEach((img) => {
                extractedDocumentImages.push({ 
                  nombre: img.nombre, 
                  contenido: img.contenido 
                });
              });
            } catch (imgErr) {
              console.error(`Error extracting images from ${file.name}:`, imgErr);
            }
          }
        }
      }

      if (extractedDocumentImages.length > 0) {
        // Renombrar a imagen_1, imagen_2, etc. para que coincida con el prompt
        extractedDocumentImages.forEach((img, idx) => {
          img.nombre = `imagen_${idx + 1}`;
        });
        consolidatedText += `\n\n[IMÁGENES EXTRAÍDAS DISPONIBLES]:\n${extractedDocumentImages.map(img => img.nombre).join(', ')}\nSi un paso corresponde a alguna de estas imágenes, usa su nombre exacto en el campo 'id_imagen'.`;
      }

      const manualTimestampsString = [
        ...manualTimestamps.filter(mt => mt.time.trim() !== '').map(mt => `Paso ${mt.step}: ${mt.time}`),
        ...temporalIndex.map(ti => `Hito: ${ti.timestamp} - ${ti.label}`)
      ].join(', ');

      let finalDoc: Partial<ERFDocument> = {};
      const CHUNK_DURATION = 300; // 5 minutes for better precision in long videos
      
      if (videoFile) {
        const duration = await getVideoDuration(videoFile);
        const totalChunks = Math.ceil(duration / CHUNK_DURATION);
        setAnalysisProgress({ current: 0, total: totalChunks });
        
        let allSteps: any[] = [];
        let fullTranscription = "";
        let lastStepNumber = 0;
        let lastSummary = "";

        for (let i = 0; i < totalChunks; i++) {
          setAnalysisProgress(prev => ({ ...prev, current: i + 1 }));
          const startTime = i * CHUNK_DURATION;
          const endTime = Math.min((i + 1) * CHUNK_DURATION, duration);
          
          const partialDoc = await (async () => {
            let chunkMediaParts: MultimodalPart[] = [...mediaParts];

            if (videoFile.size > 50 * 1024 * 1024) {
              // FOR LARGE VIDEOS: Don't send the whole video. 
              // Send a sequence of frames for this chunk (1 frame every 5 seconds)
              const frameInterval = 5;
              const framePromises: Promise<string>[] = [];
              for (let t = startTime; t < endTime; t += frameInterval) {
                const mm = Math.floor(t / 60).toString().padStart(2, '0');
                const ss = (t % 60).toString().padStart(2, '0');
                framePromises.push(extractFrameFromVideo(videoFile, `${mm}:${ss}`, 0.6));
              }
              const frames = await Promise.all(framePromises);
              frames.filter(f => f !== '').forEach(f => {
                chunkMediaParts.push({ inlineData: { data: f.split(',')[1], mimeType: 'image/jpeg' } });
              });
            } else {
              // SMALL VIDEOS: Send the whole video for better context
              const videoBase64 = await new Promise<string>((resolve) => {
                const reader = new FileReader();
                reader.onload = () => resolve((reader.result as string).split(',')[1]);
                reader.readAsDataURL(videoFile);
              });
              chunkMediaParts.push({ inlineData: { data: videoBase64, mimeType: videoFile.type } });
            }

            const result = await analyzeProcess(
              consolidatedText, 
              chunkMediaParts, 
              fileNames, 
              manualTimestampsString,
              { index: i, total: totalChunks, startTime, endTime, videoDuration: duration },
              i > 0 ? { lastStepNumber, summary: lastSummary } : undefined,
              true // isAudioOnly = true
            );
            
            return result;
          })();

          if (i === 0) {
            finalDoc = { ...partialDoc };
          }
          
          if (partialDoc.pasosProceso) {
            const uniqueNewSteps = [];
            for (const step of partialDoc.pasosProceso) {
              // Filter out steps with timestamps beyond the video duration
              if (step.timestamp) {
                const parts = step.timestamp.split(':');
                if (parts.length === 2) {
                  const mm = parseInt(parts[0], 10);
                  const ss = parseInt(parts[1], 10);
                  if (!isNaN(mm) && !isNaN(ss)) {
                    const stepTimeInSeconds = mm * 60 + ss;
                    if (stepTimeInSeconds > duration) {
                      continue; // Skip this step as it's an hallucination beyond video length
                    }
                  }
                }
              }

              const lastAdded = uniqueNewSteps.length > 0 
                ? uniqueNewSteps[uniqueNewSteps.length - 1] 
                : (allSteps.length > 0 ? allSteps[allSteps.length - 1] : null);
              
              if (!lastAdded) {
                uniqueNewSteps.push(step);
                continue;
              }

              // Consider duplicate if description is identical or timestamp is identical
              const isSameDesc = step.descripcion.trim().toLowerCase() === lastAdded.descripcion.trim().toLowerCase();
              const isSameTime = step.timestamp && lastAdded.timestamp && step.timestamp === lastAdded.timestamp;
              
              if (!isSameDesc && !isSameTime) {
                uniqueNewSteps.push(step);
              }
            }

            allSteps = [...allSteps, ...uniqueNewSteps];
            const lastStep = allSteps[allSteps.length - 1];
            lastStepNumber = lastStep?.paso || lastStepNumber;
            lastSummary = lastStep?.descripcion || lastSummary;
          }
          
          if (partialDoc.transcripcion) {
            fullTranscription += (fullTranscription ? "\n\n" : "") + `[BLOQUE ${i+1} (${Math.floor(startTime/60)}m - ${Math.floor(endTime/60)}m)]:\n` + partialDoc.transcripcion;
          }
          
          await new Promise(r => setTimeout(r, 500));
        }
        
        // Re-number steps sequentially after deduplication
        finalDoc.pasosProceso = allSteps.map((step, index) => ({ ...step, paso: index + 1 }));
        finalDoc.transcripcion = fullTranscription;
      } else {
        finalDoc = await analyzeProcess(
          consolidatedText, 
          mediaParts, 
          fileNames, 
          manualTimestampsString,
          undefined,
          undefined,
          false, // isAudioOnly
          false  // hasVideo
        );
      }
      
      // LAZY CAPTURE: We do NOT extract any frames initially.
      // The user will click "Ver Captura" to load them individually.

      setDoc(prev => {
        let processedSteps = finalDoc.pasosProceso || prev.pasosProceso;
        const mappedImageNames = new Set<string>();
        
        if (processedSteps) {
          processedSteps = processedSteps.map((step: any) => {
            if (step.referenciaImagen) {
              const matchedImage = extractedDocumentImages.find(img => img.nombre === step.referenciaImagen);
              if (matchedImage) {
                mappedImageNames.add(step.referenciaImagen);
                return { ...step, pantallaAsociada: matchedImage.contenido };
              }
            }
            return step;
          });
        }

        const aiAnexos = (finalDoc.anexos || []).map((a: any) => typeof a === 'string' ? { nombre: a } : a);
        const mergedAnexos = [...prev.anexos, ...aiAnexos, ...extractedDocumentImages.filter(img => !mappedImageNames.has(img.nombre))];
        
        // Remove duplicates by name just in case
        const uniqueAnexos = mergedAnexos.filter((anexo, index, self) =>
          index === self.findIndex((t) => t.nombre === anexo.nombre)
        );

        const updatedDoc = {
          ...prev,
          metadata: { ...prev.metadata, ...finalDoc.metadata },
          identificacion: { ...prev.identificacion, ...finalDoc.identificacion },
          antecedentes: finalDoc.antecedentes || prev.antecedentes,
          requerimiento: { ...prev.requerimiento, ...finalDoc.requerimiento },
          reglasNegocio: finalDoc.reglasNegocio || prev.reglasNegocio,
          criteriosAceptacion: finalDoc.criteriosAceptacion || prev.criteriosAceptacion,
          pasosProceso: processedSteps,
          anexos: uniqueAnexos,
          hipervinculos: finalDoc.hipervinculos || prev.hipervinculos,
          transcripcion: finalDoc.transcripcion || prev.transcripcion
        };
        const newHistory = saveToHistory(updatedDoc);
        setHistory(newHistory);
        return updatedDoc;
      });
      setStepsConfirmed(false);
      setActiveSection('steps');
    } catch (err) {
      console.error("AI Analysis error:", err);
      alert("Error procesando con IA.");
    } finally { setIsAnalyzing(false); }
  };

  const handleDownloadWord = async () => {
    const videoFile = uploadedFiles.find(f => f.type.includes('video'));
    
    // Before downloading, we MUST ensure all steps have their frames
    // If they don't, we extract them now (Lazy Capture on final doc)
    setIsExportingPDF(true); // Using this as a generic loading state
    try {
      const updatedPasos = await Promise.all(doc.pasosProceso.map(async (p) => {
        if (videoFile && p.timestamp && !p.pantallaAsociada) {
          try {
            const frame = await extractFrameFromVideo(videoFile, p.timestamp, 0.8);
            return { ...p, pantallaAsociada: frame };
          } catch (e) { return p; }
        }
        return p;
      }));

      const finalDoc = { ...doc, pasosProceso: updatedPasos };
      await downloadERFWord(finalDoc);
    } catch (e) {
      console.error("Word generation error:", e);
      alert("Error al generar el documento Word.");
    } finally {
      setIsExportingPDF(false);
    }
  };

  const handleRefreshFrame = async (index: number) => {
    const videoFile = uploadedFiles.find(f => f.type.includes('video'));
    const step = doc.pasosProceso[index];
    if (!videoFile || !step.timestamp) {
      alert("Se requiere un video cargado y un timestamp válido.");
      return;
    }

    setRefreshingStep(index);
    try {
      const frameBase64 = await extractFrameFromVideo(videoFile, step.timestamp);
      const newPasos = [...doc.pasosProceso];
      newPasos[index].pantallaAsociada = frameBase64;
      setDoc({ ...doc, pasosProceso: newPasos });
    } catch (e) {
      console.error("Error refreshing frame:", e);
      alert("Error al extraer la captura.");
    } finally {
      setRefreshingStep(null);
    }
  };

  const handleAutoLoadCaptures = async () => {
    const videoFile = uploadedFiles.find(f => f.type.includes('video'));
    if (!videoFile) {
      alert("Se requiere un video cargado para extraer capturas.");
      return;
    }

    if (!window.confirm("Esto procesará todas las capturas automáticamente basado en los timestamps detectados. ¿Deseas continuar?")) {
      return;
    }

    setRefreshingStep(-1); // Special state for bulk processing
    setAnalysisProgress({ current: 0, total: doc.pasosProceso.length });

    try {
      const newPasos = [...doc.pasosProceso];
      
      // Process sequentially to avoid browser memory issues with too many canvas operations
      for (let i = 0; i < newPasos.length; i++) {
        const step = newPasos[i];
        if (step.timestamp && !step.pantallaAsociada) {
          try {
            // Use the optimized timestamp from AI which aligns audio description with visual
            const frameBase64 = await extractFrameFromVideo(videoFile, step.timestamp, 0.7);
            if (frameBase64) {
              newPasos[i].pantallaAsociada = frameBase64;
              // Update state incrementally to show progress
              setDoc(prev => ({ ...prev, pasosProceso: [...newPasos] }));
            }
          } catch (e) {
            console.error(`Error extracting frame for step ${step.paso}:`, e);
          }
        }
        setAnalysisProgress(prev => ({ ...prev, current: i + 1 }));
        // Small delay to yield to UI thread
        await new Promise(r => setTimeout(r, 100));
      }
    } catch (e) {
      console.error("Error in bulk frame extraction:", e);
      alert("Hubo un error durante la extracción automática.");
    } finally {
      setRefreshingStep(null);
    }
  };

  const handleDownloadPDF = async () => {
    if (!previewRef.current) return;
    const videoFile = uploadedFiles.find(f => f.type.includes('video'));
    
    setIsExportingPDF(true);
    try {
      // Lazy Capture: Ensure all frames are present before PDF export
      const updatedPasos = await Promise.all(doc.pasosProceso.map(async (p) => {
        if (videoFile && p.timestamp && !p.pantallaAsociada) {
          try {
            const frame = await extractFrameFromVideo(videoFile, p.timestamp, 0.8);
            return { ...p, pantallaAsociada: frame };
          } catch (e) { return p; }
        }
        return p;
      }));

      // Temporarily update doc for preview rendering
      setDoc(prev => ({ ...prev, pasosProceso: updatedPasos }));
      
      // Wait for React to re-render with new images
      await new Promise(r => setTimeout(r, 1000));

      const fileName = `ERF_${doc.identificacion.sistema || "SYS"}_${(doc.identificacion.proceso || "PROCESO").replace(/\s+/g, "_")}.pdf`;
      await exportToPDF(previewRef.current, fileName);
    } catch (err) {
      console.error("PDF generation error:", err);
      alert("Error exportando a PDF.");
    } finally {
      setIsExportingPDF(false);
    }
  };

  const handleDeleteStep = (index: number) => {
    const newPasos = doc.pasosProceso
      .filter((_, i) => i !== index)
      .map((p, i) => ({ ...p, paso: i + 1 }));
    setDoc({ ...doc, pasosProceso: newPasos });
    setStepsConfirmed(false);
  };

  const handleAddStepBetween = (index: number) => {
    const newPasos = [...doc.pasosProceso];
    newPasos.splice(index + 1, 0, { 
      paso: 0, 
      descripcion: '', 
      timestamp: '', 
      resultadoEsperado: '', 
      ocultarDetalles: false 
    });
    const renumberedPasos = newPasos.map((p, i) => ({ ...p, paso: i + 1 }));
    setDoc({ ...doc, pasosProceso: renumberedPasos });
    setStepsConfirmed(false);
  };

  const handleOpenFromHistory = (item: HistoryItem) => {
    setDoc(item.fullDoc);
    setActiveSection('steps');
    setShowHistoryPrompt(false);
  };

  const handleDeleteHistory = (e: React.MouseEvent, id: string) => {
    e.stopPropagation();
    setDeleteConfirmation({ isOpen: true, itemId: id });
  };

  const confirmDeleteHistory = () => {
    if (deleteConfirmation.itemId) {
      const newHistory = removeFromHistory(deleteConfirmation.itemId);
      setHistory(newHistory);
      setDeleteConfirmation({ isOpen: false, itemId: null });
    }
  };

  const cancelDeleteHistory = () => {
    setDeleteConfirmation({ isOpen: false, itemId: null });
  };

  const handleNewDocument = () => {
    if (window.confirm('¿Estás seguro de que deseas empezar un nuevo documento? Se perderán los cambios no guardados en el historial.')) {
      setDoc(createNewERF());
      setUploadedFiles([]);
      setInputText('');
      setManualTimestamps([]);
      setTemporalIndex([]);
      setStepsConfirmed(false);
      setActiveSection('input');
    }
  };

  const handleCommand = (e: React.FormEvent) => {
    e.preventDefault();
    const cmd = commandText.toLowerCase().trim();
    if (cmd.startsWith('abrir documento ')) {
      const target = cmd.replace('abrir documento ', '').trim();
      // Try by index first
      const index = parseInt(target) - 1;
      if (!isNaN(index) && history[index]) {
        handleOpenFromHistory(history[index]);
      } else {
        // Try by title
        const item = history.find(h => h.title.toLowerCase().includes(target));
        if (item) {
          handleOpenFromHistory(item);
        } else {
          alert('Documento no encontrado en el historial.');
        }
      }
    }
    setCommandText('');
  };

  const renderHistoryWindow = () => (
    <div className="mt-12 pt-8 border-t border-slate-200">
      <div className="bg-slate-900 text-white p-8 rounded-2xl shadow-2xl">
        <div className="flex items-center justify-between mb-6">
          <h3 className="text-xs font-black uppercase tracking-[0.3em] text-slate-400 flex items-center">
            <i className="fas fa-history mr-3 text-pink-500"></i>
            Ventana de Historial
          </h3>
          <span className="text-[10px] bg-slate-800 text-slate-400 px-2 py-1 rounded font-bold">Últimos 5 documentos</span>
        </div>
        
        <div className="space-y-3">
          {history.length === 0 ? (
            <div className="py-8 text-center border border-dashed border-slate-700 rounded-xl">
              <p className="text-slate-500 text-sm italic">No hay documentos en el historial.</p>
            </div>
          ) : (
            <div className="space-y-3">
              {history.map((item, idx) => (
                <div key={item.id} className="relative group">
                  <button 
                    onClick={() => handleOpenFromHistory(item)}
                    className="w-full text-left p-4 bg-slate-800/50 border border-slate-700 rounded-xl hover:border-pink-500/50 hover:bg-slate-800 transition-all"
                  >
                    <div className="flex items-start space-x-4 pr-8">
                      <span className="text-pink-500 font-black text-sm mt-0.5">{idx + 1}.</span>
                      <div className="flex-1 min-w-0">
                        <h4 className="font-bold text-sm text-slate-100 group-hover:text-white transition-colors truncate">
                          {item.title}
                        </h4>
                        <p className="text-slate-500 text-[11px] mt-1 line-clamp-1 font-medium">
                          {item.summary}
                        </p>
                      </div>
                    </div>
                  </button>
                  <button 
                    onClick={(e) => handleDeleteHistory(e, item.id)}
                    className="absolute right-4 top-1/2 -translate-y-1/2 text-slate-600 hover:text-red-500 p-2 transition-colors opacity-0 group-hover:opacity-100 z-10"
                    title="Eliminar del historial"
                  >
                    <i className="fas fa-trash-alt text-xs"></i>
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>
        
        <form onSubmit={handleCommand} className="mt-8 relative">
          <input 
            type="text" 
            placeholder='Escribe "Abrir documento [Número o Título]"'
            className="w-full bg-slate-800 border border-slate-700 rounded-xl pl-12 pr-24 py-3 text-sm focus:ring-2 focus:ring-pink-500 focus:border-transparent outline-none transition-all"
            value={commandText}
            onChange={(e) => setCommandText(e.target.value)}
          />
          <div className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-500">
            <i className="fas fa-terminal text-xs"></i>
          </div>
          <button type="submit" className="absolute right-2 top-1/2 -translate-y-1/2 bg-pink-600 hover:bg-pink-700 px-4 py-1.5 rounded-lg text-xs font-bold transition-colors shadow-lg shadow-pink-900/20">
            Ejecutar
          </button>
        </form>
      </div>
    </div>
  );

  const renderSection = () => {
    const commonInputClass = "w-full p-3 bg-slate-50 border rounded-lg text-sm focus:ring-2 focus:ring-pink-500 outline-none";
    const labelClass = "text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1 block";

    switch (activeSection) {
      case 'input':
        return (
          <div className="max-w-4xl mx-auto space-y-6">
            {history.length > 0 && !hideHistoryBox && (
              <div className="bg-pink-50 border border-pink-100 p-6 rounded-2xl flex items-center justify-between animate-in fade-in slide-in-from-top-4 duration-500">
                <div className="flex items-center space-x-4">
                  <div className="w-12 h-12 bg-pink-600 rounded-full flex items-center justify-center text-white shadow-lg shadow-pink-200">
                    <i className="fas fa-history text-xl"></i>
                  </div>
                  <div>
                    <h3 className="text-sm font-bold text-pink-900">¿Deseas recuperar tu trabajo anterior?</h3>
                    <p className="text-xs text-pink-600 mt-0.5 font-medium">Tienes {history.length} {history.length === 1 ? 'documento guardado' : 'documentos guardados'} en el historial.</p>
                  </div>
                </div>
                <div className="flex items-center space-x-4">
                  <button 
                    onClick={() => setHideHistoryBox(true)}
                    className="text-xs font-bold text-pink-400 hover:text-pink-600 transition-colors"
                  >
                    Ignorar
                  </button>
                  <button 
                    onClick={() => setActiveSection('history')}
                    className="bg-pink-600 text-white px-6 py-2.5 rounded-xl text-xs font-bold hover:bg-pink-700 transition-all shadow-md shadow-pink-200"
                  >
                    Ver Historial
                  </button>
                </div>
              </div>
            )}

            <div className="bg-white p-8 rounded-2xl border shadow-sm">
              <h2 className="text-xl font-bold mb-4">Carga de Soporte</h2>
              <p className="text-xs text-slate-500 mb-6 italic">Sube grabaciones de pantalla para el flujo y documentos PDF/Word como soporte técnico.</p>
              <label className="flex flex-col items-center justify-center w-full h-32 border-2 border-dashed rounded-xl cursor-pointer hover:bg-slate-50 transition-all mb-4 group">
                <i className="fas fa-cloud-upload-alt text-2xl text-pink-600 mb-2 group-hover:scale-110 transition-transform"></i>
                <span className="text-sm font-medium text-slate-600">Click para subir archivos (Video, PDF, Word)</span>
                <input type="file" className="hidden" multiple onChange={handleFileUpload} />
              </label>
              <div className="grid grid-cols-2 gap-2 mb-4">
                {uploadedFiles.map((f, i) => (
                  <div key={i} className="flex justify-between items-center bg-slate-50 p-3 rounded-lg border text-xs">
                    <div className="flex items-center space-x-2 truncate">
                       <i className={`fas ${f.type.includes('video') ? 'fa-video text-blue-500' : 'fa-file-pdf text-red-500'} w-4`}></i>
                       <span className="truncate">{f.name}</span>
                    </div>
                    <button onClick={() => removeFile(i)} className="text-slate-300 hover:text-red-500 transition-colors ml-2"><i className="fas fa-trash"></i></button>
                  </div>
                ))}
              </div>

              {temporalIndex.length > 0 && (
                <div className="mb-6 p-4 bg-slate-50 rounded-xl border border-slate-200 animate-in fade-in slide-in-from-top-4">
                  <h3 className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-3 flex items-center">
                    <i className="fas fa-list-ol mr-2 text-pink-500"></i> Índice Temporal Generado
                  </h3>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-2">
                    {temporalIndex.map((ti, idx) => (
                      <div key={idx} className="flex items-center space-x-2 text-[11px] bg-white p-2 rounded border">
                        <span className="font-mono font-bold text-pink-600">{ti.timestamp}</span>
                        <span className="text-slate-600 truncate">{ti.label}</span>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              <button 
                onClick={processWithAI} 
                disabled={isAnalyzing || isIndexing} 
                className="w-full mt-6 py-4 bg-pink-600 text-white rounded-xl font-bold hover:bg-pink-700 disabled:opacity-50 shadow-lg shadow-pink-200 active:scale-[0.98] transition-all"
              >
                {isIndexing ? (
                  <><i className="fas fa-list-ol fa-spin mr-2"></i> Generando Índice Temporal...</>
                ) : isAnalyzing ? (
                  <><i className="fas fa-spinner fa-spin mr-2"></i> Analizando Bloque {analysisProgress.current}/{analysisProgress.total}...</>
                ) : (
                  <><i className="fas fa-wand-magic-sparkles mr-2"></i> Generar Documentación con IA</>
                )}
              </button>
            </div>
          </div>
        );

      case 'header':
        return (
          <div className="max-w-4xl mx-auto bg-white p-8 rounded-2xl border shadow-sm space-y-6">
            <h2 className="text-xl font-bold border-b pb-4">Identificación del Proceso</h2>
            <div className="grid grid-cols-2 gap-4">
              <div><label className={labelClass}>PEP / Folio Remedy</label><input className={commonInputClass} value={doc.identificacion.pepRemedy || ''} onChange={e => updateDoc('identificacion', { pepRemedy: e.target.value })} /></div>
              <div><label className={labelClass}>Sistema</label><input className={commonInputClass} value={doc.identificacion.sistema || ''} onChange={e => updateDoc('identificacion', { sistema: e.target.value })} /></div>
              <div className="col-span-2"><label className={labelClass}>Nombre del Proceso</label><input className={commonInputClass} value={doc.identificacion.proceso || ''} onChange={e => updateDoc('identificacion', { proceso: e.target.value })} /></div>
              <div><label className={labelClass}>Solicitado Por</label><input className={commonInputClass} value={doc.identificacion.requeridoPor || ''} onChange={e => updateDoc('identificacion', { requeridoPor: e.target.value })} /></div>
              <div><label className={labelClass}>Fecha Solicitud</label><input type="date" className={commonInputClass} value={doc.identificacion.requeridoFecha || ''} onChange={e => updateDoc('identificacion', { requeridoFecha: e.target.value })} /></div>
            </div>
          </div>
        );

      case 'changes':
        return (
          <div className="max-w-4xl mx-auto bg-white p-8 rounded-2xl border shadow-sm space-y-4">
            <h2 className="text-xl font-bold border-b pb-4">Control de Cambios</h2>
            {doc.controlCambios.map((c, i) => (
              <div key={i} className="grid grid-cols-4 gap-2 border-b pb-2">
                <input className={commonInputClass} value={c.version} placeholder="Ver" readOnly />
                <input className={commonInputClass + " col-span-2"} value={c.descripcion} placeholder="Descripción" readOnly />
                <input className={commonInputClass} value={c.fecha} placeholder="Fecha" readOnly />
              </div>
            ))}
          </div>
        );

      case 'content':
        return (
          <div className="max-w-4xl mx-auto bg-white p-8 rounded-2xl border shadow-sm space-y-6">
            <h2 className="text-xl font-bold border-b pb-4">Antecedentes y Objetivos</h2>
            <div><label className={labelClass}>Situación Actual</label><textarea className={commonInputClass + " h-32"} value={doc.antecedentes || ''} onChange={e => setDoc({...doc, antecedentes: e.target.value})} /></div>
            <div><label className={labelClass}>Objetivo General</label><textarea className={commonInputClass + " h-20"} value={doc.requerimiento.objetivo || ''} onChange={e => updateDoc('requerimiento', { objetivo: e.target.value })} /></div>
            <div><label className={labelClass}>Necesidad de Negocio</label><textarea className={commonInputClass + " h-20"} value={doc.requerimiento.necesidad || ''} onChange={e => updateDoc('requerimiento', { necesidad: e.target.value })} /></div>
          </div>
        );

      case 'steps':
        return (
          <div className="max-w-5xl mx-auto bg-white p-8 rounded-2xl border shadow-sm space-y-6">
            <div className="flex justify-between items-center border-b pb-4">
              <div>
                <h2 className="text-xl font-bold">Flujo de Proceso Detallado</h2>
                <div className="flex items-center space-x-2 mt-1">
                  <span className="text-[10px] bg-pink-100 text-pink-600 px-2 py-1 rounded font-black">MÁXIMA PRECISIÓN REQUERIDA</span>
                  <span className="text-[10px] bg-blue-100 text-blue-600 px-2 py-1 rounded font-black flex items-center">
                    <i className="fas fa-wave-square mr-1"></i> ALINEACIÓN AUDIO-VISUAL PRIORITARIA
                  </span>
                </div>
              </div>
              <button 
                onClick={handleAutoLoadCaptures}
                disabled={refreshingStep !== null}
                className="bg-indigo-600 text-white px-4 py-2 rounded-lg text-xs font-bold hover:bg-indigo-700 disabled:opacity-50 transition-all shadow-md shadow-indigo-200 flex items-center"
              >
                {refreshingStep === -1 ? (
                  <><i className="fas fa-spinner fa-spin mr-2"></i> Procesando {Math.round((analysisProgress.current / (doc.pasosProceso.length || 1)) * 100)}%...</>
                ) : (
                  <><i className="fas fa-images mr-2"></i> Auto-cargar Capturas</>
                )}
              </button>
            </div>
            {doc.pasosProceso.map((p, i) => (
              <React.Fragment key={i}>
                <div className="flex space-x-4 items-start p-4 bg-slate-50 rounded-lg relative group">
                  <button 
                    onClick={() => handleDeleteStep(i)}
                    className="absolute top-2 right-2 text-slate-300 hover:text-red-500 opacity-0 group-hover:opacity-100 transition-all p-2"
                    title="Eliminar paso"
                  >
                    <i className="fas fa-trash-can"></i>
                  </button>
                  <span className="bg-slate-900 text-white w-8 h-8 rounded flex items-center justify-center font-bold text-xs shrink-0">{p.paso}</span>
                  <div className="flex-1 grid grid-cols-2 gap-4">
                    <div className="col-span-2">
                      <div className="flex justify-between items-center mb-1">
                        <label className={labelClass}>Descripción del Paso (Literal)</label>
                        <button 
                          onClick={() => {
                            const newPasos = [...doc.pasosProceso];
                            newPasos[i].ocultarDetalles = !newPasos[i].ocultarDetalles;
                            setDoc({...doc, pasosProceso: newPasos});
                          }}
                          className={`text-[10px] font-bold flex items-center px-2 py-1 rounded transition-colors ${
                            p.ocultarDetalles 
                            ? 'bg-green-100 text-green-700 hover:bg-green-200' 
                            : 'bg-red-100 text-red-700 hover:bg-red-200'
                          }`}
                        >
                          <i className={`fas ${p.ocultarDetalles ? 'fa-eye' : 'fa-eye-slash'} mr-1`}></i>
                          {p.ocultarDetalles ? 'Mostrar Detalles (Captura, Tiempo...)' : 'Quitar Detalles (Captura, Tiempo...)'}
                        </button>
                      </div>
                      <textarea className={commonInputClass} value={p.descripcion || ''} onChange={e => {
                        const newPasos = [...doc.pasosProceso];
                        newPasos[i].descripcion = e.target.value;
                        setDoc({...doc, pasosProceso: newPasos});
                      }} />
                    </div>
                    
                    {!p.ocultarDetalles && (
                      <>
                        <div className="col-span-2 md:col-span-1">
                          <label className={labelClass}>Timestamp Video (MM:SS)</label>
                          <div className="flex space-x-2">
                            <input className={commonInputClass} value={p.timestamp || ''} onChange={e => {
                              const newPasos = [...doc.pasosProceso];
                              newPasos[i].timestamp = e.target.value;
                              setDoc({...doc, pasosProceso: newPasos});
                              setStepsConfirmed(false);
                            }} />
                            <button 
                              onClick={() => handleRefreshFrame(i)}
                              disabled={refreshingStep === i}
                              className="bg-slate-900 text-white px-3 rounded-lg text-[10px] font-bold hover:bg-slate-800 disabled:opacity-50 transition-all whitespace-nowrap"
                            >
                              {refreshingStep === i ? <i className="fas fa-spinner fa-spin"></i> : 'Ver Captura'}
                            </button>
                          </div>
                        </div>
                        <div className="col-span-2 md:col-span-1">
                          <label className={labelClass}>Resultado Esperado</label>
                          <input className={commonInputClass} value={p.resultadoEsperado || ''} onChange={e => {
                            const newPasos = [...doc.pasosProceso];
                            newPasos[i].resultadoEsperado = e.target.value;
                            setDoc({...doc, pasosProceso: newPasos});
                          }} />
                        </div>
                        {p.pantallaAsociada ? (
                          <div className="col-span-2 mt-2">
                            <div className="flex justify-between items-end mb-1">
                              <label className={labelClass}>Previsualización de Captura</label>
                              <button 
                                onClick={() => {
                                  const newPasos = [...doc.pasosProceso];
                                  newPasos[i].pantallaAsociada = undefined;
                                  setDoc({...doc, pasosProceso: newPasos});
                                }}
                                className="text-[10px] text-red-500 hover:text-red-700 font-bold flex items-center"
                              >
                                <i className="fas fa-trash-can mr-1"></i> Quitar Captura
                              </button>
                            </div>
                            <div className="border rounded-lg overflow-hidden bg-black aspect-video max-w-sm">
                              <img src={p.pantallaAsociada} alt="Preview" className="w-full h-full object-contain" loading="lazy" />
                            </div>
                          </div>
                        ) : (
                          <div className="col-span-2 mt-2">
                            <button 
                              onClick={() => handleRefreshFrame(i)}
                              className="text-xs text-slate-500 hover:text-slate-900 flex items-center space-x-2 bg-slate-100 px-4 py-2 rounded-lg border border-dashed border-slate-300 w-full justify-center"
                            >
                              <i className="fas fa-image"></i>
                              <span>Cargar previsualización de captura</span>
                            </button>
                          </div>
                        )}
                      </>
                    )}
                  </div>
                </div>

                {/* Botón para añadir paso entre existentes */}
                {i < doc.pasosProceso.length - 1 && (
                  <div className="flex justify-center -my-3 relative z-10 group/add">
                    <div className="absolute inset-0 flex items-center" aria-hidden="true">
                      <div className="w-full border-t border-slate-100 group-hover/add:border-pink-200 transition-colors"></div>
                    </div>
                    <button 
                      onClick={() => handleAddStepBetween(i)}
                      className="relative w-8 h-8 bg-white border-2 border-slate-100 rounded-full flex items-center justify-center text-slate-400 hover:text-pink-600 hover:border-pink-600 hover:shadow-md transition-all"
                      title="Añadir paso aquí"
                    >
                      <i className="fas fa-plus text-xs"></i>
                    </button>
                  </div>
                )}
              </React.Fragment>
            ))}
            <div className="flex justify-between items-center pt-6 border-t">
              <button className="text-pink-600 font-bold text-sm hover:underline" onClick={() => {
                setDoc({...doc, pasosProceso: [...doc.pasosProceso, { paso: doc.pasosProceso.length + 1, descripcion: '' }]});
                setStepsConfirmed(false);
              }}>
                + Añadir Paso Manualmente
              </button>
              
              <button 
                onClick={() => {
                  setStepsConfirmed(true);
                  setActiveSection('preview');
                }}
                className={`px-8 py-3 rounded-xl font-bold shadow-lg transition-all ${
                  stepsConfirmed 
                  ? 'bg-green-600 text-white hover:bg-green-700' 
                  : 'bg-pink-600 text-white hover:bg-pink-700'
                }`}
              >
                {stepsConfirmed ? <i className="fas fa-check-circle mr-2"></i> : <i className="fas fa-save mr-2"></i>}
                Confirmar Flujo y Ver Vista Previa
              </button>
            </div>
          </div>
        );

      case 'transcription':
        return (
          <div className="max-w-4xl mx-auto bg-white p-8 rounded-2xl border shadow-sm space-y-6">
            <div className="flex justify-between items-center border-b pb-4">
              <h2 className="text-xl font-bold">Transcripción Literal del Video</h2>
              <span className="text-[10px] bg-blue-100 text-blue-600 px-2 py-1 rounded font-black">AUDIO COMPLETO</span>
            </div>
            <p className="text-xs text-slate-500 italic">Esta sección contiene la transcripción palabra por palabra de todo el audio detectado en el video cargado.</p>
            <textarea 
              className={commonInputClass + " h-[500px] font-mono leading-relaxed"} 
              value={doc.transcripcion || ''} 
              onChange={e => setDoc({...doc, transcripcion: e.target.value})}
              placeholder="La transcripción aparecerá aquí después del análisis..."
            />
          </div>
        );

      case 'rules':
        return (
          <div className="max-w-4xl mx-auto bg-white p-8 rounded-2xl border shadow-sm space-y-4">
            <h2 className="text-xl font-bold border-b pb-4">Reglas de Negocio</h2>
            {doc.reglasNegocio.map((r, i) => (
              <div key={i} className="flex space-x-2">
                <span className="font-bold text-pink-600 min-w-[50px]">RN-{i+1}</span>
                <input className={commonInputClass} value={r || ''} onChange={e => {
                  const newRules = [...doc.reglasNegocio];
                  newRules[i] = e.target.value;
                  setDoc({...doc, reglasNegocio: newRules});
                }} />
                <button onClick={() => setDoc({...doc, reglasNegocio: doc.reglasNegocio.filter((_, idx) => idx !== i)})} className="text-slate-300 hover:text-red-500"><i className="fas fa-times"></i></button>
              </div>
            ))}
            <button className="text-pink-600 font-bold text-sm" onClick={() => setDoc({...doc, reglasNegocio: [...doc.reglasNegocio, '']})}>
              + Añadir Regla
            </button>
          </div>
        );

      case 'criteria':
        return (
          <div className="max-w-4xl mx-auto bg-white p-8 rounded-2xl border shadow-sm space-y-4">
            <h2 className="text-xl font-bold border-b pb-4">Criterios de aceptación</h2>
            {doc.criteriosAceptacion.map((c, i) => (
              <div key={i} className="flex space-x-2">
                <span className="font-bold text-pink-600 min-w-[50px]">CA-{i+1}</span>
                <input className={commonInputClass} value={c || ''} onChange={e => {
                  const newCriteria = [...doc.criteriosAceptacion];
                  newCriteria[i] = e.target.value;
                  setDoc({...doc, criteriosAceptacion: newCriteria});
                }} />
                <button onClick={() => setDoc({...doc, criteriosAceptacion: doc.criteriosAceptacion.filter((_, idx) => idx !== i)})} className="text-slate-300 hover:text-red-500"><i className="fas fa-times"></i></button>
              </div>
            ))}
            <button className="text-pink-600 font-bold text-sm" onClick={() => setDoc({...doc, criteriosAceptacion: [...doc.criteriosAceptacion, '']})}>
              + Añadir Criterio
            </button>
          </div>
        );

      case 'history':
        return (
          <div className="max-w-4xl mx-auto space-y-6">
            <div className="bg-white p-8 rounded-2xl border shadow-sm">
              <div className="flex justify-between items-center mb-8">
                <div>
                  <h2 className="text-xl font-bold flex items-center text-slate-900">
                    <i className="fas fa-history mr-3 text-pink-600"></i>
                    Historial de Documentos
                  </h2>
                  <p className="text-xs text-slate-500 mt-1">Gestión de los últimos 5 procesos generados</p>
                </div>
              </div>
              
              <div className="space-y-4">
                {history.length === 0 ? (
                  <div className="text-center py-20 bg-slate-50 rounded-2xl border border-dashed border-slate-200">
                    <div className="w-16 h-16 bg-white rounded-full flex items-center justify-center mx-auto mb-4 shadow-sm">
                      <i className="fas fa-folder-open text-2xl text-slate-300"></i>
                    </div>
                    <p className="text-slate-500 font-medium">No hay documentos en el historial todavía.</p>
                    <p className="text-[10px] text-slate-400 mt-1">Genera tu primer ERF para verlo aquí.</p>
                  </div>
                ) : (
                  <div className="grid grid-cols-1 gap-4">
                    {history.map((item, idx) => (
                      <div key={item.id} className="flex items-center justify-between p-5 bg-white rounded-2xl border border-slate-100 hover:border-pink-200 hover:shadow-xl hover:shadow-pink-500/5 transition-all group cursor-pointer" onClick={() => handleOpenFromHistory(item)}>
                        <div className="flex items-center space-x-5 flex-1 min-w-0">
                          <div className="w-10 h-10 bg-slate-50 rounded-xl flex items-center justify-center font-black text-slate-400 group-hover:bg-pink-50 group-hover:text-pink-600 transition-colors">
                            {idx + 1}
                          </div>
                          <div className="flex-1 min-w-0">
                            <h3 className="font-bold text-slate-900 group-hover:text-pink-600 transition-colors truncate">{item.title}</h3>
                            <p className="text-xs text-slate-500 mt-1 line-clamp-1 font-medium">{item.summary}</p>
                            <div className="flex items-center space-x-3 mt-2">
                              <span className="text-[9px] font-black text-slate-400 uppercase tracking-tighter bg-slate-100 px-2 py-0.5 rounded">
                                {new Date(item.timestamp).toLocaleDateString()}
                              </span>
                              <span className="text-[9px] font-black text-slate-400 uppercase tracking-tighter bg-slate-100 px-2 py-0.5 rounded">
                                {new Date(item.timestamp).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                              </span>
                            </div>
                          </div>
                        </div>
                        <div className="flex items-center space-x-4">
                          <button 
                            onClick={(e) => handleDeleteHistory(e, item.id)}
                            className="w-8 h-8 bg-slate-100 text-slate-400 rounded-full flex items-center justify-center hover:bg-red-50 hover:text-red-600 transition-all z-10 relative"
                            title="Eliminar"
                          >
                            <i className="fas fa-trash-alt text-xs"></i>
                          </button>
                          <div className="opacity-0 group-hover:opacity-100 transition-all transform translate-x-2 group-hover:translate-x-0">
                            <div className="w-8 h-8 bg-pink-600 text-white rounded-full flex items-center justify-center shadow-lg shadow-pink-200">
                              <i className="fas fa-arrow-right text-xs"></i>
                            </div>
                          </div>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          </div>
        );

      case 'annexes':
        return (
          <div className="max-w-4xl mx-auto space-y-8">
            <div className="bg-white p-8 rounded-2xl border shadow-sm space-y-4">
              <h2 className="text-xl font-bold border-b pb-4">Anexos Visuales</h2>
              <p className="text-xs text-slate-500 italic">Los archivos de soporte detectados aparecerán como marcadores de posición. Puedes subir imágenes o PDFs para incluirlos visualmente.</p>
              
              <div className="mb-6">
                <label className="flex flex-col items-center justify-center w-full h-24 border-2 border-dashed rounded-xl cursor-pointer hover:bg-slate-50 transition-all group">
                  <i className="fas fa-paperclip text-xl text-pink-600 mb-1 group-hover:scale-110 transition-transform"></i>
                  <span className="text-[10px] font-bold text-slate-500 uppercase">Subir Anexo Visual (PDF, JPG, PNG)</span>
                  <input type="file" className="hidden" multiple accept="image/*,application/pdf" onChange={handleAnnexUpload} />
                </label>
                {isProcessingAnnex && (
                  <div className="mt-2 text-center text-xs text-pink-600 font-bold animate-pulse">
                    <i className="fas fa-spinner fa-spin mr-2"></i> Procesando archivos...
                  </div>
                )}
              </div>

              <div className="space-y-4">
                {doc.anexos.map((a, i) => (
                  <div key={i} className="bg-slate-50 p-4 rounded-xl border border-slate-200 group relative">
                    <button 
                      onClick={() => setDoc({...doc, anexos: doc.anexos.filter((_, idx) => idx !== i)})} 
                      className="absolute top-2 right-2 text-slate-300 hover:text-red-500 transition-colors"
                    >
                      <i className="fas fa-trash-can"></i>
                    </button>
                    <div className="flex items-center space-x-2 mb-2">
                      <i className="fas fa-file-alt text-pink-600 text-xs"></i>
                      <input 
                        className="bg-transparent text-xs font-bold outline-none flex-1" 
                        value={a.nombre || ''} 
                        onChange={e => {
                          const newAnnexes = [...doc.anexos];
                          newAnnexes[i].nombre = e.target.value;
                          setDoc({...doc, anexos: newAnnexes});
                        }} 
                      />
                    </div>
                    {a.contenido && (
                      <div className="mt-2 border rounded-lg overflow-hidden bg-white max-h-48">
                        <img src={a.contenido} alt={a.nombre} className="w-full h-full object-contain" />
                      </div>
                    )}
                  </div>
                ))}
              </div>
              
              <button className="text-pink-600 font-bold text-sm mt-4" onClick={() => setDoc({...doc, anexos: [...doc.anexos, { nombre: 'Nuevo Anexo' }]})}>
                + Añadir Marcador Manual
              </button>
            </div>

            <div className="bg-white p-8 rounded-2xl border shadow-sm space-y-4">
              <h2 className="text-xl font-bold border-b pb-4">Hipervínculos y Enlaces Externos</h2>
              <p className="text-xs text-slate-500 italic">Agrega enlaces a Google Drive, repositorios o documentación externa relevante.</p>
              
              <div className="space-y-3">
                {doc.hipervinculos.map((h, i) => (
                  <div key={i} className="flex items-center space-x-3 bg-slate-50 p-3 rounded-xl border group">
                    <div className="flex-1 grid grid-cols-2 gap-3">
                      <div>
                        <label className="text-[9px] font-bold text-slate-400 uppercase block mb-1">Etiqueta / Nombre</label>
                        <input 
                          className="w-full p-2 bg-white border rounded text-xs outline-none focus:ring-1 focus:ring-pink-500"
                          placeholder="Ej: Carpeta Google Drive"
                          value={h.etiqueta || ''}
                          onChange={e => {
                            const newLinks = [...doc.hipervinculos];
                            newLinks[i].etiqueta = e.target.value;
                            setDoc({...doc, hipervinculos: newLinks});
                          }}
                        />
                      </div>
                      <div>
                        <label className="text-[9px] font-bold text-slate-400 uppercase block mb-1">URL / Enlace</label>
                        <input 
                          className="w-full p-2 bg-white border rounded text-xs outline-none focus:ring-1 focus:ring-pink-500 font-mono"
                          placeholder="https://drive.google.com/..."
                          value={h.url || ''}
                          onChange={e => {
                            const newLinks = [...doc.hipervinculos];
                            newLinks[i].url = e.target.value;
                            setDoc({...doc, hipervinculos: newLinks});
                          }}
                        />
                      </div>
                    </div>
                    <button 
                      onClick={() => setDoc({...doc, hipervinculos: doc.hipervinculos.filter((_, idx) => idx !== i)})}
                      className="text-slate-300 hover:text-red-500 p-2 transition-colors"
                    >
                      <i className="fas fa-trash-can text-sm"></i>
                    </button>
                  </div>
                ))}
              </div>

              <button 
                className="flex items-center text-pink-600 font-bold text-sm mt-4 hover:text-pink-700 transition-colors"
                onClick={() => setDoc({...doc, hipervinculos: [...doc.hipervinculos, { etiqueta: '', url: '' }]})}
              >
                <i className="fas fa-link mr-2"></i> Añadir Hipervínculo
              </button>
            </div>
          </div>
        );

      case 'preview':
        return (
          <div className="space-y-6 pb-20">
            {!stepsConfirmed && (
              <div className="max-w-4xl mx-auto bg-amber-50 border border-amber-200 p-4 rounded-xl flex items-center space-x-3 text-amber-800 mb-6">
                <i className="fas fa-exclamation-triangle text-xl"></i>
                <div className="text-sm">
                  <p className="font-bold">Flujo no confirmado</p>
                  <p>Debes confirmar los pasos en la sección <b>Flujo de Proceso</b> antes de poder exportar el documento.</p>
                </div>
                <button 
                  onClick={() => setActiveSection('steps')}
                  className="ml-auto bg-amber-600 text-white px-4 py-1.5 rounded-lg font-bold text-xs hover:bg-amber-700 transition-all"
                >
                  Ir a Flujo
                </button>
              </div>
            )}
            <div className="no-print flex justify-end space-x-4 mb-4 pr-4">
              <button 
                onClick={handleDownloadPDF} 
                disabled={isExportingPDF || !stepsConfirmed} 
                className="bg-red-600 text-white px-6 py-2.5 rounded-xl font-bold shadow-lg hover:bg-red-700 disabled:opacity-50 disabled:grayscale transition-all"
              >
                {isExportingPDF ? <i className="fas fa-spinner fa-spin mr-2"></i> : <i className="fas fa-file-pdf mr-2"></i>} PDF (A4)
              </button>
              <button 
                onClick={handleDownloadWord} 
                disabled={isExportingPDF || !stepsConfirmed}
                className="bg-blue-600 text-white px-6 py-2.5 rounded-xl font-bold shadow-lg hover:bg-blue-700 disabled:opacity-50 disabled:grayscale transition-all"
              >
                {isExportingPDF ? <i className="fas fa-spinner fa-spin mr-2"></i> : <i className="fas fa-file-word mr-2"></i>} Word
              </button>
            </div>
            
            <div ref={previewRef} className="pdf-page mx-auto space-y-10 relative">
              {/* 1. PORTADA */}
              <div className="min-h-[260mm] flex flex-col relative">
                <div className="watermark">Eficiencia de Procesos C.C.</div>
                <div className="flex justify-between items-center border-b-2 border-black pb-6">
                  <div className="flex items-center space-x-2">
                    <div className="w-8 h-8 bg-pink-600 text-white flex items-center justify-center font-bold text-lg">L</div>
                    <h2 className="text-3xl font-light text-purple-700">Liverpool<sup>®</sup></h2>
                  </div>
                  <div className="bg-orange-400 p-1.5 text-white font-bold text-2xl italic">MTi</div>
                </div>

                <div className="text-center my-24 space-y-6 flex-1 flex flex-col justify-center">
                  <h3 className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-400">Metodología de Trabajo de Informática</h3>
                  <h1 className="text-4xl font-black uppercase text-slate-900 leading-tight">Especificación de Requerimiento Funcional (ERF)</h1>
                  <div className="bg-slate-50 py-16 rounded-xl border border-slate-200">
                    <h2 className="text-3xl font-black text-slate-900 px-10">{(doc.identificacion.proceso || 'PROCESO').toUpperCase()}</h2>
                    <div className="mt-8 flex justify-center space-x-10 text-slate-500 font-bold uppercase text-[9px] tracking-widest">
                      <span>Versión: {doc.metadata.version}</span>
                      <span>Fecha: {doc.metadata.fecha}</span>
                      <span>Código: {doc.metadata.codigo}</span>
                    </div>
                  </div>
                </div>

                <div className="mt-auto border-t pt-4 text-[9px] text-slate-600 leading-normal space-y-2">
                  <p className="font-bold">Documento Confidencial.</p>
                  <p>Derechos Reservados. El contenido y formato de este documento son propiedad de Servicios Liverpool, S. A. de C. V. y es confidencial. Cualquier reproducción del todo o en parte está estrictamente prohibida sin el permiso escrito de Servicios Liverpool, S. A. de C. V.</p>
                  <p>Este documento está sujeto a cambios. Comentarios, correcciones o preguntas deberán ser dirigidas al autor.</p>
                </div>
              </div>

              {/* 2. ÍNDICE */}
              <div className="page-break"></div>
              <section className="py-10 min-h-[260mm] relative">
                <h3 className="text-lg font-black border-b-2 border-black mb-8 pb-3 uppercase tracking-widest">ÍNDICE DE CONTENIDO</h3>
                <div className="space-y-4">
                  <div className="toc-link"><span>1. IDENTIFICACIÓN</span><span>{'{PAGEREF}'}</span></div>
                  <div className="toc-link"><span>2. CONTROL DE CAMBIOS</span><span>{'{PAGEREF}'}</span></div>
                  <div className="toc-link"><span>3. ANTECEDENTES Y OBJETIVOS</span><span>{'{PAGEREF}'}</span></div>
                  <div className="toc-link"><span>4. FLUJO DE PROCESO</span><span>{'{PAGEREF}'}</span></div>
                  <div className="toc-link"><span>5. REGLAS DE NEGOCIO</span><span>{'{PAGEREF}'}</span></div>
                  <div className="toc-link"><span>6. CRITERIOS DE ACEPTACIÓN</span><span>{'{PAGEREF}'}</span></div>
                  <div className="toc-link"><span>7. ANEXOS</span><span>{'{PAGEREF}'}</span></div>
                </div>
              </section>

              {/* 3. IDENTIFICACIÓN */}
              <div className="page-break"></div>
              <section className="py-10 relative">
                <h3 className="text-sm font-black bg-black text-white p-3 mb-6 uppercase tracking-widest">1. IDENTIFICACIÓN</h3>
                <table className="report-table mb-8">
                  <tbody>
                    <tr><th className="w-1/3">PEP / FOLIO REMEDY</th><td>{doc.identificacion.pepRemedy || ''}</td></tr>
                    <tr><th>SISTEMA / PLATAFORMA</th><td>{doc.identificacion.sistema || ''}</td></tr>
                    <tr><th>MÓDULO</th><td>{doc.identificacion.modulo || ''}</td></tr>
                    <tr><th>PROCESO</th><td>{doc.identificacion.proceso || ''}</td></tr>
                    <tr><th>SOLICITANTE</th><td>{doc.identificacion.requeridoPor || ''}</td></tr>
                    <tr><th>FECHA SOLICITUD</th><td>{doc.identificacion.requeridoFecha || ''}</td></tr>
                  </tbody>
                </table>
              </section>

              {/* 4. CONTROL DE CAMBIOS */}
              <div className="page-break"></div>
              <section className="py-10 relative">
                <h3 className="text-sm font-black bg-black text-white p-3 mb-6 uppercase tracking-widest">2. CONTROL DE CAMBIOS</h3>
                <h4 className="text-xs font-bold mb-4 uppercase">Historial de Revisiones</h4>
                <table className="report-table">
                  <thead><tr><th>VER.</th><th>DESCRIPCIÓN</th><th>AUTOR</th><th>FECHA</th></tr></thead>
                  <tbody>{doc.controlCambios.map((c, i) => <tr key={i}><td className="text-center">{c.version}</td><td>{c.descripcion}</td><td>{c.autor}</td><td className="text-center">{c.fecha}</td></tr>)}</tbody>
                </table>
              </section>

              {/* 5. ANTECEDENTES Y OBJETIVOS */}
              <div className="page-break"></div>
              <section className="py-10 relative">
                <h3 className="text-sm font-black bg-black text-white p-3 mb-6 uppercase tracking-widest">3. ANTECEDENTES Y OBJETIVOS</h3>
                <div className="mb-10">
                  <h4 className="text-[10px] font-black text-pink-600 uppercase mb-2">Antecedentes y Justificación</h4>
                  <p className="text-[11px] leading-relaxed whitespace-pre-wrap p-4 bg-slate-50 border rounded">{doc.antecedentes || ''}</p>
                </div>
                <div className="space-y-4">
                  <div><h4 className="text-[10px] font-black text-pink-600 uppercase">Objetivo General</h4><p className="text-[11px] border-l-2 border-pink-200 pl-4 py-1">{doc.requerimiento.objetivo || ''}</p></div>
                  <div><h4 className="text-[10px] font-black text-pink-600 uppercase">Necesidad de Negocio</h4><p className="text-[11px] border-l-2 border-pink-200 pl-4 py-1">{doc.requerimiento.necesidad || ''}</p></div>
                </div>
              </section>

              {/* 6. FLUJO DE PROCESO */}
              <div className="page-break"></div>
              <section className="py-10 relative">
                <h3 className="text-sm font-black bg-black text-white p-3 mb-6 uppercase tracking-widest">4. FLUJO DE PROCESO</h3>
                <div className="space-y-6">
                  {doc.pasosProceso.map((p, i) => (
                    <div key={i} className="border rounded-lg overflow-hidden">
                      <div className="bg-slate-900 text-white p-2 flex justify-between items-center text-[10px] font-bold">
                        <span>PASO {p.paso}</span>
                        {!p.ocultarDetalles && p.timestamp && <span>REF: {p.timestamp}</span>}
                      </div>
                      <div className="p-4 space-y-2">
                        <p className="text-[11px] font-medium">{p.descripcion}</p>
                        {!p.ocultarDetalles && (
                          <>
                            {p.resultadoEsperado && <div className="text-[10px] bg-green-50 p-2 border border-green-100 rounded"><b>Resultado Esperado:</b> {p.resultadoEsperado}</div>}
                          </>
                        )}
                        {p.pantallaAsociada && (
                          <div className="mt-4 border rounded overflow-hidden bg-slate-100">
                            <img src={p.pantallaAsociada} alt={`Paso ${p.paso}`} className="w-full h-auto max-h-[300px] object-contain" />
                            <div className="bg-slate-200 text-[9px] px-2 py-1 text-slate-600 font-mono">
                              [CAPTURA_DE_PANTALLA: {p.timestamp}]
                            </div>
                          </div>
                        )}
                      </div>
                    </div>
                  ))}
                </div>
              </section>

              {/* 7. REGLAS DE NEGOCIO */}
              <div className="page-break"></div>
              <section className="py-10 relative">
                <h3 className="text-sm font-black bg-black text-white p-3 mb-6 uppercase tracking-widest">5. REGLAS DE NEGOCIO</h3>
                <div className="space-y-2 mb-10">
                  {doc.reglasNegocio.length > 0 ? doc.reglasNegocio.map((rn, i) => <div key={i} className="flex space-x-2 text-[11px]"><span className="font-bold text-pink-600 min-w-[50px]">RN-{i+1}</span><p>{rn}</p></div>) : null}
                </div>
              </section>

              {/* 8. CRITERIOS DE ACEPTACIÓN */}
              <div className="page-break"></div>
              <section className="py-10 relative">
                <h3 className="text-sm font-black bg-black text-white p-3 mb-6 uppercase tracking-widest">6. CRITERIOS DE ACEPTACIÓN</h3>
                <div className="space-y-2">
                  {doc.criteriosAceptacion.length > 0 ? doc.criteriosAceptacion.map((ca, i) => <div key={i} className="flex space-x-2 text-[11px]"><span className="font-bold text-pink-600 min-w-[50px]">CA-{i+1}</span><p>{ca}</p></div>) : null}
                </div>
              </section>

              {/* 9. ANEXOS */}
              <div className="page-break"></div>
              <section className="py-10 relative">
                <h3 className="text-sm font-black bg-black text-white p-3 mb-6 uppercase tracking-widest">7. ANEXOS</h3>
                
                {doc.hipervinculos.length > 0 && (
                  <div className="mb-10">
                    <h4 className="text-[10px] font-black text-pink-600 uppercase mb-3">Hipervínculos y Enlaces Externos</h4>
                    <div className="grid grid-cols-1 gap-2">
                      {doc.hipervinculos.map((h, i) => (
                        <div key={i} className="flex items-center space-x-2 text-[11px] p-2 border rounded bg-slate-50">
                          <i className="fas fa-link text-pink-500"></i>
                          <span className="font-bold">{h.etiqueta}:</span>
                          <a href={h.url} target="_blank" rel="noopener noreferrer" className="text-blue-600 underline truncate">{h.url}</a>
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                <div className="space-y-6">
                  {doc.anexos.length > 0 ? doc.anexos.map((a, i) => (
                    <div key={i} className="flex flex-col space-y-2 p-4 border rounded-xl bg-slate-50 text-[11px]">
                      <div className="flex items-center space-x-2 font-bold text-pink-600">
                        <i className="fas fa-paperclip"></i>
                        <span>ANEXO {i+1}: {a.nombre}</span>
                      </div>
                      {a.contenido ? (
                        <div className="mt-2 border rounded-lg overflow-hidden bg-white">
                          <img src={a.contenido} alt={a.nombre} className="w-full h-auto" />
                        </div>
                      ) : (
                        <p className="font-mono text-[10px] bg-white p-2 border border-slate-200 rounded">
                          [INSERTAR DOCUMENTO COMPLETO: {a.nombre}]
                        </p>
                      )}
                    </div>
                  )) : null}
                </div>
              </section>
            </div>
          </div>
        );

      default:
        return null;
    }
  };

  return (
    <>
      <Layout activeSection={activeSection} onSectionChange={setActiveSection}>
        {renderSection()}
        {activeSection !== 'input' && activeSection !== 'history' && renderHistoryWindow()}
      </Layout>

      {/* Delete Confirmation Modal */}
      {deleteConfirmation.isOpen && (
        <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-2xl max-w-md w-full p-6 animate-in fade-in zoom-in duration-200">
            <div className="flex items-center space-x-4 mb-4">
              <div className="w-12 h-12 bg-red-100 rounded-full flex items-center justify-center text-red-600">
                <i className="fas fa-exclamation-triangle text-xl"></i>
              </div>
              <div>
                <h3 className="text-lg font-bold text-slate-900">Confirmar eliminación</h3>
                <p className="text-sm text-slate-500">Esta acción no se puede deshacer.</p>
              </div>
            </div>
            <p className="text-slate-600 text-sm mb-6">
              ¿Estás seguro de que deseas eliminar este documento del historial? Se borrará permanentemente del almacenamiento local.
            </p>
            <div className="flex justify-end space-x-3">
              <button
                onClick={cancelDeleteHistory}
                className="px-4 py-2 rounded-xl text-sm font-bold text-slate-600 hover:bg-slate-100 transition-colors"
              >
                Cancelar
              </button>
              <button
                onClick={confirmDeleteHistory}
                className="px-4 py-2 rounded-xl text-sm font-bold text-white bg-red-600 hover:bg-red-700 shadow-lg shadow-red-200 transition-colors"
              >
                Eliminar
              </button>
            </div>
          </div>
        </div>
      )}

      {showHistoryPrompt && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4 animate-in fade-in duration-300">
          <div className="bg-white rounded-2xl shadow-2xl max-w-md w-full overflow-hidden animate-in zoom-in-95 duration-300">
            <div className="bg-pink-600 p-6 text-white">
              <h3 className="text-xl font-bold flex items-center">
                <i className="fas fa-history mr-3"></i>
                Recuperar Historial
              </h3>
              <p className="text-pink-100 text-sm mt-2">Hemos detectado documentos guardados de sesiones anteriores.</p>
            </div>
            <div className="p-6 space-y-4">
              <p className="text-sm text-slate-600">¿Deseas ver el historial de los últimos 5 documentos generados o empezar uno nuevo?</p>
              <div className="flex flex-col space-y-2">
                <button 
                  onClick={() => {
                    setActiveSection('history');
                    setShowHistoryPrompt(false);
                  }}
                  className="w-full py-3 bg-pink-600 text-white rounded-xl font-bold hover:bg-pink-700 transition-all shadow-lg shadow-pink-100"
                >
                  Ver Historial Guardado
                </button>
                <button 
                  onClick={() => {
                    setDoc(createNewERF());
                    setShowHistoryPrompt(false);
                  }}
                  className="w-full py-3 bg-slate-100 text-slate-600 rounded-xl font-bold hover:bg-slate-200 transition-all"
                >
                  Empezar Nuevo Documento
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </>
  );
};

export default App;
