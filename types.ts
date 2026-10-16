export interface ERFDocument {
  id: string;
  metadata: {
    version: string;
    fecha: string;
    codigo: string;
    nombre: string;
  };
  controlCambios: Array<{
    version: string;
    descripcion: string;
    autor: string;
    autoriza: string;
    fecha: string;
  }>;
  identificacion: {
    pepRemedy: string;
    controlCambiosNum: string;
    tipoRequerimiento: string;
    sistema: string;
    modulo: string;
    proceso: string;
    descripcion: string;
    requeridoPor: string;
    requeridoFecha: string;
    aprobadoPor: string;
    aprobadoFecha: string;
  };
  antecedentes: string;
  requerimiento: {
    objetivo: string;
    necesidad: string;
    navegacion: string[];
  };
  reglasNegocio: string[];
  criteriosAceptacion: string[];
  anexos: Array<{ nombre: string; contenido?: string }>;
  hipervinculos: Array<{ etiqueta: string; url: string }>;
  transcripcion: string;
  pasosProceso: Array<{
    paso: number;
    descripcion: string;
    timestamp?: string;
    resultadoEsperado?: string;
    pantallaAsociada?: string; 
    id_imagen?: string;
    ocultarDetalles?: boolean;
  }>;
  missingFields: string[];
  imagenesExtraidas?: Array<{ id: string; contenido: string }>;
}

export interface HistoryItem {
  id: string;
  title: string;
  date: string;
  summary: string;
  fullDoc: ERFDocument;
}

const generateId = () => {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return crypto.randomUUID();
  }
  return Math.random().toString(36).substring(2, 11) + Date.now().toString(36);
};

export const createNewERF = (nombre: string = 'Nuevo Proceso'): ERFDocument => ({
  id: generateId(),
  metadata: { version: '01', fecha: new Date().toLocaleDateString(), codigo: 'MTI-II-ERF', nombre },
  controlCambios: [{ version: '01', descripcion: 'Versión original', autor: '', autoriza: '', fecha: new Date().toLocaleDateString() }],
  identificacion: {
    pepRemedy: '',
    controlCambiosNum: '',
    tipoRequerimiento: '',
    sistema: '',
    modulo: '',
    proceso: nombre,
    descripcion: '',
    requeridoPor: '',
    requeridoFecha: new Date().toISOString().split('T')[0],
    aprobadoPor: '',
    aprobadoFecha: ''
  },
  antecedentes: '',
  transcripcion: '',
  requerimiento: {
    objetivo: '',
    necesidad: '',
    navegacion: []
  },
  reglasNegocio: [],
  criteriosAceptacion: [],
  anexos: [],
  hipervinculos: [],
  pasosProceso: [],
  missingFields: []
});