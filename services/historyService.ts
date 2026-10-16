
import { ERFDocument } from '../types';

export interface HistoryItem {
  id: string;
  title: string;
  summary: string;
  fullDoc: ERFDocument;
  timestamp: number;
}

const HISTORY_KEY = 'erf_document_history';
const MAX_HISTORY = 5;

export const saveToHistory = (doc: ERFDocument) => {
  const history = getHistory();
  
  // Create a summary (first 100 chars of antecedentes or objective)
  const summary = doc.antecedentes?.substring(0, 100) || doc.requerimiento.objetivo?.substring(0, 100) || 'Sin resumen disponible';
  
  const newItem: HistoryItem = {
    id: doc.id,
    title: doc.identificacion.proceso || 'Documento sin título',
    summary: summary + (summary.length >= 100 ? '...' : ''),
    fullDoc: doc,
    timestamp: Date.now()
  };

  // Remove if already exists (update)
  const filteredHistory = history.filter(item => item.id !== doc.id);
  
  // Add to end
  const newHistory = [...filteredHistory, newItem];
  
  // Keep only last 5 (FIFO: remove oldest from beginning)
  if (newHistory.length > MAX_HISTORY) {
    newHistory.shift();
  }

  try {
    localStorage.setItem(HISTORY_KEY, JSON.stringify(newHistory));
  } catch (e) {
    console.error('Error saving history', e);
  }
  return newHistory;
};

export const getHistory = (): HistoryItem[] => {
  try {
    const stored = localStorage.getItem(HISTORY_KEY);
    if (!stored) return [];
    return JSON.parse(stored);
  } catch (e) {
    console.error('Error accessing or parsing history', e);
    return [];
  }
};

export const removeFromHistory = (id: string) => {
  try {
    if (id === 'all') {
      localStorage.removeItem(HISTORY_KEY);
      return [];
    }
    const history = getHistory();
    const newHistory = history.filter(item => item.id !== id);
    localStorage.setItem(HISTORY_KEY, JSON.stringify(newHistory));
    return newHistory;
  } catch (e) {
    console.error('Error removing history', e);
    return getHistory();
  }
};

export const getHistoryItem = (idOrTitle: string): HistoryItem | undefined => {
  const history = getHistory();
  return history.find(item => item.id === idOrTitle || item.title.toLowerCase() === idOrTitle.toLowerCase());
};
