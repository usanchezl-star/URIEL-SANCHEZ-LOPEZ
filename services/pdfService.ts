
import html2pdf from 'html2pdf.js';

export const exportToPDF = async (element: HTMLElement, filename: string) => {
  // Intentar obtener la función constructora de varias fuentes posibles (default export, module export o global)
  // html2pdf.js es propenso a este error si no se detecta el wrapper ESM correctamente
  const h2p: any = (html2pdf as any).default || html2pdf || (window as any).html2pdf;

  if (typeof h2p !== 'function') {
    console.error("html2pdf loading error: expected a function but got", typeof h2p);
    throw new Error("La librería de PDF no se cargó correctamente. Por favor intenta de nuevo.");
  }

  const opt = {
    margin: [0, 0, 0, 0], // Sin márgenes para manejar el padding desde el CSS de la página
    filename: filename,
    image: { type: 'jpeg', quality: 1 },
    html2canvas: { 
      scale: 2, 
      useCORS: true, 
      logging: false,
      letterRendering: true,
      windowWidth: 794 // 210mm a 96dpi
    },
    jsPDF: { 
      unit: 'mm', 
      format: 'a4', 
      orientation: 'portrait',
      compress: true
    },
    pagebreak: { mode: ['avoid-all', 'css', 'legacy'] }
  };

  try {
    // Iniciar el proceso usando la función constructora identificada y el API de Worker
    await h2p().set(opt).from(element).save();
  } catch (error) {
    console.error("Error al generar PDF:", error);
    throw error;
  }
};
