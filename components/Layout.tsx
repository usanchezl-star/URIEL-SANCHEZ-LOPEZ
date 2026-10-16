import React from 'react';

interface LayoutProps {
  children: React.ReactNode;
  activeSection: string;
  onSectionChange: (section: string) => void;
}

const Layout: React.FC<LayoutProps> = ({ children, activeSection, onSectionChange }) => {
  const menuItems = [
    { id: 'input', icon: 'fa-file-import', label: 'Carga de Datos' },
    { id: 'header', icon: 'fa-id-card', label: 'Identificación' },
    { id: 'changes', icon: 'fa-history', label: 'Control de Cambios' },
    { id: 'content', icon: 'fa-align-left', label: 'Antecedentes & Obj' },
    { id: 'steps', icon: 'fa-list-check', label: 'Flujo de Proceso' },
    { id: 'transcription', icon: 'fa-quote-left', label: 'Transcripción' },
    { id: 'rules', icon: 'fa-gavel', label: 'Reglas de Negocio' },
    { id: 'criteria', icon: 'fa-check-double', label: 'Criterios de aceptación' },
    { id: 'annexes', icon: 'fa-paperclip', label: 'Anexos' },
    { id: 'history', icon: 'fa-history', label: 'Historial' },
    { id: 'preview', icon: 'fa-eye', label: 'Vista Previa (Word)' },
  ];

  return (
    <div className="flex h-screen overflow-hidden bg-slate-50">
      <aside className="w-64 bg-slate-900 text-white flex-shrink-0 flex flex-col no-print no-scrollbar overflow-y-auto">
        <div className="p-6 flex items-center space-x-3 border-b border-slate-800">
          <div className="w-8 h-8 bg-pink-600 rounded flex items-center justify-center font-bold">L</div>
          <span className="font-bold text-lg tracking-tight">DocuGen ERF</span>
        </div>
        
        <nav className="flex-1 mt-6 px-4 space-y-1">
          {menuItems.map((item) => (
            <button
              key={item.id}
              onClick={() => onSectionChange(item.id)}
              className={`w-full flex items-center space-x-3 px-4 py-3 rounded-lg transition-all duration-200 ${
                activeSection === item.id 
                ? 'bg-pink-600 text-white shadow-lg' 
                : 'text-slate-400 hover:bg-slate-800 hover:text-white'
              }`}
            >
              <i className={`fas ${item.icon} w-5`}></i>
              <span className="font-medium">{item.label}</span>
            </button>
          ))}
        </nav>

        <div className="p-6 border-t border-slate-800 text-xs text-slate-500 mt-auto">
          <p>© 2024 Servicios Liverpool</p>
        </div>
      </aside>

      <main className="flex-1 flex flex-col min-w-0 overflow-hidden">
        <header className="h-16 bg-white border-b flex items-center justify-between px-8 flex-shrink-0 no-print">
          <h1 className="text-xl font-semibold text-slate-800 capitalize">
            {menuItems.find(i => i.id === activeSection)?.label}
          </h1>
          <div className="flex items-center space-x-4">
            <span className="text-sm text-slate-500">
              <i className="fas fa-circle text-green-500 mr-2 text-[10px]"></i>
              Gemini AI Online
            </span>
          </div>
        </header>

        <div className="flex-1 overflow-y-auto p-8 no-scrollbar">
            {children}
        </div>
      </main>
    </div>
  );
};

export default Layout;