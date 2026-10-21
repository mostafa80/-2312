import React from 'react';
import { Download } from 'lucide-react';

export type WorkspaceTab = 'grid' | 'companies' | 'formulas' | 'purged';

interface TopBarProps {
  activeTab: WorkspaceTab;
  onSelectTab: (tab: WorkspaceTab) => void;
  onExportExcel: () => void;
  purgedCount: number;
}

export const TopBar: React.FC<TopBarProps> = ({
  activeTab,
  onSelectTab,
  onExportExcel,
  purgedCount,
}) => {
  return (
    <header className="flex items-center justify-between gap-4 sm:gap-8 px-4 sm:px-6 py-3.5 bg-white border-b border-slate-200 sticky top-0 z-30">
      {/* Zone 1: Single text element wordmark */}
      <a
        href="#top"
        onClick={(e) => {
          e.preventDefault();
          onSelectTab('grid');
        }}
        className="text-base sm:text-lg font-bold tracking-tight text-slate-900 whitespace-nowrap shrink-0 font-display"
      >
        ميزان إكسيل الذكي
      </a>

      {/* Zone 2: 4 concise single-line text navigation links */}
      <nav aria-label="التنقل الرئيسي" className="hidden md:flex items-center gap-6 text-sm font-medium text-slate-600">
        <button
          type="button"
          onClick={() => onSelectTab('grid')}
          className={`py-1 transition-colors whitespace-nowrap shrink-0 border-b-2 cursor-pointer ${
            activeTab === 'grid'
              ? 'text-teal-700 border-teal-700 font-semibold'
              : 'border-transparent hover:text-slate-900'
          }`}
        >
          الجدول الموحد
        </button>
        <button
          type="button"
          onClick={() => onSelectTab('companies')}
          className={`py-1 transition-colors whitespace-nowrap shrink-0 border-b-2 cursor-pointer ${
            activeTab === 'companies'
              ? 'text-teal-700 border-teal-700 font-semibold'
              : 'border-transparent hover:text-slate-900'
          }`}
        >
          تقسيم الشركات
        </button>
        <button
          type="button"
          onClick={() => onSelectTab('formulas')}
          className={`py-1 transition-colors whitespace-nowrap shrink-0 border-b-2 cursor-pointer ${
            activeTab === 'formulas'
              ? 'text-teal-700 border-teal-700 font-semibold'
              : 'border-transparent hover:text-slate-900'
          }`}
        >
          بناء المعادلات
        </button>
        <button
          type="button"
          onClick={() => onSelectTab('purged')}
          className={`py-1 transition-colors whitespace-nowrap shrink-0 border-b-2 cursor-pointer ${
            activeTab === 'purged'
              ? 'text-teal-700 border-teal-700 font-semibold'
              : 'border-transparent hover:text-slate-900'
          }`}
        >
          المستبعدات ({purgedCount})
        </button>
      </nav>

      {/* Zone 3: 1 primary action */}
      <div className="flex items-center shrink-0">
        <button
          type="button"
          onClick={onExportExcel}
          className="inline-flex items-center gap-1.5 px-3 sm:px-4 py-2 text-xs font-semibold text-white bg-teal-700 rounded-lg hover:bg-teal-800 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
        >
          <Download className="w-3.5 h-3.5" />
          <span>تصدير Excel (.xlsx)</span>
        </button>
      </div>
    </header>
  );
};
