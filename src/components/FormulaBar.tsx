import React, { useState, useEffect } from 'react';
import { SelectedCellInfo } from '../types/excel';
import { Check, Plus } from 'lucide-react';

interface FormulaBarProps {
  selectedCell: SelectedCellInfo | null;
  onApplyFormulaOrValue: (input: string) => void;
  onOpenFormulaBuilder: () => void;
}

const COMMON_FORMULAS = [
  { label: 'مجموع الفروقات المالية', formula: '=SUM(I:I)' },
  { label: 'متوسط الأسعار', formula: '=AVERAGE(E:E)' },
  { label: 'أعلى سعر صنف', formula: '=MAX(E:E)' },
  { label: 'إجمالي العجز المالي', formula: '=SUMIF(J:J, "عجز", I:I)' },
  { label: 'إجمالي الزيادة المالية', formula: '=SUMIF(J:J, "زيادة", I:I)' },
  { label: 'عدد أصناف العجز', formula: '=COUNTIF(J:J, "عجز")' },
];

export const FormulaBar: React.FC<FormulaBarProps> = ({
  selectedCell,
  onApplyFormulaOrValue,
  onOpenFormulaBuilder,
}) => {
  const [fxInput, setFxInput] = useState('=SUM(I:I)');

  useEffect(() => {
    if (selectedCell) {
      setFxInput(selectedCell.formula || String(selectedCell.rawValue));
    }
  }, [selectedCell]);

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!fxInput.trim()) return;
    onApplyFormulaOrValue(fxInput.trim());
  };

  return (
    <div className="bg-white border border-slate-200 rounded-xl px-4 py-2.5 flex flex-col md:flex-row md:items-center gap-3">
      {/* Cell coordinate indicator & fx symbol */}
      <div className="flex items-center gap-2 shrink-0">
        <div
          aria-label="إحداثيات الخلية المحددة"
          className="px-3 py-1 bg-slate-100 border border-slate-200 rounded text-xs font-mono font-semibold text-slate-800 tabular-nums min-w-[64px] text-center"
          dir="ltr"
        >
          {selectedCell ? `${selectedCell.colLetter}${selectedCell.rowIndex}` : 'FX-ALL'}
        </div>
        <span className="text-xs text-slate-500 whitespace-nowrap">
          {selectedCell ? selectedCell.colLabel : 'شريط المعادلات الذكي'}
        </span>
        <span
          className="px-2 py-0.5 text-xs font-mono font-bold italic text-teal-800 bg-teal-50 rounded select-none"
          dir="ltr"
        >
          fx
        </span>
      </div>

      {/* Formula / Cell Value Input */}
      <form onSubmit={handleSubmit} className="flex-1 flex items-center gap-2">
        <label htmlFor="excel-fx-input" className="sr-only">
          شريط معادلات الإكسيل وقيم الخلايا
        </label>
        <input
          id="excel-fx-input"
          type="text"
          dir="ltr"
          value={fxInput}
          onChange={(e) => setFxInput(e.target.value)}
          placeholder="=SUM(I:I) or =AVERAGE(E:E) or =G1-F1"
          className="w-full px-3 py-1.5 text-xs font-mono bg-slate-50 border border-slate-200 rounded-md text-slate-900 focus:outline-none focus:bg-white focus:border-teal-700"
        />
        <button
          type="submit"
          className="inline-flex items-center gap-1 px-3 py-1.5 text-xs font-semibold text-white bg-teal-700 rounded-md hover:bg-teal-800 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
        >
          <Check className="w-3.5 h-3.5" />
          <span>حساب</span>
        </button>
      </form>

      {/* Quick Formula Insert & Custom Column Button */}
      <div className="flex items-center gap-1.5 overflow-x-auto pb-1 md:pb-0 shrink-0">
        {COMMON_FORMULAS.slice(0, 3).map((f) => (
          <button
            key={f.formula}
            type="button"
            onClick={() => {
              setFxInput(f.formula);
              onApplyFormulaOrValue(f.formula);
            }}
            className="px-2.5 py-1 text-xs font-medium text-slate-600 bg-slate-100 hover:bg-slate-200 hover:text-slate-900 rounded transition-colors whitespace-nowrap shrink-0 cursor-pointer"
          >
            {f.label}
          </button>
        ))}
        <button
          type="button"
          onClick={onOpenFormulaBuilder}
          className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-semibold text-teal-800 bg-teal-50 hover:bg-teal-100 rounded transition-colors whitespace-nowrap shrink-0 cursor-pointer"
        >
          <Plus className="w-3.5 h-3.5" />
          <span>عمود معادلة جديد</span>
        </button>
      </div>
    </div>
  );
};
