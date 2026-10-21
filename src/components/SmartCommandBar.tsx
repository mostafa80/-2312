import React, { useState } from 'react';
import { Play, RotateCcw, Sparkles, Trash2, ArrowUpDown, Calculator, Layers } from 'lucide-react';
import { EvaluatedFormulaBanner } from '../types/excel';

interface SmartCommandBarProps {
  onExecuteCommand: (commandText: string) => Promise<void>;
  onQuickPurgeZero: () => void;
  onQuickFullOrganize: () => void;
  onUndoLast: () => void;
  canUndo: boolean;
  isProcessing: boolean;
  lastFeedback: string | null;
  evaluatedBanner?: EvaluatedFormulaBanner;
}

const QUICK_EXAMPLES = [
  {
    label: 'حذف المطابق والفارق (0) فوراً',
    command: 'احذف الأصناف المطابقة أو التي الفارق فيها صفر',
    icon: Trash2,
  },
  {
    label: 'تنظيم شامل للشركات والحالات والأسعار',
    command: 'نظم الشيت تنظيماً شاملاً ودقيقاً حسب الشركة والحالة والسعر',
    icon: Layers,
  },
  {
    label: 'ترتيب الأصناف بالسعر من الأعلى للأقل',
    command: 'رتب الأصناف حسب السعر من الأعلى للأقل',
    icon: ArrowUpDown,
  },
  {
    label: 'إضافة عمود السعر بالضريبة 14%',
    command: 'اضف عمود جديد يحسب السعر بالضريبة 14%',
    icon: Calculator,
  },
  {
    label: 'حساب إجمالي قيمة العجز المالي',
    command: 'احسب مجموع قيمة العجز في الشيت',
    icon: Calculator,
  },
];

export const SmartCommandBar: React.FC<SmartCommandBarProps> = ({
  onExecuteCommand,
  onQuickPurgeZero,
  onQuickFullOrganize,
  onUndoLast,
  canUndo,
  isProcessing,
  lastFeedback,
  evaluatedBanner,
}) => {
  const [commandInput, setCommandInput] = useState('');
  const [validationError, setValidationError] = useState<string | null>(null);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    const trimmed = commandInput.trim();
    if (!trimmed) {
      setValidationError('يرجى كتابة الأمر أو شرح ما تريد تنفيذه على الشيت أولاً.');
      return;
    }
    setValidationError(null);
    await onExecuteCommand(trimmed);
    setCommandInput('');
  };

  return (
    <section
      aria-label="المساعد الذكي لتنفيذ الأوامر الكتابية والمعادلات"
      className="bg-white border border-slate-200 rounded-xl p-5"
    >
      <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 mb-4">
        <div>
          <h2 className="text-base font-bold text-slate-900 font-display">
            اشرح ما تريد تنفيذه كتابياً أو اطلب أي معادلة بدون تعقيد
          </h2>
          <p className="text-xs text-slate-600 mt-0.5">
            اكتب أوامرك باللغة العربية العامية أو الفصحى أو معادلة إكسيل مباشرة (مثل: احذف المطابق والفارق صفر، رتب بالسعر، اضف عمود ضريبة 14%، أو =SUM(I:I))
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2 shrink-0">
          <button
            type="button"
            onClick={onQuickPurgeZero}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-rose-800 bg-rose-50 border border-rose-200 rounded-lg hover:bg-rose-100 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
          >
            <Trash2 className="w-3.5 h-3.5" />
            <span>حذف المطابق والفارق (0)</span>
          </button>

          <button
            type="button"
            onClick={onQuickFullOrganize}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-teal-800 bg-teal-50 border border-teal-200 rounded-lg hover:bg-teal-100 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
          >
            <Sparkles className="w-3.5 h-3.5" />
            <span>تنظيم وترتيب دقيق للشيت</span>
          </button>

          {canUndo && (
            <button
              type="button"
              onClick={onUndoLast}
              className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-medium text-slate-700 bg-slate-100 rounded-lg hover:bg-slate-200 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              <span>تراجع عن آخر خطوة</span>
            </button>
          )}
        </div>
      </div>

      <form onSubmit={handleSubmit} className="flex flex-col sm:flex-row gap-2.5">
        <div className="relative flex-1">
          <label htmlFor="smart-command-input" className="sr-only">
            اكتب أمرك الكتابي لتنظيم الشيت أو حساب المعادلات
          </label>
          <input
            id="smart-command-input"
            type="text"
            value={commandInput}
            onChange={(e) => {
              setCommandInput(e.target.value);
              if (validationError) setValidationError(null);
            }}
            placeholder="اكتب أمرك هنا... مثال: احذف المطابق والفارق صفر ورتب الأصناف حسب السعر من الأعلى للأقل"
            className="w-full px-4 py-2.5 text-sm bg-slate-50 border border-slate-300 rounded-lg text-slate-900 placeholder:text-slate-400 focus:outline-none focus:bg-white focus:border-teal-700 focus:ring-2 focus:ring-teal-700/15 transition-all"
          />
        </div>
        <button
          type="submit"
          disabled={isProcessing}
          className="inline-flex items-center justify-center gap-2 px-5 py-2.5 text-sm font-semibold text-white bg-slate-900 rounded-lg hover:bg-slate-800 disabled:opacity-50 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
        >
          <Play className="w-3.5 h-3.5 fill-current" />
          <span>{isProcessing ? 'جاري التنفيذ...' : 'تنفيذ الأمر الذكي'}</span>
        </button>
      </form>

      {validationError && (
        <p role="alert" className="text-xs font-medium text-rose-700 mt-2">
          {validationError}
        </p>
      )}

      {/* Quick 1-click command buttons */}
      <div className="mt-3 pt-3 border-t border-slate-100 flex flex-wrap items-center gap-2">
        <span className="text-xs font-medium text-slate-500 ml-1">أوامر سريعة بضغطة واحدة:</span>
        {QUICK_EXAMPLES.map((item) => {
          const IconComponent = item.icon;
          return (
            <button
              key={item.label}
              type="button"
              onClick={() => onExecuteCommand(item.command)}
              className="inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200/80 hover:text-slate-900 rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer"
            >
              <IconComponent className="w-3 h-3 text-slate-500" />
              <span>{item.label}</span>
            </button>
          );
        })}
      </div>

      {/* Execution Feedback & Live Formula Result */}
      {(lastFeedback || evaluatedBanner) && (
        <div className="mt-3 pt-3 border-t border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-2 text-xs">
          {lastFeedback && (
            <div className="text-teal-800 font-medium">
              <span>تم التنفيذ: </span>
              <span>{lastFeedback}</span>
            </div>
          )}
          {evaluatedBanner && (
            <div className="flex items-center gap-3 bg-slate-900 text-white px-3.5 py-2 rounded-lg">
              <span className="text-slate-300">{evaluatedBanner.label}:</span>
              <code className="font-mono text-amber-300 dir-ltr" dir="ltr">
                {evaluatedBanner.formula}
              </code>
              <span className="text-slate-400">=</span>
              <span className="font-mono font-bold text-emerald-400 tabular-nums text-sm">
                {typeof evaluatedBanner.result === 'number'
                  ? evaluatedBanner.result.toLocaleString('en-US')
                  : evaluatedBanner.result}
              </span>
            </div>
          )}
        </div>
      )}
    </section>
  );
};
