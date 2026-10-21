import React, { useState } from 'react';
import { CustomColumn, InventoryRow } from '../types/excel';
import { evaluateExcelFormula } from '../utils/excelEngine';
import { Calculator, Plus, Play, Trash2 } from 'lucide-react';

interface FormulaStudioViewProps {
  activeRows: InventoryRow[];
  customColumns: CustomColumn[];
  onAddFormulaColumn: (title: string, formulaTemplate: string) => void;
  onRemoveCustomColumn: (colLetter: string) => void;
  onEvaluateFormulaInSheet: (formula: string, explanation: string) => void;
}

const PRESET_FORMULA_COLUMNS = [
  {
    title: 'السعر شامل ضريبة القيمة المضافة (14%)',
    formulaTemplate: '=ROUND(E{row}*1.14, 2)',
    description: 'يضرب سعر الوحدة (العمود E) × 1.14 مع التقريب لمنزلتين عشريتين لكل صنف.',
  },
  {
    title: 'إجمالي القيمة الفعلية للمخزون',
    formulaTemplate: '=ROUND(G{row}*E{row}, 2)',
    description: 'يضرب الجرد الفعلي (العمود G) × سعر الوحدة (العمود E) لحساب قيمة البضاعة الفعلية.',
  },
  {
    title: 'إجمالي القيمة الدفترية المسجلة',
    formulaTemplate: '=ROUND(F{row}*E{row}, 2)',
    description: 'يضرب الرصيد الدفتري (العمود F) × سعر الوحدة (العمود E).',
  },
  {
    title: 'السعر بعد خصم تجاري (10%)',
    formulaTemplate: '=ROUND(E{row}*0.90, 2)',
    description: 'يحسب صافي سعر الصنف بعد خصم 10% من سعر الوحدة.',
  },
  {
    title: 'نسبة الفارق المئوية %',
    formulaTemplate: '=ROUND((H{row}/F{row})*100, 1)',
    description: 'يحسب نسبة العجز أو الزيادة مقارنة بالرصيد الدفتري.',
  },
];

const BUILTIN_SUMMARY_EQUATIONS = [
  {
    label: 'إجمالي قيمة العجز المالي (SUMIF)',
    formula: '=SUMIF(J:J, "عجز", I:I)',
    explanation: 'جمع قيم الفارق المالية لجميع الأصناف التي حالتها عجز فقط',
  },
  {
    label: 'إجمالي قيمة الزيادة المالية (SUMIF)',
    formula: '=SUMIF(J:J, "زيادة", I:I)',
    explanation: 'جمع قيم الفارق المالية لجميع الأصناف التي حالتها زيادة فقط',
  },
  {
    label: 'صافي الفروقات الكلي للشيت (SUM)',
    formula: '=SUM(I:I)',
    explanation: 'مجموع عمود قيمة الفارق (I) بالكامل لجميع الشركات',
  },
  {
    label: 'متوسط أسعار الأصناف (AVERAGE)',
    formula: '=AVERAGE(E:E)',
    explanation: 'حساب المتوسط الحسابي لأسعار الوحدات في الشيت',
  },
  {
    label: 'أعلى سعر وحدة (MAX)',
    formula: '=MAX(E:E)',
    explanation: 'استخراج أغلى سعر صنف مسجل في الجدول',
  },
  {
    label: 'عدد الأصناف التي بها عجز (COUNTIF)',
    formula: '=COUNTIF(J:J, "عجز")',
    explanation: 'عد الأصناف التي تعاني من عجز في الجرد الفعلي',
  },
];

export const FormulaStudioView: React.FC<FormulaStudioViewProps> = ({
  activeRows,
  customColumns,
  onAddFormulaColumn,
  onRemoveCustomColumn,
  onEvaluateFormulaInSheet,
}) => {
  const [customTitle, setCustomTitle] = useState('');
  const [customFormula, setCustomFormula] = useState('=ROUND(E{row}*1.14, 2)');
  const [errorMsg, setErrorMsg] = useState<string | null>(null);

  const handleCreateCustomCol = (e: React.FormEvent) => {
    e.preventDefault();
    if (!customTitle.trim()) {
      setErrorMsg('يرجى كتابة اسم العمود الجديد (مثال: السعر بعد الضريبة).');
      return;
    }
    if (!customFormula.trim().startsWith('=')) {
      setErrorMsg('يجب أن تبدأ المعادلة بعلامة يساوي (=) مثل =E{row}*1.14');
      return;
    }
    setErrorMsg(null);
    onAddFormulaColumn(customTitle.trim(), customFormula.trim());
    setCustomTitle('');
  };

  return (
    <div className="space-y-6">
      {/* Section 1: Add Dynamic Formula Columns Without Complication */}
      <section className="bg-white border border-slate-200 rounded-xl p-5">
        <h2 className="text-base font-bold text-slate-900 font-display">
          بناء أعمدة المعادلات الذكية (بديل معادلات Excel بدون تعقيد)
        </h2>
        <p className="text-xs text-slate-600 mt-1">
          اختر معادلة جاهزة بضغطة واحدة ليتم إضافتها كعمود جديد في الشيت وتطبيقها على كل الأصناف والشركات تلقائياً، أو صمم معادلتك الخاصة.
        </p>

        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 mt-4">
          {PRESET_FORMULA_COLUMNS.map((preset) => (
            <div
              key={preset.title}
              className="border border-slate-200 rounded-lg p-4 flex flex-col justify-between bg-slate-50/50"
            >
              <div>
                <h3 className="text-sm font-bold text-slate-900">{preset.title}</h3>
                <code
                  className="inline-block mt-1 text-xs font-mono text-teal-800 bg-teal-50 px-2 py-0.5 rounded"
                  dir="ltr"
                >
                  {preset.formulaTemplate}
                </code>
                <p className="text-xs text-slate-600 mt-2 leading-relaxed">{preset.description}</p>
              </div>
              <button
                type="button"
                onClick={() => onAddFormulaColumn(preset.title, preset.formulaTemplate)}
                className="mt-4 inline-flex items-center justify-center gap-1.5 w-full px-3 py-2 text-xs font-semibold text-white bg-teal-700 rounded-lg hover:bg-teal-800 transition-colors cursor-pointer"
              >
                <Plus className="w-3.5 h-3.5" />
                <span>إضافة هذا العمود للشيت</span>
              </button>
            </div>
          ))}
        </div>

        {/* Custom Formula Column Form */}
        <form
          onSubmit={handleCreateCustomCol}
          className="mt-6 pt-5 border-t border-slate-200 grid grid-cols-1 md:grid-cols-12 gap-3 items-end"
        >
          <div className="md:col-span-5">
            <label htmlFor="custom-col-title" className="block text-xs font-semibold text-slate-700 mb-1">
              عنوان العمود الجديد
            </label>
            <input
              id="custom-col-title"
              type="text"
              value={customTitle}
              onChange={(e) => {
                setCustomTitle(e.target.value);
                if (errorMsg) setErrorMsg(null);
              }}
              placeholder="مثال: صافي القيمة بعد الضريبة"
              className="w-full px-3 py-2 text-xs bg-slate-50 border border-slate-300 rounded-lg focus:outline-none focus:bg-white focus:border-teal-700"
            />
          </div>
          <div className="md:col-span-5">
            <label htmlFor="custom-col-formula" className="block text-xs font-semibold text-slate-700 mb-1">
              معادلة العمود (استخدم E للسعر، F للدفتر، G للفعلي، H للفارق)
            </label>
            <input
              id="custom-col-formula"
              type="text"
              dir="ltr"
              value={customFormula}
              onChange={(e) => {
                setCustomFormula(e.target.value);
                if (errorMsg) setErrorMsg(null);
              }}
              placeholder="=ROUND(E{row}*1.14, 2)"
              className="w-full px-3 py-2 text-xs font-mono bg-slate-50 border border-slate-300 rounded-lg focus:outline-none focus:bg-white focus:border-teal-700"
            />
          </div>
          <div className="md:col-span-2">
            <button
              type="submit"
              className="w-full inline-flex items-center justify-center gap-1.5 px-4 py-2 text-xs font-semibold text-white bg-slate-900 rounded-lg hover:bg-slate-800 transition-colors cursor-pointer"
            >
              <Plus className="w-3.5 h-3.5" />
              <span>إنشاء العمود</span>
            </button>
          </div>
        </form>
        {errorMsg && (
          <p role="alert" className="text-xs font-medium text-rose-700 mt-2">
            {errorMsg}
          </p>
        )}

        {/* Active Custom Columns List */}
        {customColumns.length > 0 && (
          <div className="mt-5 pt-4 border-t border-slate-100">
            <h4 className="text-xs font-bold text-slate-700 mb-2">
              الأعمدة المحسوبة المضافة حالياً في الشيت ({customColumns.length}):
            </h4>
            <div className="flex flex-wrap gap-2">
              {customColumns.map((col) => (
                <div
                  key={col.id}
                  className="inline-flex items-center gap-2 px-3 py-1.5 bg-teal-50 border border-teal-200 rounded-lg text-xs text-teal-900"
                >
                  <span className="font-mono font-bold">[{col.colLetter}]</span>
                  <span className="font-semibold">{col.title}</span>
                  <code className="font-mono text-teal-700" dir="ltr">
                    {col.formulaTemplate}
                  </code>
                  <button
                    type="button"
                    onClick={() => onRemoveCustomColumn(col.colLetter)}
                    aria-label={`إزالة عمود ${col.title}`}
                    className="text-rose-600 hover:text-rose-800 cursor-pointer"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                  </button>
                </div>
              ))}
            </div>
          </div>
        )}
      </section>

      {/* Section 2: Live Aggregate Excel Functions Dashboard */}
      <section className="bg-white border border-slate-200 rounded-xl p-5">
        <div className="flex items-center gap-2 mb-3">
          <Calculator className="w-4 h-4 text-teal-700" />
          <h2 className="text-base font-bold text-slate-900 font-display">
            نتائج الدوال الإحصائية والشرطية الحية (SUM / SUMIF / COUNTIF / AVERAGE / MAX)
          </h2>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          {BUILTIN_SUMMARY_EQUATIONS.map((eq) => {
            const liveValue = evaluateExcelFormula(eq.formula, activeRows);
            return (
              <div
                key={eq.formula}
                className="border border-slate-200 rounded-lg p-4 flex flex-col justify-between"
              >
                <div>
                  <div className="text-xs font-semibold text-slate-800">{eq.label}</div>
                  <code
                    className="inline-block mt-1 text-xs font-mono text-slate-600 bg-slate-100 px-2 py-0.5 rounded"
                    dir="ltr"
                  >
                    {eq.formula}
                  </code>
                  <p className="text-[11px] text-slate-500 mt-1.5">{eq.explanation}</p>
                </div>

                <div className="mt-4 pt-3 border-t border-slate-100 flex items-center justify-between">
                  <span className="text-base font-mono font-bold text-slate-900 tabular-nums" dir="ltr">
                    {typeof liveValue === 'number' ? liveValue.toLocaleString('en-US') : liveValue}
                  </span>
                  <button
                    type="button"
                    onClick={() => onEvaluateFormulaInSheet(eq.formula, eq.label)}
                    className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-medium text-teal-800 bg-teal-50 hover:bg-teal-100 rounded transition-colors cursor-pointer"
                  >
                    <Play className="w-3 h-3" />
                    <span>تثبيت في الشريط</span>
                  </button>
                </div>
              </div>
            );
          })}
        </div>
      </section>
    </div>
  );
};
