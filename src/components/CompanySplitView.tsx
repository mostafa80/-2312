import React from 'react';
import { CompanySummary, InventoryRow } from '../types/excel';
import { Download, Filter } from 'lucide-react';
import { exportOrganizedWorkbook } from '../utils/xlsxIo';

interface CompanySplitViewProps {
  summaries: CompanySummary[];
  activeRows: InventoryRow[];
  onFilterByCompany: (companyName: string) => void;
}

export const CompanySplitView: React.FC<CompanySplitViewProps> = ({
  summaries,
  activeRows,
  onFilterByCompany,
}) => {
  return (
    <div className="space-y-6">
      <div className="bg-white border border-slate-200 rounded-xl p-5 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
        <div>
          <h2 className="text-base font-bold text-slate-900 font-display">
            تقسيم الشيت الموحد إلى جداول مستقلة لكل شركة ({summaries.length} شركات)
          </h2>
          <p className="text-xs text-slate-600 mt-0.5">
            تم فصل أصناف كل شركة تلقائياً وترتيبها حسب الحالة والسعر مع احتساب صافي العجز والزيادة لكل مورد.
          </p>
        </div>
      </div>

      {summaries.map((summary) => {
        const companyItems = activeRows.filter((r) => r.company === summary.company);

        return (
          <section
            key={summary.company}
            aria-label={`جدول ${summary.company}`}
            className="bg-white border border-slate-200 rounded-xl overflow-hidden"
          >
            {/* Company Header Bar */}
            <div className="px-5 py-4 bg-slate-50 border-b border-slate-200 flex flex-col lg:flex-row lg:items-center justify-between gap-4">
              <div>
                <h3 className="text-base font-bold text-slate-900 font-display">
                  {summary.company}
                </h3>
                {/* Zero-Pill metadata row with subtle separators */}
                <div className="flex flex-wrap items-center gap-2 text-xs text-slate-600 mt-1">
                  <span>الأصناف النشطة: {summary.activeCount}</span>
                  <span aria-hidden="true">·</span>
                  <span>المطابق المحذوف (فارق 0): {summary.purgedCount}</span>
                  <span aria-hidden="true">·</span>
                  <span className="text-rose-700 font-medium">
                    إجمالي العجز: {summary.totalShortageValue.toLocaleString('en-US')} ج.م
                  </span>
                  <span aria-hidden="true">·</span>
                  <span className="text-emerald-700 font-medium">
                    إجمالي الزيادة: {summary.totalSurplusValue.toLocaleString('en-US')} ج.م
                  </span>
                </div>
              </div>

              <div className="flex items-center gap-2 shrink-0">
                <button
                  type="button"
                  onClick={() => onFilterByCompany(summary.company)}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-lg hover:bg-slate-100 transition-colors whitespace-nowrap cursor-pointer"
                >
                  <Filter className="w-3.5 h-3.5" />
                  <span>فتح في الجدول الرئيسي</span>
                </button>
                <button
                  type="button"
                  onClick={() =>
                    exportOrganizedWorkbook(
                      companyItems,
                      [],
                      [],
                      `شيت_${summary.company.replace(/\s+/g, '_')}.xlsx`
                    )
                  }
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-slate-900 rounded-lg hover:bg-slate-800 transition-colors whitespace-nowrap cursor-pointer"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>تحميل شيت الشركة منفرداً</span>
                </button>
              </div>
            </div>

            {/* Company Items Sub-table */}
            {companyItems.length === 0 ? (
              <div className="p-6 text-center text-xs text-slate-500">
                جميع أصناف هذه الشركة كانت مطابقة (فارق 0) وتم استبعادها.
              </div>
            ) : (
              <div className="overflow-x-auto">
                <table className="w-full text-right border-collapse text-xs">
                  <thead>
                    <tr className="border-b border-slate-200 text-slate-500 bg-slate-50/50">
                      <th className="py-2 px-3">الكود</th>
                      <th className="py-2 px-3">اسم الصنف</th>
                      <th className="py-2 px-3 text-left">سعر الوحدة</th>
                      <th className="py-2 px-3 text-left">الدفتر</th>
                      <th className="py-2 px-3 text-left">الفعلي</th>
                      <th className="py-2 px-3 text-left">الفارق</th>
                      <th className="py-2 px-3 text-left">قيمة الفارق</th>
                      <th className="py-2 px-3">الحالة</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {companyItems.map((item) => (
                      <tr key={item.id} className="hover:bg-slate-50">
                        <td className="py-2 px-3 font-mono text-slate-500 tabular-nums" dir="ltr">
                          {item.code}
                        </td>
                        <td className="py-2 px-3 font-semibold text-slate-900">{item.itemName}</td>
                        <td className="py-2 px-3 text-left font-mono tabular-nums" dir="ltr">
                          {item.unitPrice.toLocaleString('en-US')}
                        </td>
                        <td className="py-2 px-3 text-left font-mono text-slate-600 tabular-nums" dir="ltr">
                          {item.bookQty}
                        </td>
                        <td className="py-2 px-3 text-left font-mono text-slate-800 tabular-nums" dir="ltr">
                          {item.actualQty}
                        </td>
                        <td
                          className={`py-2 px-3 text-left font-mono font-bold tabular-nums ${
                            item.variance < 0 ? 'text-rose-700' : 'text-emerald-700'
                          }`}
                          dir="ltr"
                        >
                          {item.variance > 0 ? `+${item.variance}` : item.variance}
                        </td>
                        <td
                          className={`py-2 px-3 text-left font-mono font-semibold tabular-nums ${
                            item.varianceValue < 0 ? 'text-rose-700' : 'text-emerald-700'
                          }`}
                          dir="ltr"
                        >
                          {item.varianceValue > 0
                            ? `+${item.varianceValue.toLocaleString('en-US')}`
                            : item.varianceValue.toLocaleString('en-US')}
                        </td>
                        <td
                          className={`py-2 px-3 font-semibold ${
                            item.status === 'عجز'
                              ? 'text-rose-700'
                              : item.status === 'زيادة'
                              ? 'text-emerald-700'
                              : 'text-amber-700'
                          }`}
                        >
                          {item.status}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </section>
        );
      })}
    </div>
  );
};
