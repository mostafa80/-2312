import React, { useState } from 'react';
import {
  CustomColumn,
  InventoryRow,
  ItemStatus,
  SelectedCellInfo,
  SortDirection,
  SortField,
} from '../types/excel';
import { ArrowUpDown, Trash2, Plus } from 'lucide-react';

interface SpreadsheetGridProps {
  rows: InventoryRow[];
  customColumns: CustomColumn[];
  selectedCell: SelectedCellInfo | null;
  onSelectCell: (cell: SelectedCellInfo) => void;
  onUpdateRow: (updatedRow: InventoryRow) => void;
  onDeleteRow: (rowId: string) => void;
  onRemoveCustomColumn: (colLetter: string) => void;
  onSortColumn: (field: SortField) => void;
  sortField: SortField;
  sortDirection: SortDirection;
  onOpenAddRow: () => void;
}

const STATUS_OPTIONS: ItemStatus[] = ['عجز', 'زيادة', 'مطابق', 'تالف', 'مرتجع'];

export function formatStatusText(status: ItemStatus): {
  prefix: string;
  colorClass: string;
} {
  switch (status) {
    case 'عجز':
      return { prefix: '−', colorClass: 'text-rose-700 font-semibold' };
    case 'زيادة':
      return { prefix: '+', colorClass: 'text-emerald-700 font-semibold' };
    case 'تالف':
      return { prefix: '!', colorClass: 'text-amber-700 font-semibold' };
    case 'مرتجع':
      return { prefix: '↩', colorClass: 'text-indigo-700 font-semibold' };
    case 'مطابق':
    default:
      return { prefix: '=', colorClass: 'text-slate-500 font-medium' };
  }
}

export const SpreadsheetGrid: React.FC<SpreadsheetGridProps> = ({
  rows,
  customColumns,
  selectedCell,
  onSelectCell,
  onUpdateRow,
  onDeleteRow,
  onRemoveCustomColumn,
  onSortColumn,
  sortField,
  sortDirection,
  onOpenAddRow,
}) => {
  const [editingCell, setEditingCell] = useState<{ rowId: string; field: string } | null>(null);
  const [editValue, setEditValue] = useState<string>('');

  const startEdit = (rowId: string, field: string, currentVal: string | number) => {
    setEditingCell({ rowId, field });
    setEditValue(String(currentVal));
  };

  const commitEdit = (row: InventoryRow, field: string) => {
    if (!editingCell) return;
    const updated = { ...row };
    if (field === 'unitPrice') {
      const n = Number.parseFloat(editValue);
      if (!Number.isNaN(n) && n >= 0) updated.unitPrice = n;
    } else if (field === 'bookQty') {
      const n = Number.parseFloat(editValue);
      if (!Number.isNaN(n) && n >= 0) {
        updated.bookQty = n;
        updated.variance = updated.actualQty - n;
      }
    } else if (field === 'actualQty') {
      const n = Number.parseFloat(editValue);
      if (!Number.isNaN(n) && n >= 0) {
        updated.actualQty = n;
        updated.variance = n - updated.bookQty;
      }
    } else if (field === 'variance') {
      const n = Number.parseFloat(editValue);
      if (!Number.isNaN(n)) {
        updated.variance = n;
        updated.actualQty = Math.max(0, updated.bookQty + n);
      }
    } else if (field === 'itemName' && editValue.trim()) {
      updated.itemName = editValue.trim();
    } else if (field === 'company' && editValue.trim()) {
      updated.company = editValue.trim();
    }
    setEditingCell(null);
    onUpdateRow(updated);
  };

  const totalBookQty = rows.reduce((sum, r) => sum + r.bookQty, 0);
  const totalActualQty = rows.reduce((sum, r) => sum + r.actualQty, 0);
  const totalVariance = rows.reduce((sum, r) => sum + r.variance, 0);
  const totalVarianceValue = rows.reduce((sum, r) => sum + r.varianceValue, 0);

  if (rows.length === 0) {
    return (
      <div className="bg-white border border-slate-200 rounded-xl p-10 text-center">
        <p className="text-base font-semibold text-slate-900">
          لا توجد أصناف مطابقة للفلتر الحالي في الجدول
        </p>
        <p className="text-xs text-slate-500 mt-1 mb-4">
          يمكنك استعادة الأصناف المستبعدة أو إضافة صنف جديد أو إعادة تحميل البيانات التجريبية.
        </p>
        <button
          type="button"
          onClick={onOpenAddRow}
          className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold text-white bg-teal-700 rounded-lg hover:bg-teal-800 transition-colors cursor-pointer"
        >
          <Plus className="w-3.5 h-3.5" />
          <span>إضافة صنف جديد للشيت</span>
        </button>
      </div>
    );
  }

  return (
    <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
      <div className="overflow-x-auto">
        <table className="w-full text-right border-collapse text-xs">
          <thead>
            {/* Row 1: Excel Column Coordinate Letters (A, B, C, D, E, F, G, H, I, J, K...) */}
            <tr className="bg-slate-100 border-b border-slate-200 text-slate-500 font-mono text-[11px] select-none">
              <th className="py-1.5 px-2.5 text-center border-l border-slate-200 w-10">#</th>
              <th className="py-1.5 px-3 text-center border-l border-slate-200">A</th>
              <th className="py-1.5 px-2.5 text-center border-l border-slate-200">B</th>
              <th className="py-1.5 px-3 text-center border-l border-slate-200">C</th>
              <th className="py-1.5 px-2.5 text-center border-l border-slate-200">D</th>
              <th className="py-1.5 px-3 text-center border-l border-slate-200">E</th>
              <th className="py-1.5 px-2.5 text-center border-l border-slate-200">F</th>
              <th className="py-1.5 px-2.5 text-center border-l border-slate-200">G</th>
              <th className="py-1.5 px-2.5 text-center border-l border-slate-200">H (=G-F)</th>
              <th className="py-1.5 px-3 text-center border-l border-slate-200">I (=H*E)</th>
              <th className="py-1.5 px-3 text-center border-l border-slate-200">J (=IF)</th>
              {customColumns.map((col) => (
                <th
                  key={col.id}
                  className="py-1.5 px-3 text-center border-l border-slate-200 bg-teal-50/60 text-teal-800"
                >
                  <div className="flex items-center justify-between gap-2">
                    <span>{col.colLetter}</span>
                    <button
                      type="button"
                      onClick={() => onRemoveCustomColumn(col.colLetter)}
                      aria-label={`حذف عمود ${col.title}`}
                      className="text-slate-400 hover:text-rose-600 transition-colors cursor-pointer"
                    >
                      ×
                    </button>
                  </div>
                </th>
              ))}
              <th className="py-1.5 px-2.5 text-center w-12">إجراء</th>
            </tr>

            {/* Row 2: Semantic Column Headers with Sort Controls */}
            <tr className="bg-slate-50 border-b border-slate-200 text-slate-700 font-semibold">
              <th className="py-2.5 px-2.5 text-center border-l border-slate-200 font-mono text-slate-400">
                م
              </th>
              <th className="py-2.5 px-3 border-l border-slate-200">
                <button
                  type="button"
                  onClick={() => onSortColumn('company')}
                  className="inline-flex items-center gap-1 hover:text-teal-700 transition-colors whitespace-nowrap cursor-pointer"
                >
                  <span>الشركة</span>
                  <ArrowUpDown className="w-3 h-3 text-slate-400" />
                </button>
              </th>
              <th className="py-2.5 px-2.5 border-l border-slate-200 whitespace-nowrap">الكود</th>
              <th className="py-2.5 px-3 border-l border-slate-200">
                <button
                  type="button"
                  onClick={() => onSortColumn('itemName')}
                  className="inline-flex items-center gap-1 hover:text-teal-700 transition-colors whitespace-nowrap cursor-pointer"
                >
                  <span>اسم الصنف</span>
                  <ArrowUpDown className="w-3 h-3 text-slate-400" />
                </button>
              </th>
              <th className="py-2.5 px-2.5 border-l border-slate-200 whitespace-nowrap">التصنيف</th>
              <th className="py-2.5 px-3 border-l border-slate-200 text-left">
                <button
                  type="button"
                  onClick={() => onSortColumn('unitPrice')}
                  className={`inline-flex items-center gap-1 hover:text-teal-700 transition-colors whitespace-nowrap cursor-pointer ${
                    sortField === 'unitPrice' ? 'text-teal-700' : ''
                  }`}
                >
                  <span>سعر الوحدة</span>
                  <ArrowUpDown className="w-3 h-3 text-slate-400" />
                  {sortField === 'unitPrice' && (
                    <span className="text-[10px]">{sortDirection === 'asc' ? '↑' : '↓'}</span>
                  )}
                </button>
              </th>
              <th className="py-2.5 px-2.5 border-l border-slate-200 text-left">
                <button
                  type="button"
                  onClick={() => onSortColumn('bookQty')}
                  className="inline-flex items-center gap-1 hover:text-teal-700 transition-colors whitespace-nowrap cursor-pointer"
                >
                  <span>الدفتر</span>
                  <ArrowUpDown className="w-3 h-3 text-slate-400" />
                </button>
              </th>
              <th className="py-2.5 px-2.5 border-l border-slate-200 text-left">
                <button
                  type="button"
                  onClick={() => onSortColumn('actualQty')}
                  className="inline-flex items-center gap-1 hover:text-teal-700 transition-colors whitespace-nowrap cursor-pointer"
                >
                  <span>الفعلي</span>
                  <ArrowUpDown className="w-3 h-3 text-slate-400" />
                </button>
              </th>
              <th className="py-2.5 px-2.5 border-l border-slate-200 text-left">
                <button
                  type="button"
                  onClick={() => onSortColumn('absVariance')}
                  className={`inline-flex items-center gap-1 hover:text-teal-700 transition-colors whitespace-nowrap cursor-pointer ${
                    sortField === 'absVariance' || sortField === 'variance' ? 'text-teal-700' : ''
                  }`}
                >
                  <span>الفارق</span>
                  <ArrowUpDown className="w-3 h-3 text-slate-400" />
                </button>
              </th>
              <th className="py-2.5 px-3 border-l border-slate-200 text-left">
                <button
                  type="button"
                  onClick={() => onSortColumn('varianceValue')}
                  className="inline-flex items-center gap-1 hover:text-teal-700 transition-colors whitespace-nowrap cursor-pointer"
                >
                  <span>قيمة الفارق</span>
                  <ArrowUpDown className="w-3 h-3 text-slate-400" />
                </button>
              </th>
              <th className="py-2.5 px-3 border-l border-slate-200">
                <button
                  type="button"
                  onClick={() => onSortColumn('status')}
                  className={`inline-flex items-center gap-1 hover:text-teal-700 transition-colors whitespace-nowrap cursor-pointer ${
                    sortField === 'status' ? 'text-teal-700' : ''
                  }`}
                >
                  <span>الحالة</span>
                  <ArrowUpDown className="w-3 h-3 text-slate-400" />
                </button>
              </th>
              {customColumns.map((col) => (
                <th
                  key={col.id}
                  className="py-2.5 px-3 border-l border-slate-200 bg-teal-50/40 text-teal-900 text-left whitespace-nowrap"
                >
                  <div>{col.title}</div>
                  <div className="text-[10px] font-mono text-teal-700 font-normal" dir="ltr">
                    {col.formulaTemplate.replace(/\{row\}/g, 'n')}
                  </div>
                </th>
              ))}
              <th className="py-2.5 px-2.5 text-center">حذف</th>
            </tr>
          </thead>

          <tbody className="divide-y divide-slate-200">
            {rows.map((row, index) => {
              const rowNum = index + 1;
              const statusStyle = formatStatusText(row.status);
              const isSelectedRow = selectedCell?.rowId === row.id;

              return (
                <tr
                  key={row.id}
                  className={`transition-colors hover:bg-slate-50/90 ${
                    isSelectedRow ? 'bg-teal-50/30' : ''
                  }`}
                >
                  {/* Row Number */}
                  <td className="py-2 px-2.5 text-center font-mono text-slate-400 bg-slate-50/60 border-l border-slate-200 tabular-nums">
                    {rowNum}
                  </td>

                  {/* Col A: Company */}
                  <td
                    onClick={() =>
                      onSelectCell({
                        rowIndex: rowNum,
                        rowId: row.id,
                        colLetter: 'A',
                        colLabel: 'الشركة',
                        rawValue: row.company,
                        formula: row.company,
                      })
                    }
                    onDoubleClick={() => startEdit(row.id, 'company', row.company)}
                    className="py-2 px-3 border-l border-slate-100 font-medium text-slate-800 whitespace-nowrap cursor-pointer"
                  >
                    {editingCell?.rowId === row.id && editingCell.field === 'company' ? (
                      <input
                        type="text"
                        aria-label="تعديل اسم الشركة"
                        autoFocus
                        value={editValue}
                        onChange={(e) => setEditValue(e.target.value)}
                        onBlur={() => commitEdit(row, 'company')}
                        onKeyDown={(e) => e.key === 'Enter' && commitEdit(row, 'company')}
                        className="w-36 px-1.5 py-0.5 text-xs border border-teal-600 rounded bg-white"
                      />
                    ) : (
                      row.company
                    )}
                  </td>

                  {/* Col B: Code */}
                  <td
                    onClick={() =>
                      onSelectCell({
                        rowIndex: rowNum,
                        rowId: row.id,
                        colLetter: 'B',
                        colLabel: 'الكود',
                        rawValue: row.code,
                        formula: row.code,
                      })
                    }
                    className="py-2 px-2.5 border-l border-slate-100 font-mono text-slate-500 whitespace-nowrap tabular-nums"
                    dir="ltr"
                  >
                    {row.code}
                  </td>

                  {/* Col C: Item Name */}
                  <td
                    onClick={() =>
                      onSelectCell({
                        rowIndex: rowNum,
                        rowId: row.id,
                        colLetter: 'C',
                        colLabel: 'اسم الصنف',
                        rawValue: row.itemName,
                        formula: row.itemName,
                      })
                    }
                    onDoubleClick={() => startEdit(row.id, 'itemName', row.itemName)}
                    className="py-2 px-3 border-l border-slate-100 font-semibold text-slate-900 cursor-pointer"
                  >
                    {editingCell?.rowId === row.id && editingCell.field === 'itemName' ? (
                      <input
                        type="text"
                        aria-label="تعديل اسم الصنف"
                        autoFocus
                        value={editValue}
                        onChange={(e) => setEditValue(e.target.value)}
                        onBlur={() => commitEdit(row, 'itemName')}
                        onKeyDown={(e) => e.key === 'Enter' && commitEdit(row, 'itemName')}
                        className="w-48 px-1.5 py-0.5 text-xs border border-teal-600 rounded bg-white"
                      />
                    ) : (
                      row.itemName
                    )}
                  </td>

                  {/* Col D: Category (unboxed quiet metadata) */}
                  <td className="py-2 px-2.5 border-l border-slate-100 text-slate-500 whitespace-nowrap">
                    {row.category}
                  </td>

                  {/* Col E: Unit Price (click or double click to edit) */}
                  <td
                    onClick={() =>
                      onSelectCell({
                        rowIndex: rowNum,
                        rowId: row.id,
                        colLetter: 'E',
                        colLabel: 'سعر الوحدة',
                        rawValue: row.unitPrice,
                        formula: String(row.unitPrice),
                      })
                    }
                    onDoubleClick={() => startEdit(row.id, 'unitPrice', row.unitPrice)}
                    className="py-2 px-3 border-l border-slate-100 text-left font-mono font-semibold text-slate-900 tabular-nums cursor-pointer"
                    dir="ltr"
                    title="انقر مرتين لتعديل السعر"
                  >
                    {editingCell?.rowId === row.id && editingCell.field === 'unitPrice' ? (
                      <input
                        type="number"
                        aria-label="تعديل سعر الوحدة"
                        autoFocus
                        value={editValue}
                        onChange={(e) => setEditValue(e.target.value)}
                        onBlur={() => commitEdit(row, 'unitPrice')}
                        onKeyDown={(e) => e.key === 'Enter' && commitEdit(row, 'unitPrice')}
                        className="w-20 px-1.5 py-0.5 text-xs font-mono border border-teal-600 rounded bg-white text-left"
                      />
                    ) : (
                      row.unitPrice.toLocaleString('en-US', {
                        minimumFractionDigits: 0,
                        maximumFractionDigits: 2,
                      })
                    )}
                  </td>

                  {/* Col F: Book Qty */}
                  <td
                    onClick={() =>
                      onSelectCell({
                        rowIndex: rowNum,
                        rowId: row.id,
                        colLetter: 'F',
                        colLabel: 'الرصيد الدفتري',
                        rawValue: row.bookQty,
                        formula: String(row.bookQty),
                      })
                    }
                    onDoubleClick={() => startEdit(row.id, 'bookQty', row.bookQty)}
                    className="py-2 px-2.5 border-l border-slate-100 text-left font-mono text-slate-600 tabular-nums cursor-pointer"
                    dir="ltr"
                  >
                    {editingCell?.rowId === row.id && editingCell.field === 'bookQty' ? (
                      <input
                        type="number"
                        aria-label="تعديل الرصيد الدفتري"
                        autoFocus
                        value={editValue}
                        onChange={(e) => setEditValue(e.target.value)}
                        onBlur={() => commitEdit(row, 'bookQty')}
                        onKeyDown={(e) => e.key === 'Enter' && commitEdit(row, 'bookQty')}
                        className="w-16 px-1.5 py-0.5 text-xs font-mono border border-teal-600 rounded bg-white text-left"
                      />
                    ) : (
                      row.bookQty
                    )}
                  </td>

                  {/* Col G: Actual Qty */}
                  <td
                    onClick={() =>
                      onSelectCell({
                        rowIndex: rowNum,
                        rowId: row.id,
                        colLetter: 'G',
                        colLabel: 'الجرد الفعلي',
                        rawValue: row.actualQty,
                        formula: String(row.actualQty),
                      })
                    }
                    onDoubleClick={() => startEdit(row.id, 'actualQty', row.actualQty)}
                    className="py-2 px-2.5 border-l border-slate-100 text-left font-mono text-slate-800 font-medium tabular-nums cursor-pointer"
                    dir="ltr"
                  >
                    {editingCell?.rowId === row.id && editingCell.field === 'actualQty' ? (
                      <input
                        type="number"
                        aria-label="تعديل الجرد الفعلي"
                        autoFocus
                        value={editValue}
                        onChange={(e) => setEditValue(e.target.value)}
                        onBlur={() => commitEdit(row, 'actualQty')}
                        onKeyDown={(e) => e.key === 'Enter' && commitEdit(row, 'actualQty')}
                        className="w-16 px-1.5 py-0.5 text-xs font-mono border border-teal-600 rounded bg-white text-left"
                      />
                    ) : (
                      row.actualQty
                    )}
                  </td>

                  {/* Col H: Variance (=G-F) */}
                  <td
                    onClick={() =>
                      onSelectCell({
                        rowIndex: rowNum,
                        rowId: row.id,
                        colLetter: 'H',
                        colLabel: 'الفارق (العدد)',
                        rawValue: row.variance,
                        formula: `=G${rowNum}-F${rowNum}`,
                      })
                    }
                    onDoubleClick={() => startEdit(row.id, 'variance', row.variance)}
                    className={`py-2 px-2.5 border-l border-slate-100 text-left font-mono font-bold tabular-nums cursor-pointer ${
                      row.variance < 0
                        ? 'text-rose-700'
                        : row.variance > 0
                        ? 'text-emerald-700'
                        : 'text-slate-400'
                    }`}
                    dir="ltr"
                  >
                    {editingCell?.rowId === row.id && editingCell.field === 'variance' ? (
                      <input
                        type="number"
                        aria-label="تعديل الفارق"
                        autoFocus
                        value={editValue}
                        onChange={(e) => setEditValue(e.target.value)}
                        onBlur={() => commitEdit(row, 'variance')}
                        onKeyDown={(e) => e.key === 'Enter' && commitEdit(row, 'variance')}
                        className="w-16 px-1.5 py-0.5 text-xs font-mono border border-teal-600 rounded bg-white text-left"
                      />
                    ) : row.variance > 0 ? (
                      `+${row.variance}`
                    ) : (
                      row.variance
                    )}
                  </td>

                  {/* Col I: Variance Value (=H*E) */}
                  <td
                    onClick={() =>
                      onSelectCell({
                        rowIndex: rowNum,
                        rowId: row.id,
                        colLetter: 'I',
                        colLabel: 'قيمة الفارق المالية',
                        rawValue: row.varianceValue,
                        formula: `=H${rowNum}*E${rowNum}`,
                      })
                    }
                    className={`py-2 px-3 border-l border-slate-100 text-left font-mono font-semibold tabular-nums cursor-pointer ${
                      row.varianceValue < 0
                        ? 'text-rose-700'
                        : row.varianceValue > 0
                        ? 'text-emerald-700'
                        : 'text-slate-400'
                    }`}
                    dir="ltr"
                  >
                    {row.varianceValue > 0
                      ? `+${row.varianceValue.toLocaleString('en-US')}`
                      : row.varianceValue.toLocaleString('en-US')}
                  </td>

                  {/* Col J: Status (Zero-Pill clean unboxed text + inline select on change) */}
                  <td
                    onClick={() =>
                      onSelectCell({
                        rowIndex: rowNum,
                        rowId: row.id,
                        colLetter: 'J',
                        colLabel: 'الحالة',
                        rawValue: row.status,
                        formula: `=IF(H${rowNum}<0,"عجز",IF(H${rowNum}>0,"زيادة","مطابق"))`,
                      })
                    }
                    className="py-2 px-3 border-l border-slate-100 whitespace-nowrap"
                  >
                    <select
                      aria-label={`حالة الصنف ${row.itemName}`}
                      value={row.status}
                      onChange={(e) =>
                        onUpdateRow({
                          ...row,
                          status: e.target.value as ItemStatus,
                        })
                      }
                      className={`bg-transparent text-xs focus:outline-none cursor-pointer ${statusStyle.colorClass}`}
                    >
                      {STATUS_OPTIONS.map((st) => (
                        <option key={st} value={st} className="text-slate-900 bg-white">
                          {st}
                        </option>
                      ))}
                    </select>
                  </td>

                  {/* Custom Formula Columns (K, L...) */}
                  {customColumns.map((col) => {
                    const val = row.customValues?.[col.colLetter] ?? 0;
                    const resolvedFormula = col.formulaTemplate.replace(/\{row\}/gi, String(rowNum));
                    return (
                      <td
                        key={col.id}
                        onClick={() =>
                          onSelectCell({
                            rowIndex: rowNum,
                            rowId: row.id,
                            colLetter: col.colLetter,
                            colLabel: col.title,
                            rawValue: val,
                            formula: resolvedFormula,
                          })
                        }
                        className="py-2 px-3 border-l border-slate-100 bg-teal-50/20 text-left font-mono font-semibold text-teal-900 tabular-nums cursor-pointer"
                        dir="ltr"
                      >
                        {typeof val === 'number' ? val.toLocaleString('en-US') : val}
                      </td>
                    );
                  })}

                  {/* Delete Row Button */}
                  <td className="py-2 px-2.5 text-center">
                    <button
                      type="button"
                      onClick={() => onDeleteRow(row.id)}
                      aria-label={`حذف الصنف ${row.itemName}`}
                      className="p-1 text-slate-400 hover:text-rose-600 rounded transition-colors cursor-pointer"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  </td>
                </tr>
              );
            })}
          </tbody>

          {/* Live Excel Formula Summary Footer */}
          <tfoot>
            <tr className="bg-slate-100 border-t-2 border-slate-300 font-semibold text-slate-900">
              <td colSpan={6} className="py-2.5 px-3 text-right">
                <span>الإجمالي التلقائي بالمعادلات ({rows.length} صنف نشط)</span>
                <span className="text-slate-400 mx-2">·</span>
                <span className="text-[11px] font-mono text-slate-500" dir="ltr">
                  =SUM(F:I)
                </span>
              </td>
              <td className="py-2.5 px-2.5 text-left font-mono tabular-nums border-l border-slate-200" dir="ltr">
                {totalBookQty.toLocaleString('en-US')}
              </td>
              <td className="py-2.5 px-2.5 text-left font-mono tabular-nums border-l border-slate-200" dir="ltr">
                {totalActualQty.toLocaleString('en-US')}
              </td>
              <td
                className={`py-2.5 px-2.5 text-left font-mono font-bold tabular-nums border-l border-slate-200 ${
                  totalVariance < 0
                    ? 'text-rose-700'
                    : totalVariance > 0
                    ? 'text-emerald-700'
                    : 'text-slate-600'
                }`}
                dir="ltr"
              >
                {totalVariance > 0 ? `+${totalVariance}` : totalVariance}
              </td>
              <td
                className={`py-2.5 px-3 text-left font-mono font-bold tabular-nums border-l border-slate-200 ${
                  totalVarianceValue < 0
                    ? 'text-rose-700'
                    : totalVarianceValue > 0
                    ? 'text-emerald-700'
                    : 'text-slate-600'
                }`}
                dir="ltr"
              >
                {totalVarianceValue > 0
                  ? `+${totalVarianceValue.toLocaleString('en-US')}`
                  : totalVarianceValue.toLocaleString('en-US')}
              </td>
              <td colSpan={2 + customColumns.length} className="py-2.5 px-3 text-xs text-slate-500">
                انقر على أي خلية لعرض معادلتها أو انقر مرتين لتعديل القيم مباشرة
              </td>
            </tr>
          </tfoot>
        </table>
      </div>
    </div>
  );
};
