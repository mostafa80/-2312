import React, { useState } from 'react';
import { InventoryRow, ItemStatus } from '../types/excel';
import { downloadSampleMessyMultiCompanyExcel, readExcelFile } from '../utils/xlsxIo';
import { recalculateRowMetrics } from '../utils/excelEngine';
import { Upload, FileSpreadsheet, X, Plus, Download } from 'lucide-react';

interface UploadModalProps {
  isOpen: boolean;
  mode: 'upload' | 'add-row';
  onClose: () => void;
  onImportRows: (rows: InventoryRow[], autoPurgeZero: boolean) => void;
  onAddSingleRow: (row: InventoryRow) => void;
  existingCompanies: string[];
}

export const UploadModal: React.FC<UploadModalProps> = ({
  isOpen,
  mode,
  onClose,
  onImportRows,
  onAddSingleRow,
  existingCompanies,
}) => {
  const [autoPurgeOnUpload, setAutoPurgeOnUpload] = useState(true);
  const [uploadError, setUploadError] = useState<string | null>(null);
  const [isReading, setIsReading] = useState(false);

  // Add single row state
  const [company, setCompany] = useState(existingCompanies[0] || 'شركة النور للأدوية والتوزيع');
  const [code, setCode] = useState('');
  const [itemName, setItemName] = useState('');
  const [category, setCategory] = useState('عام');
  const [unitPrice, setUnitPrice] = useState('');
  const [bookQty, setBookQty] = useState('');
  const [actualQty, setActualQty] = useState('');
  const [status, setStatus] = useState<ItemStatus>('عجز');
  const [formError, setFormError] = useState<string | null>(null);

  if (!isOpen) return null;

  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setIsReading(true);
    setUploadError(null);
    try {
      const { rows } = await readExcelFile(file);
      if (rows.length === 0) {
        setUploadError('لم يتم العثور على صفوف بيانات صالحة داخل الملف المرفوع. تأكد من وجود أسماء الأصناف والأسعار.');
      } else {
        onImportRows(rows, autoPurgeOnUpload);
        onClose();
      }
    } catch {
      setUploadError('تعذر قراءة ملف الإكسيل. يرجى التأكد من أن الملف بصيغة .xlsx أو .xls أو .csv');
    } finally {
      setIsReading(false);
    }
  };

  const handleAddRowSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!company.trim() || !itemName.trim() || !unitPrice.trim()) {
      setFormError('يرجى إدخال اسم الشركة، واسم الصنف، وسعر الوحدة.');
      return;
    }
    const priceNum = Number.parseFloat(unitPrice);
    const bookNum = Number.parseFloat(bookQty || '0');
    const actualNum = Number.parseFloat(actualQty || '0');

    if (Number.isNaN(priceNum) || priceNum <= 0) {
      setFormError('يرجى إدخال سعر وحدة صحيح أكبر من صفر.');
      return;
    }

    const variance = actualNum - bookNum;
    const newRow = recalculateRowMetrics(
      {
        id: `manual-${Date.now()}`,
        company: company.trim(),
        code: code.trim() || `ITM-${Math.floor(100 + Math.random() * 900)}`,
        itemName: itemName.trim(),
        category: category.trim() || 'عام',
        unitPrice: priceNum,
        bookQty: bookNum,
        actualQty: actualNum,
        variance,
        varianceValue: variance * priceNum,
        status,
      },
      true
    );

    onAddSingleRow(newRow);
    setItemName('');
    setCode('');
    setUnitPrice('');
    setBookQty('');
    setActualQty('');
    setFormError(null);
    onClose();
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="modal-title"
      className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 p-4"
    >
      <div className="bg-white border border-slate-200 rounded-xl max-w-lg w-full p-6 shadow-xl">
        <div className="flex items-center justify-between pb-3 border-b border-slate-200">
          <h2 id="modal-title" className="text-base font-bold text-slate-900 font-display">
            {mode === 'upload'
              ? 'رفع شيت إكسيل مجمع لعدة شركات في نفس الصفحة'
              : 'إضافة صنف جديد إلى الشيت'}
          </h2>
          <button
            type="button"
            onClick={onClose}
            aria-label="إغلاق النافذة"
            className="p-1 text-slate-400 hover:text-slate-700 rounded cursor-pointer"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {mode === 'upload' ? (
          <div className="mt-4 space-y-4">
            <p className="text-xs text-slate-600 leading-relaxed">
              ارفع ملف إكسيل واحد (`.xlsx`, `.xls`, `.csv`) يحتوي على عدة شركات في نفس الورقة (سواء كانت الشركات مفصولة بعناوين أقسام أو داخل عمود الشركة). سيقوم البرنامج بالتعرف عليها وترتيبها تلقائياً.
            </p>

            <label
              htmlFor="excel-file-upload"
              className="flex flex-col items-center justify-center border-2 border-dashed border-slate-300 hover:border-teal-600 rounded-xl p-6 bg-slate-50 cursor-pointer transition-colors text-center"
            >
              <Upload className="w-8 h-8 text-teal-700 mb-2" />
              <span className="text-sm font-semibold text-slate-900">
                {isReading ? 'جاري قراءة وتحليل الشيت...' : 'اختر ملف Excel أو CSV من جهازك'}
              </span>
              <span className="text-xs text-slate-500 mt-1">
                يدعم ملفات .xlsx و .xls و .csv متعددة الشركات
              </span>
              <input
                id="excel-file-upload"
                type="file"
                accept=".xlsx,.xls,.csv"
                onChange={handleFileChange}
                className="sr-only"
              />
            </label>

            <label className="flex items-center gap-2 text-xs font-medium text-slate-800 cursor-pointer">
              <input
                type="checkbox"
                checked={autoPurgeOnUpload}
                onChange={(e) => setAutoPurgeOnUpload(e.target.checked)}
                className="rounded border-slate-300 text-teal-700 focus:ring-teal-600"
              />
              <span>
                حذف الأصناف المطابقة أو ذات الفارق (0) تلقائياً فور رفع الشيت وترتيب الباقي بدقة
              </span>
            </label>

            {uploadError && (
              <p role="alert" className="text-xs font-medium text-rose-700 bg-rose-50 p-2.5 rounded-lg">
                {uploadError}
              </p>
            )}

            <div className="pt-3 border-t border-slate-200 flex items-center justify-between gap-2">
              <span className="text-xs text-slate-500">ليس لديك ملف جاهز للتجربة الآن؟</span>
              <button
                type="button"
                onClick={downloadSampleMessyMultiCompanyExcel}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-teal-800 bg-teal-50 hover:bg-teal-100 rounded-lg transition-colors cursor-pointer"
              >
                <Download className="w-3.5 h-3.5" />
                <span>تحميل ملف إكسيل تجريبي لعدة شركات</span>
              </button>
            </div>
          </div>
        ) : (
          <form onSubmit={handleAddRowSubmit} className="mt-4 space-y-3">
            <div>
              <label htmlFor="add-row-company" className="block text-xs font-semibold text-slate-700 mb-1">
                اسم الشركة
              </label>
              <input
                id="add-row-company"
                type="text"
                value={company}
                onChange={(e) => setCompany(e.target.value)}
                placeholder="مثال: شركة النور للأدوية والتوزيع"
                className="w-full px-3 py-2 text-xs border border-slate-300 rounded-lg"
              />
            </div>

            <div className="grid grid-cols-3 gap-3">
              <div>
                <label htmlFor="add-row-code" className="block text-xs font-semibold text-slate-700 mb-1">
                  كود الصنف
                </label>
                <input
                  id="add-row-code"
                  type="text"
                  value={code}
                  onChange={(e) => setCode(e.target.value)}
                  placeholder="NR-109"
                  className="w-full px-3 py-2 text-xs font-mono border border-slate-300 rounded-lg"
                />
              </div>
              <div className="col-span-2">
                <label htmlFor="add-row-name" className="block text-xs font-semibold text-slate-700 mb-1">
                  اسم الصنف
                </label>
                <input
                  id="add-row-name"
                  type="text"
                  value={itemName}
                  onChange={(e) => setItemName(e.target.value)}
                  placeholder="مثال: أموكسيسيلين 500 مجم"
                  className="w-full px-3 py-2 text-xs border border-slate-300 rounded-lg"
                />
              </div>
            </div>

            <div className="grid grid-cols-3 gap-3">
              <div>
                <label htmlFor="add-row-price" className="block text-xs font-semibold text-slate-700 mb-1">
                  سعر الوحدة
                </label>
                <input
                  id="add-row-price"
                  type="number"
                  step="0.01"
                  value={unitPrice}
                  onChange={(e) => setUnitPrice(e.target.value)}
                  placeholder="120"
                  className="w-full px-3 py-2 text-xs font-mono border border-slate-300 rounded-lg"
                />
              </div>
              <div>
                <label htmlFor="add-row-book" className="block text-xs font-semibold text-slate-700 mb-1">
                  الرصيد الدفتري
                </label>
                <input
                  id="add-row-book"
                  type="number"
                  value={bookQty}
                  onChange={(e) => setBookQty(e.target.value)}
                  placeholder="50"
                  className="w-full px-3 py-2 text-xs font-mono border border-slate-300 rounded-lg"
                />
              </div>
              <div>
                <label htmlFor="add-row-actual" className="block text-xs font-semibold text-slate-700 mb-1">
                  الجرد الفعلي
                </label>
                <input
                  id="add-row-actual"
                  type="number"
                  value={actualQty}
                  onChange={(e) => setActualQty(e.target.value)}
                  placeholder="45"
                  className="w-full px-3 py-2 text-xs font-mono border border-slate-300 rounded-lg"
                />
              </div>
            </div>

            <div>
              <label htmlFor="add-row-status" className="block text-xs font-semibold text-slate-700 mb-1">
                الحالة
              </label>
              <select
                id="add-row-status"
                value={status}
                onChange={(e) => setStatus(e.target.value as ItemStatus)}
                className="w-full px-3 py-2 text-xs border border-slate-300 rounded-lg bg-white"
              >
                <option value="عجز">عجز</option>
                <option value="زيادة">زيادة</option>
                <option value="مطابق">مطابق (فارق 0)</option>
                <option value="تالف">تالف</option>
                <option value="مرتجع">مرتجع</option>
              </select>
            </div>

            {formError && (
              <p role="alert" className="text-xs font-medium text-rose-700">
                {formError}
              </p>
            )}

            <div className="pt-3 flex items-center justify-end gap-2">
              <button
                type="button"
                onClick={onClose}
                className="px-4 py-2 text-xs font-medium text-slate-600 hover:bg-slate-100 rounded-lg cursor-pointer"
              >
                إلغاء
              </button>
              <button
                type="submit"
                className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold text-white bg-teal-700 hover:bg-teal-800 rounded-lg cursor-pointer"
              >
                <Plus className="w-3.5 h-3.5" />
                <span>حفظ وإضافة للجدول</span>
              </button>
            </div>
          </form>
        )}
      </div>
    </div>
  );
};
