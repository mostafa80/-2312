import * as XLSX from 'xlsx';
import { CustomColumn, InventoryRow } from '../types/excel';
import { calculateCompanySummaries, parseMultiCompanySheetRows } from './excelEngine';
import { getInitialMultiCompanyDataset } from './sampleData';

/**
 * Reads an uploaded Excel (.xlsx, .xls, .csv) file and extracts all multi-company rows
 * even if multiple companies are stacked on the same worksheet or across worksheets.
 */
export async function readExcelFile(file: File): Promise<{
  rows: InventoryRow[];
  sheetNames: string[];
}> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });

  let allRows: InventoryRow[] = [];
  for (const sheetName of workbook.SheetNames) {
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) continue;
    const rawGrid = XLSX.utils.sheet_to_json<unknown[]>(worksheet, {
      header: 1,
      defval: '',
    });
    const parsedFromSheet = parseMultiCompanySheetRows(rawGrid);
    if (parsedFromSheet.length > 0) {
      // If company wasn't set inside the sheet and the sheet tab name looks like a company name, assign it
      const normalizedRows = parsedFromSheet.map((r) => ({
        ...r,
        company:
          r.company === 'شركة غير محددة' && workbook.SheetNames.length > 1
            ? sheetName
            : r.company,
      }));
      allRows = [...allRows, ...normalizedRows];
    }
  }

  return {
    rows: allRows,
    sheetNames: workbook.SheetNames,
  };
}

/**
 * Exports the organized multi-company data to a real .xlsx file with:
 * 1. "الشيت الموحد المنظم" (All active organized rows + custom formula columns)
 * 2. "ملخص الشركات والفروقات" (Per-company financial & discrepancy summary)
 * 3. Separate dedicated worksheet for each company!
 */
export function exportOrganizedWorkbook(
  activeRows: InventoryRow[],
  purgedRows: InventoryRow[],
  customColumns: CustomColumn[] = [],
  filename = 'شيت_الشركات_المنظم.xlsx'
): void {
  const wb = XLSX.utils.book_new();

  const formatRowForExport = (row: InventoryRow, index: number) => {
    const base: Record<string, string | number> = {
      'م': index + 1,
      'الشركة': row.company,
      'كود الصنف': row.code,
      'اسم الصنف': row.itemName,
      'التصنيف': row.category,
      'سعر الوحدة': row.unitPrice,
      'الرصيد الدفتري': row.bookQty,
      'الجرد الفعلي': row.actualQty,
      'الفارق (العدد)': row.variance,
      'قيمة الفارق المالية': row.varianceValue,
      'الحالة': row.status,
    };

    for (const col of customColumns) {
      base[col.title] = row.customValues?.[col.colLetter] ?? '';
    }

    base['ملاحظات'] = row.notes ?? '';
    return base;
  };

  // Sheet 1: Unified Organized Sheet
  const unifiedData = activeRows.map((r, i) => formatRowForExport(r, i));
  const wsUnified = XLSX.utils.json_to_sheet(unifiedData);
  XLSX.utils.book_append_sheet(wb, wsUnified, 'الشيت الموحد المنظم');

  // Sheet 2: Companies Summary
  const summaries = calculateCompanySummaries(activeRows, purgedRows);
  const summaryData = summaries.map((s, idx) => ({
    'م': idx + 1,
    'اسم الشركة': s.company,
    'الأصناف النشطة (بعد الحذف)': s.activeCount,
    'الأصناف المطابقة المحذوفة (فارق 0)': s.purgedCount,
    'عدد أصناف العجز': s.shortageCount,
    'إجمالي قيمة العجز': s.totalShortageValue,
    'عدد أصناف الزيادة': s.surplusCount,
    'إجمالي قيمة الزيادة': s.totalSurplusValue,
    'صافي قيمة الفروقات': s.netVarianceValue,
    'إجمالي قيمة الجرد الفعلي': s.totalInventoryValue,
  }));
  const wsSummary = XLSX.utils.json_to_sheet(summaryData);
  XLSX.utils.book_append_sheet(wb, wsSummary, 'ملخص الشركات');

  // Sheet 3+: Individual Sheet per Company
  const companies = Array.from(new Set(activeRows.map((r) => r.company)));
  companies.forEach((companyName, idx) => {
    const compRows = activeRows
      .filter((r) => r.company === companyName)
      .map((r, i) => formatRowForExport(r, i));
    if (compRows.length > 0) {
      const wsComp = XLSX.utils.json_to_sheet(compRows);
      // Excel sheet names max 31 chars and cannot contain : \ / ? * [ ]
      const cleanName = companyName
        .replace(/[:\\/?*\[\]]/g, ' ')
        .trim()
        .slice(0, 26);
      const tabName = `${idx + 1}-${cleanName}`;
      XLSX.utils.book_append_sheet(wb, wsComp, tabName);
    }
  });

  // Optional Purged Archive Sheet if there are purged items
  if (purgedRows.length > 0) {
    const purgedData = purgedRows.map((r, i) => formatRowForExport(r, i));
    const wsPurged = XLSX.utils.json_to_sheet(purgedData);
    XLSX.utils.book_append_sheet(wb, wsPurged, 'المستبعدات (فارق 0 ومطابق)');
  }

  XLSX.writeFile(wb, filename);
}

/**
 * Generates and downloads a sample multi-company single-sheet messy Excel file
 * so the user can test uploading a file with multiple companies on the same sheet.
 */
export function downloadSampleMessyMultiCompanyExcel(): void {
  const sampleRows = getInitialMultiCompanyDataset();
  const companies = Array.from(new Set(sampleRows.map((r) => r.company)));

  const singleSheetGrid: (string | number)[][] = [
    ['كشف جرد مجمع لعدة شركات في صفحة واحدة (نموذج تجريبي قبل التنظيم)'],
    [],
  ];

  for (const comp of companies) {
    singleSheetGrid.push([`=== ${comp} ===`]);
    singleSheetGrid.push([
      'كود الصنف',
      'اسم الصنف',
      'التصنيف',
      'سعر الوحدة',
      'الرصيد الدفتري',
      'الجرد الفعلي',
      'الفارق',
      'الحالة',
      'ملاحظات',
    ]);
    const compItems = sampleRows.filter((r) => r.company === comp);
    for (const item of compItems) {
      singleSheetGrid.push([
        item.code,
        item.itemName,
        item.category,
        item.unitPrice,
        item.bookQty,
        item.actualQty,
        item.variance,
        item.status,
        item.notes ?? '',
      ]);
    }
    singleSheetGrid.push([]);
  }

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(singleSheetGrid);
  XLSX.utils.book_append_sheet(wb, ws, 'شيت مجمع لعدة شركات');
  XLSX.writeFile(wb, 'نموذج_شيت_مجمع_لعدة_شركات.xlsx');
}
