import {
  CompanySummary,
  CustomColumn,
  EvaluatedFormulaBanner,
  InventoryRow,
  ItemStatus,
  OrganizeConfig,
  SortDirection,
  SortField,
  StructuredOperation,
} from '../types/excel';

export const STATUS_PRIORITY: Record<ItemStatus, number> = {
  'عجز': 1,
  'تالف': 2,
  'زيادة': 3,
  'مرتجع': 4,
  'مطابق': 5,
};

export const EXCEL_COLUMNS: { letter: string; key: keyof InventoryRow; label: string; isNumeric: boolean; defaultFormula?: (rowNum: number) => string }[] = [
  { letter: 'A', key: 'company', label: 'الشركة', isNumeric: false },
  { letter: 'B', key: 'code', label: 'الكود', isNumeric: false },
  { letter: 'C', key: 'itemName', label: 'اسم الصنف', isNumeric: false },
  { letter: 'D', key: 'category', label: 'التصنيف', isNumeric: false },
  { letter: 'E', key: 'unitPrice', label: 'سعر الوحدة', isNumeric: true },
  { letter: 'F', key: 'bookQty', label: 'الرصيد الدفتري', isNumeric: true },
  { letter: 'G', key: 'actualQty', label: 'الجرد الفعلي', isNumeric: true },
  { letter: 'H', key: 'variance', label: 'الفارق (العدد)', isNumeric: true, defaultFormula: (r) => `=G${r}-F${r}` },
  { letter: 'I', key: 'varianceValue', label: 'قيمة الفارق', isNumeric: true, defaultFormula: (r) => `=H${r}*E${r}` },
  { letter: 'J', key: 'status', label: 'الحالة', isNumeric: false, defaultFormula: (r) => `=IF(H${r}<0,"عجز",IF(H${r}>0,"زيادة","مطابق"))` },
];

/**
 * Normalizes Arabic numbers and strings for reliable matching
 */
export function normalizeArabicText(input: string): string {
  if (!input) return '';
  return String(input)
    .trim()
    .replace(/[أإآ]/g, 'ا')
    .replace(/ة/g, 'ه')
    .replace(/ى/g, 'ي')
    .replace(/[٠-٩]/g, (d) => String('٠١٢٣٤٥٦٧٨٩'.indexOf(d)))
    .replace(/\s+/g, ' ');
}

export function parseNumericCell(val: unknown): number {
  if (typeof val === 'number' && !Number.isNaN(val)) return val;
  if (val === null || val === undefined) return 0;
  const normalized = String(val)
    .trim()
    .replace(/[٠-٩]/g, (d) => String('٠١٢٣٤٥٦٧٨٩'.indexOf(d)))
    .replace(/[,،\sج.م$SARريال]/g, '');
  const parsed = Number.parseFloat(normalized);
  return Number.isNaN(parsed) ? 0 : parsed;
}

/**
 * Infers normalized ItemStatus from Arabic/English text and numeric variance
 */
export function inferItemStatus(rawStatus: unknown, variance: number): ItemStatus {
  const norm = normalizeArabicText(String(rawStatus ?? '')).toLowerCase();

  if (norm.includes('تالف') || norm.includes('هالك') || norm.includes('damaged') || norm.includes('expired')) {
    return 'تالف';
  }
  if (norm.includes('مرتجع') || norm.includes('رجيع') || norm.includes('return')) {
    return 'مرتجع';
  }
  if (Math.abs(variance) < 0.0001) {
    return 'مطابق';
  }
  if (norm.includes('مطابق') || norm.includes('متطابق') || norm.includes('صفر') || norm.includes('matched') || norm === 'ok') {
    if (variance < 0) return 'عجز';
    if (variance > 0) return 'زيادة';
    return 'مطابق';
  }
  if (norm.includes('عجز') || norm.includes('نقص') || norm.includes('فقد') || norm.includes('shortage') || norm.includes('deficit')) {
    return 'عجز';
  }
  if (norm.includes('زياده') || norm.includes('فائض') || norm.includes('surplus') || norm.includes('excess')) {
    return 'زيادة';
  }

  if (variance < 0) return 'عجز';
  if (variance > 0) return 'زيادة';
  return 'مطابق';
}

export function getCellValueByLetter(row: InventoryRow, colLetter: string): string | number {
  const upper = colLetter.toUpperCase();
  switch (upper) {
    case 'A':
      return row.company;
    case 'B':
      return row.code;
    case 'C':
      return row.itemName;
    case 'D':
      return row.category;
    case 'E':
      return row.unitPrice;
    case 'F':
      return row.bookQty;
    case 'G':
      return row.actualQty;
    case 'H':
      return row.variance;
    case 'I':
      return row.varianceValue;
    case 'J':
      return row.status;
    default:
      return row.customValues?.[upper] ?? 0;
  }
}

/**
 * Evaluates a row-level formula template such as `=E{row}*1.14` or `=G{row}*E{row}` or `=ABS(H{row})*E{row}`
 */
export function evaluateRowFormulaTemplate(row: InventoryRow, rowNumber: number, formulaTemplate: string): number | string {
  const expanded = formulaTemplate.replace(/\{row\}/gi, String(rowNumber));
  return evaluateExcelFormula(expanded, [row], rowNumber);
}

/**
 * Full Excel-compatible formula evaluator supporting:
 * - =SUM(E:E), =SUM(E1:E10)
 * - =AVERAGE(E:E), =MAX(E:E), =MIN(E:E), =COUNT(A:A)
 * - =COUNTIF(J:J, "عجز"), =COUNTIF(A:A, "شركة النور")
 * - =SUMIF(J:J, "عجز", I:I), =SUMIF(A:A, "شركة النور", I:I)
 * - =IF(H1<0, "عجز", "زيادة")
 * - =ROUND(E1*1.14, 2), =ABS(H1)*E1
 * - Row/cell arithmetic: =G2-F2, =E2*1.14, =(G2*E2)-100
 */
export function evaluateExcelFormula(
  rawFormula: string,
  rows: InventoryRow[],
  singleRowOverrideNumber?: number
): number | string {
  if (!rawFormula) return '';
  const trimmed = rawFormula.trim();
  if (!trimmed.startsWith('=')) {
    const num = Number(trimmed);
    return Number.isNaN(num) ? trimmed : num;
  }

  const expr = trimmed.slice(1).trim();

  // Helper to resolve a column range like "E:E" or "E1:E10" into array of values
  const getRangeValues = (rangeStr: string): (string | number)[] => {
    const clean = rangeStr.trim().toUpperCase();
    const fullColMatch = clean.match(/^([A-Z]):\1$/);
    if (fullColMatch) {
      const col = fullColMatch[1];
      return rows.map((r) => getCellValueByLetter(r, col));
    }
    const spanMatch = clean.match(/^([A-Z])(\d+):([A-Z])(\d+)$/);
    if (spanMatch) {
      const col = spanMatch[1];
      const startRow = Math.max(1, Number.parseInt(spanMatch[2], 10));
      const endRow = Math.min(rows.length, Number.parseInt(spanMatch[4], 10));
      const slice: (string | number)[] = [];
      for (let i = startRow; i <= endRow; i++) {
        const r = rows[i - 1];
        if (r) slice.push(getCellValueByLetter(r, col));
      }
      return slice;
    }
    return [];
  };

  // 1. SUMIF(range, criteria, [sum_range])
  const sumIfMatch = expr.match(/^SUMIF\s*\(\s*([^,]+)\s*,\s*([^,]+)\s*(?:,\s*([^)]+))?\)$/i);
  if (sumIfMatch) {
    const critVals = getRangeValues(sumIfMatch[1]);
    const rawCrit = sumIfMatch[2].trim().replace(/^["']|["']$/g, '');
    const sumVals = sumIfMatch[3] ? getRangeValues(sumIfMatch[3]) : critVals;
    const normCrit = normalizeArabicText(rawCrit).toLowerCase();

    let total = 0;
    for (let i = 0; i < critVals.length; i++) {
      const cellVal = critVals[i];
      const numVal = parseNumericCell(sumVals[i]);
      let matches = false;

      if (rawCrit.startsWith('<') || rawCrit.startsWith('>')) {
        const cNum = parseNumericCell(cellVal);
        const target = parseNumericCell(rawCrit.replace(/[<>=]/g, ''));
        if (rawCrit.startsWith('<=')) matches = cNum <= target;
        else if (rawCrit.startsWith('>=')) matches = cNum >= target;
        else if (rawCrit.startsWith('<')) matches = cNum < target;
        else if (rawCrit.startsWith('>')) matches = cNum > target;
      } else {
        matches = normalizeArabicText(String(cellVal)).toLowerCase().includes(normCrit);
      }

      if (matches) total += numVal;
    }
    return Number(total.toFixed(2));
  }

  // 2. COUNTIF(range, criteria)
  const countIfMatch = expr.match(/^COUNTIF\s*\(\s*([^,]+)\s*,\s*([^)]+)\)$/i);
  if (countIfMatch) {
    const critVals = getRangeValues(countIfMatch[1]);
    const rawCrit = countIfMatch[2].trim().replace(/^["']|["']$/g, '');
    const normCrit = normalizeArabicText(rawCrit).toLowerCase();

    let count = 0;
    for (const cellVal of critVals) {
      if (rawCrit.startsWith('<') || rawCrit.startsWith('>')) {
        const cNum = parseNumericCell(cellVal);
        const target = parseNumericCell(rawCrit.replace(/[<>=]/g, ''));
        if (rawCrit.startsWith('<') && cNum < target) count++;
        else if (rawCrit.startsWith('>') && cNum > target) count++;
      } else if (normalizeArabicText(String(cellVal)).toLowerCase().includes(normCrit)) {
        count++;
      }
    }
    return count;
  }

  // 3. Aggregate functions: SUM, AVERAGE, AVG, MAX, MIN, COUNT
  const aggMatch = expr.match(/^(SUM|AVERAGE|AVG|MAX|MIN|COUNT)\s*\(\s*([^)]+)\s*\)$/i);
  if (aggMatch) {
    const fn = aggMatch[1].toUpperCase();
    const vals = getRangeValues(aggMatch[2]);
    const nums = vals.map((v) => parseNumericCell(v)).filter((n) => !Number.isNaN(n));
    if (fn === 'COUNT') return vals.filter((v) => v !== '' && v !== null && v !== undefined).length;
    if (nums.length === 0) return 0;
    if (fn === 'SUM') return Number(nums.reduce((a, b) => a + b, 0).toFixed(2));
    if (fn === 'AVERAGE' || fn === 'AVG') {
      return Number((nums.reduce((a, b) => a + b, 0) / nums.length).toFixed(2));
    }
    if (fn === 'MAX') return Number(Math.max(...nums).toFixed(2));
    if (fn === 'MIN') return Number(Math.min(...nums).toFixed(2));
  }

  // Helper to resolve cell coordinates (e.g. E2, G1, H5) in row expressions
  const resolveRowTokens = (inputExpr: string): string => {
    return inputExpr.replace(/\b([A-Z])(\d+)\b/gi, (_, colLetter: string, rowNumStr: string) => {
      const rNum = Number.parseInt(rowNumStr, 10);
      const targetRow =
        singleRowOverrideNumber !== undefined && rNum === singleRowOverrideNumber
          ? rows[0]
          : rows[rNum - 1] ?? rows[0];
      if (!targetRow) return '0';
      const val = getCellValueByLetter(targetRow, colLetter);
      return typeof val === 'number' ? String(val) : `"${String(val)}"`;
    });
  };

  // 4. IF(condition, trueVal, falseVal)
  const ifMatch = expr.match(/^IF\s*\(\s*([^,]+)\s*,\s*([^,]+)\s*,\s*([^)]+)\)$/i);
  if (ifMatch) {
    const condStr = resolveRowTokens(ifMatch[1].trim());
    const trueVal = ifMatch[2].trim().replace(/^["']|["']$/g, '');
    const falseVal = ifMatch[3].trim().replace(/^["']|["']$/g, '');

    const cmpMatch = condStr.match(/^(-?\d+(?:\.\d+)?)\s*(<=|>=|<|>|=|==)\s*(-?\d+(?:\.\d+)?)$/);
    if (cmpMatch) {
      const left = Number.parseFloat(cmpMatch[1]);
      const op = cmpMatch[2];
      const right = Number.parseFloat(cmpMatch[3]);
      let passed = false;
      if (op === '<') passed = left < right;
      else if (op === '<=') passed = left <= right;
      else if (op === '>') passed = left > right;
      else if (op === '>=') passed = left >= right;
      else passed = Math.abs(left - right) < 0.0001;
      const chosen = passed ? trueVal : falseVal;
      const asNum = Number(chosen);
      return Number.isNaN(asNum) ? chosen : asNum;
    }
  }

  // 5. Arithmetic & ABS / ROUND expressions (e.g. =E1*1.14, =ABS(H1)*E1, =G1-F1)
  try {
    let mathExpr = resolveRowTokens(expr);
    mathExpr = mathExpr.replace(/ABS\s*\(([^)]+)\)/gi, 'Math.abs($1)');
    mathExpr = mathExpr.replace(/ROUND\s*\(([^,]+),\s*(\d+)\)/gi, '(Number(($1).toFixed($2)))');
    // Allow only safe numeric / Math tokens
    if (/^[0-9+\-*/().,\sA-Za-z_]+$/.test(mathExpr)) {
      const result = Function(`"use strict"; return (${mathExpr});`)();
      if (typeof result === 'number' && Number.isFinite(result)) {
        return Number(result.toFixed(2));
      }
    }
  } catch {
    return '#ERROR!';
  }

  return '#VALUE!';
}

export function recalculateRowMetrics(
  row: InventoryRow,
  autoStatus = true,
  customColumns: CustomColumn[] = [],
  rowIndex = 1
): InventoryRow {
  const unitPrice = Math.max(0, Number(row.unitPrice.toFixed(2)));
  const variance = Number(row.variance);
  const status = autoStatus
    ? row.status === 'تالف' || row.status === 'مرتجع'
      ? Math.abs(variance) < 0.0001
        ? 'مطابق'
        : row.status
      : inferItemStatus(row.status, variance)
    : row.status;
  const varianceValue = Number((variance * unitPrice).toFixed(2));

  const updatedRow: InventoryRow = {
    ...row,
    unitPrice,
    variance,
    varianceValue,
    status,
    customValues: { ...(row.customValues ?? {}) },
  };

  for (const col of customColumns) {
    updatedRow.customValues![col.colLetter] = evaluateRowFormulaTemplate(
      updatedRow,
      rowIndex,
      col.formulaTemplate
    );
  }

  return updatedRow;
}

export function recalculateAllCustomColumns(
  rows: InventoryRow[],
  customColumns: CustomColumn[]
): InventoryRow[] {
  if (customColumns.length === 0) return rows;
  return rows.map((row, idx) => recalculateRowMetrics(row, false, customColumns, idx + 1));
}

/**
 * Purges rows that are either "مطابق" (matched) or have a variance of 0
 */
export function purgeZeroAndMatched(
  rows: InventoryRow[],
  options: { removeZeroVariance?: boolean; removeMatchedStatus?: boolean } = {
    removeZeroVariance: true,
    removeMatchedStatus: true,
  }
): { keptRows: InventoryRow[]; removedRows: InventoryRow[] } {
  const removeZero = options.removeZeroVariance ?? true;
  const removeMatched = options.removeMatchedStatus ?? true;

  const keptRows: InventoryRow[] = [];
  const removedRows: InventoryRow[] = [];

  for (const row of rows) {
    const isZero = Math.abs(row.variance) < 0.0001;
    const isMatched = row.status === 'مطابق';

    if ((removeZero && isZero) || (removeMatched && isMatched)) {
      removedRows.push({
        ...row,
        status: isZero ? 'مطابق' : row.status,
      });
    } else {
      keptRows.push(row);
    }
  }

  return { keptRows, removedRows };
}

/**
 * Deduplicates items belonging to the same company and item name by combining quantities
 */
export function deduplicateCompanyItems(rows: InventoryRow[]): {
  deduplicated: InventoryRow[];
  mergedCount: number;
} {
  const map = new Map<string, InventoryRow>();
  let mergedCount = 0;

  for (const row of rows) {
    const key = `${normalizeArabicText(row.company)}::${normalizeArabicText(row.itemName)}`;
    const existing = map.get(key);
    if (!existing) {
      map.set(key, { ...row });
    } else {
      mergedCount++;
      const combinedBook = existing.bookQty + row.bookQty;
      const combinedActual = existing.actualQty + row.actualQty;
      const combinedVariance = existing.variance + row.variance;
      const latestPrice = row.unitPrice > 0 ? row.unitPrice : existing.unitPrice;
      const updated = recalculateRowMetrics(
        {
          ...existing,
          bookQty: combinedBook,
          actualQty: combinedActual,
          variance: combinedVariance,
          unitPrice: latestPrice,
          notes: existing.notes
            ? `${existing.notes} | تم دمج تكرار`
            : 'تم دمج صنف مكرر',
        },
        true
      );
      map.set(key, updated);
    }
  }

  return {
    deduplicated: Array.from(map.values()),
    mergedCount,
  };
}

function compareByField(a: InventoryRow, b: InventoryRow, field: SortField, direction: SortDirection): number {
  const dir = direction === 'asc' ? 1 : -1;

  switch (field) {
    case 'company':
      return a.company.localeCompare(b.company, 'ar') * dir;
    case 'itemName':
      return a.itemName.localeCompare(b.itemName, 'ar') * dir;
    case 'unitPrice':
      return (a.unitPrice - b.unitPrice) * dir;
    case 'bookQty':
      return (a.bookQty - b.bookQty) * dir;
    case 'actualQty':
      return (a.actualQty - b.actualQty) * dir;
    case 'variance':
      return (a.variance - b.variance) * dir;
    case 'absVariance':
      return (Math.abs(a.variance) - Math.abs(b.variance)) * dir;
    case 'varianceValue':
      return (Math.abs(a.varianceValue) - Math.abs(b.varianceValue)) * dir;
    case 'status':
      return (STATUS_PRIORITY[a.status] - STATUS_PRIORITY[b.status]) * dir;
    default:
      return 0;
  }
}

/**
 * Sorts and organizes rows according to multi-level rules
 */
export function sortInventoryRows(
  rows: InventoryRow[],
  primaryField: SortField,
  primaryDirection: SortDirection = 'asc',
  secondaryField: SortField = 'unitPrice',
  secondaryDirection: SortDirection = 'desc',
  groupByCompany = true
): InventoryRow[] {
  return [...rows].sort((a, b) => {
    if (groupByCompany && primaryField !== 'company') {
      const compCmp = a.company.localeCompare(b.company, 'ar');
      if (compCmp !== 0) return compCmp;
    }

    const pCmp = compareByField(a, b, primaryField, primaryDirection);
    if (pCmp !== 0) return pCmp;

    const sCmp = compareByField(a, b, secondaryField, secondaryDirection);
    if (sCmp !== 0) return sCmp;

    return a.itemName.localeCompare(b.itemName, 'ar');
  });
}

/**
 * Full precision organization pipeline:
 * 1. Recalculate and normalize statuses and variance values
 * 2. Optionally deduplicate same-company items
 * 3. Purge zero-variance and matched ("مطابق") items
 * 4. Group by company and sort by status & price
 */
export function organizeSheetPrecisely(
  rows: InventoryRow[],
  config: Partial<OrganizeConfig> = {},
  customColumns: CustomColumn[] = []
): {
  activeRows: InventoryRow[];
  purgedRows: InventoryRow[];
  mergedCount: number;
} {
  const fullConfig: OrganizeConfig = {
    removeZeroVariance: config.removeZeroVariance ?? true,
    removeMatchedStatus: config.removeMatchedStatus ?? true,
    deduplicateSameCompanyItems: config.deduplicateSameCompanyItems ?? true,
    autoRecalculateStatus: config.autoRecalculateStatus ?? true,
    primarySort: config.primarySort ?? 'status',
    primaryDirection: config.primaryDirection ?? 'asc',
    secondarySort: config.secondarySort ?? 'unitPrice',
    secondaryDirection: config.secondaryDirection ?? 'desc',
    groupByCompany: config.groupByCompany ?? true,
  };

  let working = rows.map((r, i) =>
    recalculateRowMetrics(r, fullConfig.autoRecalculateStatus, customColumns, i + 1)
  );
  let mergedCount = 0;

  if (fullConfig.deduplicateSameCompanyItems) {
    const dedupResult = deduplicateCompanyItems(working);
    working = dedupResult.deduplicated;
    mergedCount = dedupResult.mergedCount;
  }

  const { keptRows, removedRows } = purgeZeroAndMatched(working, {
    removeZeroVariance: fullConfig.removeZeroVariance,
    removeMatchedStatus: fullConfig.removeMatchedStatus,
  });

  const sortedActive = sortInventoryRows(
    keptRows,
    fullConfig.primarySort,
    fullConfig.primaryDirection,
    fullConfig.secondarySort,
    fullConfig.secondaryDirection,
    fullConfig.groupByCompany
  );

  return {
    activeRows: recalculateAllCustomColumns(sortedActive, customColumns),
    purgedRows: removedRows,
    mergedCount,
  };
}

/**
 * Parses a raw 2D array from an Excel sheet that contains multiple companies
 * on the same sheet (either via a Company column OR via section header rows like "شركة النور للأدوية").
 */
export function parseMultiCompanySheetRows(rawGrid: unknown[][]): InventoryRow[] {
  if (!rawGrid || rawGrid.length === 0) return [];

  let currentSectionCompany = 'شركة غير محددة';
  let colMap = {
    company: -1,
    code: -1,
    itemName: -1,
    category: -1,
    unitPrice: -1,
    bookQty: -1,
    actualQty: -1,
    variance: -1,
    status: -1,
    notes: -1,
  };

  const results: InventoryRow[] = [];
  let rowCounter = 1;

  for (let r = 0; r < rawGrid.length; r++) {
    const row = rawGrid[r];
    if (!Array.isArray(row)) continue;

    const nonEmptyCells = row
      .map((cell, idx) => ({ text: String(cell ?? '').trim(), idx }))
      .filter((c) => c.text !== '');

    if (nonEmptyCells.length === 0) continue;

    // 1. Check if this row is a Company Section Header on the sheet
    if (nonEmptyCells.length <= 2) {
      const combinedText = nonEmptyCells.map((c) => c.text).join(' ');
      const norm = normalizeArabicText(combinedText);
      if (
        norm.includes('شركه') ||
        norm.includes('مؤسسه') ||
        norm.includes('مصنع') ||
        norm.includes('مخزن') ||
        norm.includes('مورد') ||
        norm.toLowerCase().includes('company')
      ) {
        currentSectionCompany = combinedText
          .replace(/^(الشركة|شركة|اسم الشركة|المورد)\s*[:\-]?\s*/i, (m) =>
            m.includes(':') || m.includes('-') ? '' : m
          )
          .replace(/[===*#]/g, '')
          .trim();
        continue;
      }
    }

    // 2. Check if this row is a Header Row
    const normalizedCells = row.map((c) => normalizeArabicText(String(c ?? '')).toLowerCase());
    const hasItemHeader = normalizedCells.some(
      (h) =>
        h.includes('صنف') ||
        h.includes('منتج') ||
        h.includes('بيان') ||
        h.includes('اسم المادة') ||
        h.includes('item') ||
        h.includes('product')
    );
    const hasPriceOrQtyHeader = normalizedCells.some(
      (h) =>
        h.includes('سعر') ||
        h.includes('فارق') ||
        h.includes('فرق') ||
        h.includes('كميه') ||
        h.includes('رصيد') ||
        h.includes('price') ||
        h.includes('qty') ||
        h.includes('diff')
    );

    if (hasItemHeader && hasPriceOrQtyHeader) {
      colMap = {
        company: -1,
        code: -1,
        itemName: -1,
        category: -1,
        unitPrice: -1,
        bookQty: -1,
        actualQty: -1,
        variance: -1,
        status: -1,
        notes: -1,
      };

      normalizedCells.forEach((h, idx) => {
        if (!h) return;
        if (colMap.company === -1 && (h.includes('شركه') || h.includes('مورد') || h.includes('company') || h.includes('vendor'))) {
          colMap.company = idx;
        } else if (colMap.code === -1 && (h.includes('كود') || h.includes('رقم الصنف') || h.includes('باركود') || h.includes('code') || h.includes('sku'))) {
          colMap.code = idx;
        } else if (colMap.itemName === -1 && (h.includes('صنف') || h.includes('منتج') || h.includes('بيان') || h.includes('item') || h.includes('name'))) {
          colMap.itemName = idx;
        } else if (colMap.category === -1 && (h.includes('تصنيف') || h.includes('قسم') || h.includes('مجموعه') || h.includes('category'))) {
          colMap.category = idx;
        } else if (colMap.unitPrice === -1 && (h.includes('سعر') || h.includes('تكلفه') || h.includes('price') || h.includes('cost'))) {
          colMap.unitPrice = idx;
        } else if (colMap.bookQty === -1 && (h.includes('دفتري') || h.includes('مسجل') || h.includes('النظام') || h.includes('book') || h.includes('system'))) {
          colMap.bookQty = idx;
        } else if (colMap.actualQty === -1 && (h.includes('فعلي') || h.includes('جرد') || h.includes('actual') || h.includes('counted'))) {
          colMap.actualQty = idx;
        } else if (colMap.variance === -1 && (h.includes('فارق') || h.includes('فرق') || h.includes('عجز/زياده') || h.includes('diff') || h.includes('variance') || h === 'العدد')) {
          colMap.variance = idx;
        } else if (colMap.status === -1 && (h.includes('حاله') || h.includes('موقف') || h.includes('status') || h.includes('state'))) {
          colMap.status = idx;
        } else if (colMap.notes === -1 && (h.includes('ملاحظ') || h.includes('بيانات') || h.includes('note'))) {
          colMap.notes = idx;
        }
      });
      continue;
    }

    // 3. Parse Data Row
    const itemNameIdx = colMap.itemName !== -1 ? colMap.itemName : 1;
    const rawItemName = String(row[itemNameIdx] ?? '').trim();

    if (!rawItemName || normalizeArabicText(rawItemName).includes('اجمالي') || normalizeArabicText(rawItemName).includes('المجموع')) {
      continue;
    }

    const rowCompany =
      colMap.company !== -1 && String(row[colMap.company] ?? '').trim() !== ''
        ? String(row[colMap.company]).trim()
        : currentSectionCompany;

    if (colMap.company !== -1 && String(row[colMap.company] ?? '').trim() !== '') {
      currentSectionCompany = String(row[colMap.company]).trim();
    }

    const code =
      colMap.code !== -1 && String(row[colMap.code] ?? '').trim() !== ''
        ? String(row[colMap.code]).trim()
        : `ITM-${String(rowCounter).padStart(3, '0')}`;

    const category =
      colMap.category !== -1 && String(row[colMap.category] ?? '').trim() !== ''
        ? String(row[colMap.category]).trim()
        : 'عام';

    const unitPrice = colMap.unitPrice !== -1 ? parseNumericCell(row[colMap.unitPrice]) : parseNumericCell(row[2]);

    const hasBookAndActual = colMap.bookQty !== -1 && colMap.actualQty !== -1;
    const bookQty = colMap.bookQty !== -1 ? parseNumericCell(row[colMap.bookQty]) : 0;
    const actualQty = colMap.actualQty !== -1 ? parseNumericCell(row[colMap.actualQty]) : 0;

    let variance = 0;
    if (colMap.variance !== -1 && row[colMap.variance] !== undefined && String(row[colMap.variance]).trim() !== '') {
      variance = parseNumericCell(row[colMap.variance]);
    } else if (hasBookAndActual) {
      variance = actualQty - bookQty;
    } else {
      variance = parseNumericCell(row[3]);
    }

    const rawStatusCell = colMap.status !== -1 ? row[colMap.status] : row[4];
    let status = inferItemStatus(rawStatusCell, variance);

    const normStatusText = normalizeArabicText(String(rawStatusCell ?? ''));
    if (variance > 0 && (normStatusText.includes('عجز') || normStatusText.includes('نقص'))) {
      variance = -Math.abs(variance);
      status = 'عجز';
    }

    const notes = colMap.notes !== -1 ? String(row[colMap.notes] ?? '').trim() : '';

    results.push(
      recalculateRowMetrics(
        {
          id: `row-${rowCounter++}-${Date.now()}`,
          company: rowCompany,
          code,
          itemName: rawItemName,
          category,
          unitPrice,
          bookQty: hasBookAndActual ? bookQty : Math.max(0, 50 - variance),
          actualQty: hasBookAndActual ? actualQty : Math.max(0, 50),
          variance,
          varianceValue: variance * unitPrice,
          status,
          notes,
        },
        false
      )
    );
  }

  return results;
}

/**
 * Applies structured operations (including custom formula columns and live formula evaluations)
 */
export function applyStructuredOperations(
  activeRows: InventoryRow[],
  purgedRows: InventoryRow[],
  operations: StructuredOperation[],
  customColumns: CustomColumn[] = []
): {
  nextActiveRows: InventoryRow[];
  nextPurgedRows: InventoryRow[];
  nextCustomColumns: CustomColumn[];
  evaluatedFormulaBanner?: EvaluatedFormulaBanner;
  affectedCount: number;
  summaryParts: string[];
} {
  let workingActive = [...activeRows];
  let workingPurged = [...purgedRows];
  let workingCustomCols = [...customColumns];
  let evaluatedFormulaBanner: EvaluatedFormulaBanner | undefined;
  let affectedCount = 0;
  const summaryParts: string[] = [];

  for (const op of operations) {
    switch (op.type) {
      case 'PURGE_ZERO_AND_MATCHED': {
        const { keptRows, removedRows } = purgeZeroAndMatched(workingActive, {
          removeZeroVariance: true,
          removeMatchedStatus: true,
        });
        affectedCount += removedRows.length;
        workingActive = keptRows;
        workingPurged = [...workingPurged, ...removedRows];
        summaryParts.push(`تم حذف ${removedRows.length} صنف مطابق أو فارق (0)`);
        break;
      }

      case 'RESTORE_PURGED': {
        const count = workingPurged.length;
        workingActive = [...workingActive, ...workingPurged];
        workingPurged = [];
        affectedCount += count;
        summaryParts.push(`تم استعادة ${count} صنف من سجل المستبعدات`);
        break;
      }

      case 'ORGANIZE_PRECISE': {
        const res = organizeSheetPrecisely(
          workingActive,
          {
            removeZeroVariance: true,
            removeMatchedStatus: true,
            deduplicateSameCompanyItems: true,
            primarySort: op.field ?? 'status',
            primaryDirection: op.direction ?? 'asc',
            secondarySort: 'unitPrice',
            secondaryDirection: 'desc',
            groupByCompany: true,
          },
          workingCustomCols
        );
        affectedCount += res.purgedRows.length + res.mergedCount + res.activeRows.length;
        workingActive = res.activeRows;
        workingPurged = [...workingPurged, ...res.purgedRows];
        summaryParts.push(
          `تنظيم شامل: تم استبعاد ${res.purgedRows.length} صنف صفري/مطابق، ودمج ${res.mergedCount} تكرار، وترتيب الشركات حسب الحالة والسعر`
        );
        break;
      }

      case 'SORT_ROWS': {
        const field = op.field ?? 'unitPrice';
        const direction = op.direction ?? 'desc';
        workingActive = sortInventoryRows(
          workingActive,
          field,
          direction,
          'status',
          'asc',
          field === 'company'
        );
        affectedCount += workingActive.length;
        const fieldLabel: Record<SortField, string> = {
          company: 'الشركة',
          itemName: 'اسم الصنف',
          unitPrice: 'السعر',
          bookQty: 'الرصيد الدفتري',
          actualQty: 'الجرد الفعلي',
          variance: 'الفارق',
          absVariance: 'حجم الفارق',
          varianceValue: 'قيمة الفارق المالية',
          status: 'الحالة',
        };
        summaryParts.push(
          `تم ترتيب ${workingActive.length} صنف حسب ${fieldLabel[field]} (${direction === 'asc' ? 'تصاعدياً' : 'تنازلياً'})`
        );
        break;
      }

      case 'FILTER_KEEP_COMPANY': {
        if (op.companyName) {
          const target = normalizeArabicText(op.companyName);
          const before = workingActive.length;
          const kept = workingActive.filter((r) =>
            normalizeArabicText(r.company).includes(target)
          );
          const removed = workingActive.filter(
            (r) => !normalizeArabicText(r.company).includes(target)
          );
          workingActive = kept;
          workingPurged = [...workingPurged, ...removed];
          affectedCount += before - kept.length;
          summaryParts.push(`تم الإبقاء على أصناف "${op.companyName}" فقط (${kept.length} صنف)`);
        }
        break;
      }

      case 'DELETE_COMPANY': {
        if (op.companyName) {
          const target = normalizeArabicText(op.companyName);
          const kept = workingActive.filter(
            (r) => !normalizeArabicText(r.company).includes(target)
          );
          const removed = workingActive.filter((r) =>
            normalizeArabicText(r.company).includes(target)
          );
          workingActive = kept;
          workingPurged = [...workingPurged, ...removed];
          affectedCount += removed.length;
          summaryParts.push(`تم حذف ${removed.length} صنف تابع لشركة "${op.companyName}"`);
        }
        break;
      }

      case 'DELETE_BY_STATUS': {
        if (op.status) {
          const kept = workingActive.filter((r) => r.status !== op.status);
          const removed = workingActive.filter((r) => r.status === op.status);
          workingActive = kept;
          workingPurged = [...workingPurged, ...removed];
          affectedCount += removed.length;
          summaryParts.push(`تم حذف ${removed.length} صنف بحالة "${op.status}"`);
        }
        break;
      }

      case 'KEEP_ONLY_STATUS': {
        if (op.status) {
          const kept = workingActive.filter((r) => r.status === op.status);
          const removed = workingActive.filter((r) => r.status !== op.status);
          workingActive = kept;
          workingPurged = [...workingPurged, ...removed];
          affectedCount += removed.length;
          summaryParts.push(`تم حصر الجدول على حالة "${op.status}" فقط (${kept.length} صنف)`);
        }
        break;
      }

      case 'DELETE_BY_PRICE': {
        const threshold = op.value ?? 0;
        const operator = op.operator ?? 'lt';
        const shouldDelete = (price: number) => {
          if (operator === 'lt') return price < threshold;
          if (operator === 'lte') return price <= threshold;
          if (operator === 'gt') return price > threshold;
          if (operator === 'gte') return price >= threshold;
          return price === threshold;
        };
        const removed = workingActive.filter((r) => shouldDelete(r.unitPrice));
        const kept = workingActive.filter((r) => !shouldDelete(r.unitPrice));
        workingActive = kept;
        workingPurged = [...workingPurged, ...removed];
        affectedCount += removed.length;
        summaryParts.push(`تم حذف ${removed.length} صنف حسب شرط السعر (${threshold})`);
        break;
      }

      case 'KEEP_BY_PRICE': {
        const threshold = op.value ?? 0;
        const operator = op.operator ?? 'gt';
        const shouldKeep = (price: number) => {
          if (operator === 'gt') return price > threshold;
          if (operator === 'gte') return price >= threshold;
          if (operator === 'lt') return price < threshold;
          if (operator === 'lte') return price <= threshold;
          return price === threshold;
        };
        const kept = workingActive.filter((r) => shouldKeep(r.unitPrice));
        const removed = workingActive.filter((r) => !shouldKeep(r.unitPrice));
        workingActive = kept;
        workingPurged = [...workingPurged, ...removed];
        affectedCount += removed.length;
        summaryParts.push(`تم الإبقاء على ${kept.length} صنف يحقق شرط السعر (${threshold})`);
        break;
      }

      case 'ADJUST_PRICE_PERCENT': {
        const pct = op.value ?? 0;
        const targetCompany = op.companyName ? normalizeArabicText(op.companyName) : null;
        let modified = 0;
        workingActive = workingActive.map((row, idx) => {
          if (targetCompany && !normalizeArabicText(row.company).includes(targetCompany)) {
            return row;
          }
          modified++;
          const newPrice = Number((row.unitPrice * (1 + pct / 100)).toFixed(2));
          return recalculateRowMetrics({ ...row, unitPrice: newPrice }, false, workingCustomCols, idx + 1);
        });
        affectedCount += modified;
        summaryParts.push(
          `تم تعديل أسعار ${modified} صنف بنسبة ${pct > 0 ? `+${pct}%` : `${pct}%`}${
            op.companyName ? ` لشركة ${op.companyName}` : ''
          }`
        );
        break;
      }

      case 'DEDUPLICATE_ITEMS': {
        const { deduplicated, mergedCount } = deduplicateCompanyItems(workingActive);
        workingActive = deduplicated;
        affectedCount += mergedCount;
        summaryParts.push(`تم دمج ${mergedCount} صنف مكرر لنفس الشركة`);
        break;
      }

      case 'RECALCULATE_STATUSES': {
        workingActive = workingActive.map((r, idx) =>
          recalculateRowMetrics(r, true, workingCustomCols, idx + 1)
        );
        affectedCount += workingActive.length;
        summaryParts.push(`تم إعادة احتساب المعادلات والفروقات لـ ${workingActive.length} صنف`);
        break;
      }

      case 'ADD_FORMULA_COLUMN': {
        const nextLetterCode = 75 + workingCustomCols.length; // 'K', 'L', 'M'...
        const colLetter = String.fromCharCode(nextLetterCode);
        const title = op.columnTitle || `عمود معادلة (${colLetter})`;
        const formulaTemplate = op.formula || '=E{row}*1.14';

        const newCol: CustomColumn = {
          id: `custom-${colLetter}-${Date.now()}`,
          colLetter,
          title,
          formulaTemplate,
        };
        workingCustomCols = [...workingCustomCols, newCol];
        workingActive = recalculateAllCustomColumns(workingActive, workingCustomCols);
        affectedCount += workingActive.length;
        summaryParts.push(`تم إضافة عمود معادلة جديد "${title}" (${formulaTemplate.replace(/\{row\}/g, '1')})`);
        break;
      }

      case 'EVALUATE_FORMULA': {
        const formula = op.formula || '=SUM(I:I)';
        const val = evaluateExcelFormula(formula, workingActive);
        evaluatedFormulaBanner = {
          formula,
          label: op.explanation || 'نتيجة المعادلة المحسوبة',
          result: val,
        };
        affectedCount += workingActive.length;
        summaryParts.push(`${op.explanation || 'نتيجة المعادلة'} (${formula}) = ${val}`);
        break;
      }
    }
  }

  workingActive = recalculateAllCustomColumns(workingActive, workingCustomCols);

  return {
    nextActiveRows: workingActive,
    nextPurgedRows: workingPurged,
    nextCustomColumns: workingCustomCols,
    evaluatedFormulaBanner,
    affectedCount,
    summaryParts,
  };
}

/**
 * Fast, deterministic local Arabic/English written command & plain-explanation parser.
 * Understands conversational requests for both sheet organization AND Excel formulas!
 */
export function parseLocalWrittenCommand(
  commandText: string,
  availableCompanies: string[] = []
): StructuredOperation[] | null {
  const trimmed = commandText.trim();
  if (!trimmed) return null;

  // Direct Excel formula typed in command bar (e.g. "=SUM(E:E)" or "=SUMIF(J:J, "عجز", I:I)")
  if (trimmed.startsWith('=')) {
    return [
      {
        type: 'EVALUATE_FORMULA',
        formula: trimmed,
        explanation: 'تنفيذ معادلة إكسيل مباشرة',
      },
    ];
  }

  const norm = normalizeArabicText(trimmed).toLowerCase();
  const ops: StructuredOperation[] = [];

  // 1. Check if user is asking to add a new calculated/formula column ("اضف عمود", "اعمل عمود", "ضريبه", "قيمه الجرد")
  if (norm.includes('عمود')) {
    const pctMatch = norm.match(/(\d+(?:\.\d+)?)\s*%?/);
    if (norm.includes('ضريب') || norm.includes('شامل الضريب') || norm.includes('القيمه المضافه')) {
      const taxPct = pctMatch ? Number.parseFloat(pctMatch[1]) : 14;
      const multiplier = Number((1 + taxPct / 100).toFixed(4));
      ops.push({
        type: 'ADD_FORMULA_COLUMN',
        columnTitle: `السعر بالضريبة (${taxPct}%)`,
        formula: `=ROUND(E{row}*${multiplier}, 2)`,
      });
    } else if (norm.includes('خصم')) {
      const discPct = pctMatch ? Number.parseFloat(pctMatch[1]) : 10;
      const multiplier = Number((1 - discPct / 100).toFixed(4));
      ops.push({
        type: 'ADD_FORMULA_COLUMN',
        columnTitle: `السعر بعد خصم (${discPct}%)`,
        formula: `=ROUND(E{row}*${multiplier}, 2)`,
      });
    } else if (norm.includes('اجمالي الجرد') || norm.includes('قيمه الجرد') || norm.includes('قيمه المخزون')) {
      ops.push({
        type: 'ADD_FORMULA_COLUMN',
        columnTitle: 'إجمالي قيمة الجرد الفعلي',
        formula: '=ROUND(G{row}*E{row}, 2)',
      });
    } else if (norm.includes('نسبه الفارق') || norm.includes('نسبه العجز')) {
      ops.push({
        type: 'ADD_FORMULA_COLUMN',
        columnTitle: 'نسبة الفارق %',
        formula: '=ROUND((H{row}/F{row})*100, 1)',
      });
    } else {
      ops.push({
        type: 'ADD_FORMULA_COLUMN',
        columnTitle: 'إجمالي القيمة الفعلية',
        formula: '=ROUND(G{row}*E{row}, 2)',
      });
    }
    return ops;
  }

  // 2. Check if user is asking to calculate a summary equation ("احسب مجموع", "احسب متوسط", "كم عدد", "اجمالي سعر")
  if (
    norm.includes('احسب') ||
    norm.includes('مجموع') ||
    norm.includes('متوسط') ||
    norm.includes('اعلي سعر') ||
    norm.includes('اقل سعر') ||
    norm.includes('كم عدد')
  ) {
    if (norm.includes('متوسط') && norm.includes('سعر')) {
      return [
        {
          type: 'EVALUATE_FORMULA',
          formula: '=AVERAGE(E:E)',
          explanation: 'متوسط أسعار الأصناف الحالية',
        },
      ];
    }
    if (norm.includes('اعلي') || norm.includes('اكبر سعر') || norm.includes('اغلي')) {
      return [
        {
          type: 'EVALUATE_FORMULA',
          formula: '=MAX(E:E)',
          explanation: 'أعلى سعر وحدة في الجدول',
        },
      ];
    }
    if (norm.includes('عجز') && (norm.includes('مجموع') || norm.includes('اجمالي') || norm.includes('قيمه') || norm.includes('احسب'))) {
      return [
        {
          type: 'EVALUATE_FORMULA',
          formula: '=SUMIF(J:J, "عجز", I:I)',
          explanation: 'إجمالي القيمة المالية لأصناف العجز',
        },
      ];
    }
    if (norm.includes('زياده') && (norm.includes('مجموع') || norm.includes('اجمالي') || norm.includes('قيمه') || norm.includes('احسب'))) {
      return [
        {
          type: 'EVALUATE_FORMULA',
          formula: '=SUMIF(J:J, "زيادة", I:I)',
          explanation: 'إجمالي القيمة المالية لأصناف الزيادة',
        },
      ];
    }
    if (norm.includes('كم عدد') && norm.includes('عجز')) {
      return [
        {
          type: 'EVALUATE_FORMULA',
          formula: '=COUNTIF(J:J, "عجز")',
          explanation: 'عدد الأصناف التي بها عجز',
        },
      ];
    }
    return [
      {
        type: 'EVALUATE_FORMULA',
        formula: '=SUM(I:I)',
        explanation: 'صافي قيمة الفروقات المالية',
      },
    ];
  }

  // 3. Check for full precise organization command
  if (
    (norm.includes('نظم') || norm.includes('تنظيم شامل') || norm.includes('رتب الشيت')) &&
    (norm.includes('شامل') || norm.includes('دقيق') || norm.includes('كامل'))
  ) {
    ops.push({ type: 'ORGANIZE_PRECISE' });
    return ops;
  }

  // 4. Check for purge zero variance / matched
  if (
    (norm.includes('احذف') || norm.includes('استبعد') || norm.includes('ازل') || norm.includes('شيل') || norm.includes('تنظيف') || norm.includes('بدون')) &&
    (norm.includes('مطابق') || norm.includes('صفر') || norm.includes('0') || norm.includes('فارق صفر'))
  ) {
    ops.push({ type: 'PURGE_ZERO_AND_MATCHED' });
  }

  // 5. Check for restore purged
  if (norm.includes('استعاده') || norm.includes('استرجاع') || norm.includes('ارجع المحذوف')) {
    ops.push({ type: 'RESTORE_PURGED' });
    return ops;
  }

  // 6. Check for deduplication
  if (norm.includes('دمج') || norm.includes('مكرر') || norm.includes('تكرار')) {
    ops.push({ type: 'DEDUPLICATE_ITEMS' });
  }

  // 7. Check for percentage price adjustment (e.g. "زود السعر 10%" or "خصم 15%")
  const pctMatch = norm.match(/(\d+(?:\.\d+)?)\s*%/);
  if (pctMatch && (norm.includes('سعر') || norm.includes('اسعار') || norm.includes('زود') || norm.includes('خصم') || norm.includes('خفض'))) {
    const rawPct = Number.parseFloat(pctMatch[1]);
    const isDiscount = norm.includes('خصم') || norm.includes('خفض') || norm.includes('نقص') || norm.includes('قلل');
    const signedPct = isDiscount ? -rawPct : rawPct;

    let matchedCompany: string | undefined;
    for (const comp of availableCompanies) {
      const compNorm = normalizeArabicText(comp).toLowerCase();
      const shortComp = compNorm.replace('شركه', '').trim();
      if (shortComp && norm.includes(shortComp)) {
        matchedCompany = comp;
        break;
      }
    }

    ops.push({
      type: 'ADJUST_PRICE_PERCENT',
      value: signedPct,
      companyName: matchedCompany,
    });
  }

  // 8. Check for price threshold deletion or filtering
  const numberMatch = norm.match(/(?:اقل من|اصغر من|اكبر من|اعلي من|فوق|تحت)\s*(\d+(?:\.\d+)?)/);
  if (numberMatch && norm.includes('سعر')) {
    const val = Number.parseFloat(numberMatch[1]);
    const isLess = norm.includes('اقل من') || norm.includes('اصغر من') || norm.includes('تحت');
    const isDelete = norm.includes('احذف') || norm.includes('استبعد') || norm.includes('شيل');

    ops.push({
      type: isDelete ? 'DELETE_BY_PRICE' : 'KEEP_BY_PRICE',
      operator: isLess ? 'lt' : 'gt',
      value: val,
    });
  }

  // 9. Check for status filtering/deletion
  const statusTargets: { label: ItemStatus; keywords: string[] }[] = [
    { label: 'عجز', keywords: ['عجز', 'العجز', 'نقص'] },
    { label: 'زيادة', keywords: ['زياده', 'الزياده', 'فائض'] },
    { label: 'تالف', keywords: ['تالف', 'التالف', 'هالك'] },
    { label: 'مرتجع', keywords: ['مرتجع', 'المرتجع', 'رجيع'] },
    { label: 'مطابق', keywords: ['مطابق', 'المطابق'] },
  ];

  for (const st of statusTargets) {
    if (st.keywords.some((kw) => norm.includes(kw))) {
      if (st.label === 'مطابق' && ops.some((o) => o.type === 'PURGE_ZERO_AND_MATCHED')) {
        continue;
      }
      if (norm.includes('احذف') || norm.includes('استبعد') || norm.includes('بدون')) {
        ops.push({ type: 'DELETE_BY_STATUS', status: st.label });
      } else if (norm.includes('فقط') || norm.includes('اعرض') || norm.includes('فلتر') || norm.includes('حصر')) {
        ops.push({ type: 'KEEP_ONLY_STATUS', status: st.label });
      }
    }
  }

  // 10. Check for specific company filter or delete
  for (const comp of availableCompanies) {
    const compNorm = normalizeArabicText(comp).toLowerCase();
    const shortWords = compNorm
      .replace(/شركه|للتوزيع|للادويه|للتوريدات|للمستلزمات|التجاريه/g, '')
      .trim();
    if ((shortWords.length >= 3 && norm.includes(shortWords)) || norm.includes(compNorm)) {
      if (ops.some((o) => o.type === 'ADJUST_PRICE_PERCENT')) break;
      if (norm.includes('احذف') || norm.includes('استبعد')) {
        ops.push({ type: 'DELETE_COMPANY', companyName: comp });
      } else if (norm.includes('فقط') || norm.includes('اعرض') || norm.includes('فلتر') || norm.includes('شركه')) {
        ops.push({ type: 'FILTER_KEEP_COMPANY', companyName: comp });
      }
      break;
    }
  }

  // 11. Check for sorting instructions
  if (norm.includes('رتب') || norm.includes('ترتيب') || norm.includes('فرز')) {
    let field: SortField = 'unitPrice';
    let direction: SortDirection = 'desc';

    if (norm.includes('سعر') || norm.includes('الاسعار') || norm.includes('تكلفه') || norm.includes('اغلي') || norm.includes('ارخص')) {
      field = 'unitPrice';
      direction =
        norm.includes('تصاعد') || norm.includes('الاقل') || norm.includes('اصغر') || norm.includes('ارخص')
          ? 'asc'
          : 'desc';
    } else if (norm.includes('قيمه الفارق') || norm.includes('التاثير المالي')) {
      field = 'varianceValue';
      direction = norm.includes('تصاعد') || norm.includes('الاقل') ? 'asc' : 'desc';
    } else if (norm.includes('فارق') || norm.includes('فرق') || norm.includes('عجز')) {
      field = 'absVariance';
      direction = norm.includes('تصاعد') || norm.includes('الاقل') ? 'asc' : 'desc';
    } else if (norm.includes('حاله') || norm.includes('الحاله')) {
      field = 'status';
      direction = 'asc';
    } else if (norm.includes('شركه') || norm.includes('الشركات')) {
      field = 'company';
      direction = 'asc';
    } else if (norm.includes('ابجد') || norm.includes('اسم') || norm.includes('صنف')) {
      field = 'itemName';
      direction = norm.includes('تنازل') ? 'desc' : 'asc';
    }

    ops.push({ type: 'SORT_ROWS', field, direction });
  }

  return ops.length > 0 ? ops : null;
}

/**
 * Calculates per-company reconciliation summaries
 */
export function calculateCompanySummaries(
  activeRows: InventoryRow[],
  purgedRows: InventoryRow[]
): CompanySummary[] {
  const companyMap = new Map<string, CompanySummary>();

  const ensureCompany = (name: string): CompanySummary => {
    const existing = companyMap.get(name);
    if (existing) return existing;
    const init: CompanySummary = {
      company: name,
      activeCount: 0,
      purgedCount: 0,
      shortageCount: 0,
      surplusCount: 0,
      damagedCount: 0,
      returnedCount: 0,
      totalShortageValue: 0,
      totalSurplusValue: 0,
      netVarianceValue: 0,
      totalInventoryValue: 0,
    };
    companyMap.set(name, init);
    return init;
  };

  for (const row of activeRows) {
    const summary = ensureCompany(row.company);
    summary.activeCount += 1;
    summary.totalInventoryValue += row.actualQty * row.unitPrice;
    summary.netVarianceValue += row.varianceValue;

    if (row.status === 'عجز' || row.variance < 0) {
      summary.shortageCount += 1;
      summary.totalShortageValue += Math.abs(row.varianceValue);
    } else if (row.status === 'زيادة' || row.variance > 0) {
      summary.surplusCount += 1;
      summary.totalSurplusValue += Math.abs(row.varianceValue);
    }
    if (row.status === 'تالف') summary.damagedCount += 1;
    if (row.status === 'مرتجع') summary.returnedCount += 1;
  }

  for (const row of purgedRows) {
    const summary = ensureCompany(row.company);
    summary.purgedCount += 1;
  }

  return Array.from(companyMap.values()).sort((a, b) =>
    a.company.localeCompare(b.company, 'ar')
  );
}
