import test from 'node:test';
import assert from 'node:assert/strict';
import {
  applyStructuredOperations,
  calculateCompanySummaries,
  evaluateExcelFormula,
  organizeSheetPrecisely,
  parseLocalWrittenCommand,
  parseMultiCompanySheetRows,
  purgeZeroAndMatched,
} from './excelEngine';
import { InventoryRow } from '../types/excel';

const mockRows: InventoryRow[] = [
  {
    id: '1',
    company: 'شركة النور للأدوية',
    code: 'NR-101',
    itemName: 'باراسيتامول 500 مجم',
    category: 'مسكنات',
    unitPrice: 45,
    bookQty: 120,
    actualQty: 120,
    variance: 0,
    varianceValue: 0,
    status: 'مطابق',
  },
  {
    id: '2',
    company: 'شركة النور للأدوية',
    code: 'NR-102',
    itemName: 'أوجمنتين 1 جرام',
    category: 'مضادات حيوية',
    unitPrice: 180,
    bookQty: 50,
    actualQty: 42,
    variance: -8,
    varianceValue: -1440,
    status: 'عجز',
  },
  {
    id: '3',
    company: 'الشركة المتحدة للتوريدات',
    code: 'MT-201',
    itemName: 'جهاز قياس ضغط رقمي',
    category: 'أجهزة طبية',
    unitPrice: 950,
    bookQty: 15,
    actualQty: 18,
    variance: 3,
    varianceValue: 2850,
    status: 'زيادة',
  },
  {
    id: '4',
    company: 'الشركة المتحدة للتوريدات',
    code: 'MT-202',
    itemName: 'كمامات طبية علبة 50 قطعة',
    category: 'مستلزمات',
    unitPrice: 65,
    bookQty: 200,
    actualQty: 200,
    variance: 0,
    varianceValue: 0,
    status: 'مطابق',
  },
];

test('purgeZeroAndMatched removes rows with zero variance or مطابق status', () => {
  const { keptRows, removedRows } = purgeZeroAndMatched(mockRows);
  assert.equal(keptRows.length, 2);
  assert.equal(removedRows.length, 2);
  assert.equal(keptRows.every((r) => r.variance !== 0 && r.status !== 'مطابق'), true);
});

test('parseMultiCompanySheetRows parses single sheet with company section headers', () => {
  const rawSheet = [
    ['=== شركة النور للأدوية ==='],
    ['الكود', 'اسم الصنف', 'السعر', 'الرصيد الدفتري', 'الجرد الفعلي', 'الفارق', 'الحالة'],
    ['NR-1', 'بانادول إكسترا', 55, 100, 100, 0, 'مطابق'],
    ['NR-2', 'فيتامين سي فوار', 85, 40, 35, -5, 'عجز'],
    [],
    ['الشركة: مؤسسة الأمل للمستلزمات'],
    ['AM-1', 'قفازات لاتكس', 120, 60, 64, 4, 'زيادة'],
  ];

  const parsed = parseMultiCompanySheetRows(rawSheet);
  assert.equal(parsed.length, 3);
  assert.equal(parsed[0].company, 'شركة النور للأدوية');
  assert.equal(parsed[0].status, 'مطابق');
  assert.equal(parsed[1].company, 'شركة النور للأدوية');
  assert.equal(parsed[1].variance, -5);
  assert.equal(parsed[1].status, 'عجز');
  assert.equal(parsed[2].company, 'مؤسسة الأمل للمستلزمات');
  assert.equal(parsed[2].status, 'زيادة');
});

test('evaluateExcelFormula computes SUM, AVERAGE, SUMIF, COUNTIF and row arithmetic', () => {
  assert.equal(evaluateExcelFormula('=SUM(E:E)', mockRows), 1240);
  assert.equal(evaluateExcelFormula('=AVERAGE(E:E)', mockRows), 310);
  assert.equal(evaluateExcelFormula('=COUNTIF(J:J, "عجز")', mockRows), 1);
  assert.equal(evaluateExcelFormula('=SUMIF(J:J, "عجز", I:I)', mockRows), -1440);
  assert.equal(evaluateExcelFormula('=G2-F2', mockRows), -8);
  assert.equal(evaluateExcelFormula('=ROUND(E2*1.14, 2)', mockRows), 205.2);
});

test('parseLocalWrittenCommand and applyStructuredOperations handle plain Arabic explanations', () => {
  const companies = ['شركة النور للأدوية', 'الشركة المتحدة للتوريدات'];

  // 1. Delete matched and zero variance and sort by price
  const ops1 = parseLocalWrittenCommand('احذف المطابق والفارق صفر ورتب حسب السعر', companies);
  assert.ok(ops1);
  const res1 = applyStructuredOperations(mockRows, [], ops1);
  assert.equal(res1.nextActiveRows.length, 2);
  assert.equal(res1.nextPurgedRows.length, 2);
  assert.equal(res1.nextActiveRows[0].unitPrice, 950);

  // 2. Add a custom tax formula column by plain Arabic explanation
  const ops2 = parseLocalWrittenCommand('اضف عمود يحسب السعر بالضريبة 14%', companies);
  assert.ok(ops2);
  const res2 = applyStructuredOperations(res1.nextActiveRows, res1.nextPurgedRows, ops2);
  assert.equal(res2.nextCustomColumns.length, 1);
  assert.equal(res2.nextCustomColumns[0].colLetter, 'K');
  assert.equal(res2.nextActiveRows[0].customValues?.['K'], 1083); // 950 * 1.14
});

test('organizeSheetPrecisely deduplicates, purges zero variance, and sorts by company & status', () => {
  const rowsWithDup: InventoryRow[] = [
    ...mockRows,
    {
      id: '5',
      company: 'شركة النور للأدوية',
      code: 'NR-102',
      itemName: 'أوجمنتين 1 جرام',
      category: 'مضادات حيوية',
      unitPrice: 180,
      bookQty: 20,
      actualQty: 18,
      variance: -2,
      varianceValue: -360,
      status: 'عجز',
    },
  ];

  const organized = organizeSheetPrecisely(rowsWithDup);
  assert.equal(organized.mergedCount, 1);
  assert.equal(organized.purgedRows.length, 2);
  assert.equal(organized.activeRows.length, 2);

  const summaries = calculateCompanySummaries(organized.activeRows, organized.purgedRows);
  assert.equal(summaries.length, 2);
});
