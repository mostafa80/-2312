import React, { useState, useMemo } from 'react';
import {
  AuditLogEntry,
  CustomColumn,
  EvaluatedFormulaBanner,
  InventoryRow,
  ItemStatus,
  SelectedCellInfo,
  SortDirection,
  SortField,
  StructuredOperation,
} from './types/excel';
import {
  applyStructuredOperations,
  calculateCompanySummaries,
  evaluateExcelFormula,
  organizeSheetPrecisely,
  parseLocalWrittenCommand,
  purgeZeroAndMatched,
  recalculateAllCustomColumns,
  recalculateRowMetrics,
  sortInventoryRows,
} from './utils/excelEngine';
import { getInitialMultiCompanyDataset } from './utils/sampleData';
import { exportOrganizedWorkbook } from './utils/xlsxIo';
import { TopBar, WorkspaceTab } from './components/TopBar';
import { SmartCommandBar } from './components/SmartCommandBar';
import { FormulaBar } from './components/FormulaBar';
import { SpreadsheetGrid } from './components/SpreadsheetGrid';
import { CompanySplitView } from './components/CompanySplitView';
import { FormulaStudioView } from './components/FormulaStudioView';
import { UploadModal } from './components/UploadModal';
import {
  FileSpreadsheet,
  Plus,
  RotateCcw,
  Search,
  Upload,
  History,
  CheckCircle2,
} from 'lucide-react';

export default function App() {
  const [activeRows, setActiveRows] = useState<InventoryRow[]>(() =>
    getInitialMultiCompanyDataset()
  );
  const [purgedRows, setPurgedRows] = useState<InventoryRow[]>([]);
  const [customColumns, setCustomColumns] = useState<CustomColumn[]>([]);
  const [activeTab, setActiveTab] = useState<WorkspaceTab>('grid');

  // Filtering & Sorting state
  const [selectedCompany, setSelectedCompany] = useState<string>('ALL');
  const [selectedStatus, setSelectedStatus] = useState<ItemStatus | 'ALL'>('ALL');
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [sortField, setSortField] = useState<SortField>('company');
  const [sortDirection, setSortDirection] = useState<SortDirection>('asc');

  // Formula bar & AI command state
  const [selectedCell, setSelectedCell] = useState<SelectedCellInfo | null>(null);
  const [activeFormulaSpec, setActiveFormulaSpec] = useState<{ formula: string; label: string }>({
    formula: '=SUM(I:I)',
    label: 'صافي قيمة الفروقات الحالية',
  });
  const [isProcessingCommand, setIsProcessingCommand] = useState(false);
  const [lastFeedback, setLastFeedback] = useState<string | null>(null);
  const [auditLogs, setAuditLogs] = useState<AuditLogEntry[]>([]);

  // Modal state
  const [modalMode, setModalMode] = useState<'upload' | 'add-row' | null>(null);

  const allCompanies = useMemo(() => {
    const set = new Set<string>();
    activeRows.forEach((r) => set.add(r.company));
    purgedRows.forEach((r) => set.add(r.company));
    return Array.from(set).sort((a, b) => a.localeCompare(b, 'ar'));
  }, [activeRows, purgedRows]);

  const companySummaries = useMemo(
    () => calculateCompanySummaries(activeRows, purgedRows),
    [activeRows, purgedRows]
  );

  // Filtered rows for display in the main unified grid
  const visibleRows = useMemo(() => {
    return activeRows.filter((row) => {
      if (selectedCompany !== 'ALL' && row.company !== selectedCompany) return false;
      if (selectedStatus !== 'ALL' && row.status !== selectedStatus) return false;
      if (searchQuery.trim() !== '') {
        const q = searchQuery.trim().toLowerCase();
        const matchName = row.itemName.toLowerCase().includes(q);
        const matchCode = row.code.toLowerCase().includes(q);
        const matchComp = row.company.toLowerCase().includes(q);
        if (!matchName && !matchCode && !matchComp) return false;
      }
      return true;
    });
  }, [activeRows, selectedCompany, selectedStatus, searchQuery]);

  // Top-level KPI metrics
  const kpiStats = useMemo(() => {
    const zeroOrMatchedRemaining = activeRows.filter(
      (r) => Math.abs(r.variance) < 0.0001 || r.status === 'مطابق'
    ).length;
    const totalShortageVal = activeRows
      .filter((r) => r.variance < 0 || r.status === 'عجز')
      .reduce((acc, r) => acc + Math.abs(r.varianceValue), 0);
    const totalSurplusVal = activeRows
      .filter((r) => r.variance > 0 || r.status === 'زيادة')
      .reduce((acc, r) => acc + Math.abs(r.varianceValue), 0);
    const netVarianceVal = activeRows.reduce((acc, r) => acc + r.varianceValue, 0);

    return {
      activeCount: activeRows.length,
      companiesCount: allCompanies.length,
      purgedCount: purgedRows.length,
      zeroOrMatchedRemaining,
      totalShortageVal,
      totalSurplusVal,
      netVarianceVal,
    };
  }, [activeRows, purgedRows, allCompanies]);

  const evaluatedBanner = useMemo<EvaluatedFormulaBanner>(() => {
    const targetRows = visibleRows.length > 0 ? visibleRows : activeRows;
    return {
      formula: activeFormulaSpec.formula,
      label: activeFormulaSpec.label,
      result: evaluateExcelFormula(activeFormulaSpec.formula, targetRows),
    };
  }, [activeFormulaSpec, visibleRows, activeRows]);

  const pushAuditLog = (
    commandText: string,
    summaryText: string,
    affectedCount: number,
    source: 'local-smart' | 'gemini-ai' | 'manual',
    prevActive: InventoryRow[],
    prevPurged: InventoryRow[],
    prevCols: CustomColumn[]
  ) => {
    const entry: AuditLogEntry = {
      id: `log-${Date.now()}`,
      timestamp: new Date().toLocaleTimeString('ar-EG', {
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
      }),
      commandText,
      summaryText,
      affectedCount,
      source,
      previousActiveRows: prevActive,
      previousPurgedRows: prevPurged,
      previousCustomColumns: prevCols,
    };
    setAuditLogs((prev) => [entry, ...prev.slice(0, 14)]);
    setLastFeedback(summaryText);
  };

  const handleQuickPurgeZero = () => {
    const prevActive = [...activeRows];
    const prevPurged = [...purgedRows];
    const { keptRows, removedRows } = purgeZeroAndMatched(activeRows, {
      removeZeroVariance: true,
      removeMatchedStatus: true,
    });

    if (removedRows.length === 0) {
      setLastFeedback('الشيت نظيف بالفعل — لا توجد أصناف مطابقة أو ذات فارق (0) لحذفها.');
      return;
    }

    const updatedActive = recalculateAllCustomColumns(keptRows, customColumns);
    setActiveRows(updatedActive);
    setPurgedRows([...purgedRows, ...removedRows]);

    pushAuditLog(
      'حذف المطابق والفارق (0)',
      `تم حذف ${removedRows.length} صنف مطابق أو فارق (0) من الشيت بنجاح.`,
      removedRows.length,
      'manual',
      prevActive,
      prevPurged,
      customColumns
    );
  };

  const handleQuickFullOrganize = () => {
    const prevActive = [...activeRows];
    const prevPurged = [...purgedRows];
    const res = organizeSheetPrecisely(
      activeRows,
      {
        removeZeroVariance: true,
        removeMatchedStatus: true,
        deduplicateSameCompanyItems: true,
        primarySort: 'status',
        primaryDirection: 'asc',
        secondarySort: 'unitPrice',
        secondaryDirection: 'desc',
        groupByCompany: true,
      },
      customColumns
    );

    setActiveRows(res.activeRows);
    setPurgedRows([...purgedRows, ...res.purgedRows]);
    setSortField('status');
    setSortDirection('asc');

    pushAuditLog(
      'تنظيم وترتيب شامل للشيت',
      `تم تنظيم الشيت بدقة: حذف ${res.purgedRows.length} صنف صفري/مطابق، وترتيب الشركات حسب الحالة والسعر.`,
      res.purgedRows.length + res.activeRows.length,
      'manual',
      prevActive,
      prevPurged,
      customColumns
    );
  };

  const handleUndoLast = () => {
    if (auditLogs.length === 0) return;
    const [latest, ...rest] = auditLogs;
    setActiveRows(latest.previousActiveRows);
    setPurgedRows(latest.previousPurgedRows);
    setCustomColumns(latest.previousCustomColumns);
    setAuditLogs(rest);
    setLastFeedback(`تم التراجع عن: "${latest.commandText}" واستعادة الحالة السابقة.`);
  };

  const handleExecuteWrittenCommand = async (commandText: string) => {
    setIsProcessingCommand(true);
    const prevActive = [...activeRows];
    const prevPurged = [...purgedRows];
    const prevCols = [...customColumns];

    try {
      // 1. Try instant local Arabic NLP & formula parser first for <5ms speed
      const localOps = parseLocalWrittenCommand(commandText, allCompanies);
      if (localOps && localOps.length > 0) {
        const applied = applyStructuredOperations(
          activeRows,
          purgedRows,
          localOps,
          customColumns
        );
        setActiveRows(applied.nextActiveRows);
        setPurgedRows(applied.nextPurgedRows);
        setCustomColumns(applied.nextCustomColumns);
        if (applied.evaluatedFormulaBanner) {
          setActiveFormulaSpec({
            formula: applied.evaluatedFormulaBanner.formula,
            label: applied.evaluatedFormulaBanner.label,
          });
        }

        const summary = applied.summaryParts.join(' + ');
        pushAuditLog(
          commandText,
          summary,
          applied.affectedCount,
          'local-smart',
          prevActive,
          prevPurged,
          prevCols
        );
        return;
      }

      // 2. If complex conversational explanation, call server-side Gemini AI interpreter
      const response = await fetch('/api/ai-command', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          commandText,
          companies: allCompanies,
          sampleItems: activeRows.slice(0, 6).map((r) => ({
            company: r.company,
            itemName: r.itemName,
            unitPrice: r.unitPrice,
            variance: r.variance,
            status: r.status,
          })),
        }),
      });

      const data = await response.json();
      const ops: StructuredOperation[] = Array.isArray(data.operations)
        ? data.operations
        : [{ type: 'ORGANIZE_PRECISE' }];

      const applied = applyStructuredOperations(
        activeRows,
        purgedRows,
        ops,
        customColumns
      );
      setActiveRows(applied.nextActiveRows);
      setPurgedRows(applied.nextPurgedRows);
      setCustomColumns(applied.nextCustomColumns);
      if (applied.evaluatedFormulaBanner) {
        setActiveFormulaSpec({
          formula: applied.evaluatedFormulaBanner.formula,
          label: applied.evaluatedFormulaBanner.label,
        });
      }

      const summary =
        applied.summaryParts.length > 0
          ? applied.summaryParts.join(' + ')
          : data.assistantMessage || 'تم تنفيذ الأمر بنجاح.';

      pushAuditLog(
        commandText,
        summary,
        applied.affectedCount,
        data.source === 'gemini-ai' ? 'gemini-ai' : 'local-smart',
        prevActive,
        prevPurged,
        prevCols
      );
    } catch {
      setLastFeedback('حدث خطأ أثناء الاتصال بالخادم، يرجى المحاولة مرة أخرى.');
    } finally {
      setIsProcessingCommand(false);
    }
  };

  const handleSortColumn = (field: SortField) => {
    const nextDir: SortDirection =
      sortField === field && sortDirection === 'desc' ? 'asc' : 'desc';
    setSortField(field);
    setSortDirection(nextDir);

    const sorted = sortInventoryRows(
      activeRows,
      field,
      nextDir,
      'unitPrice',
      'desc',
      field === 'company'
    );
    setActiveRows(recalculateAllCustomColumns(sorted, customColumns));
  };

  const handleUpdateSingleRow = (updatedRow: InventoryRow) => {
    const next = activeRows.map((row, idx) =>
      row.id === updatedRow.id
        ? recalculateRowMetrics(updatedRow, true, customColumns, idx + 1)
        : row
    );
    setActiveRows(next);
  };

  const handleDeleteSingleRow = (rowId: string) => {
    const target = activeRows.find((r) => r.id === rowId);
    if (!target) return;
    const prevActive = [...activeRows];
    const prevPurged = [...purgedRows];

    const nextActive = recalculateAllCustomColumns(
      activeRows.filter((r) => r.id !== rowId),
      customColumns
    );
    setActiveRows(nextActive);
    setPurgedRows([...purgedRows, target]);

    pushAuditLog(
      `حذف الصنف: ${target.itemName}`,
      `تم نقل الصنف "${target.itemName}" إلى سجل المستبعدات.`,
      1,
      'manual',
      prevActive,
      prevPurged,
      customColumns
    );
  };

  const handleApplyFormulaBarInput = (input: string) => {
    if (input.startsWith('=')) {
      const val = evaluateExcelFormula(input, visibleRows.length > 0 ? visibleRows : activeRows);
      setActiveFormulaSpec({
        formula: input,
        label: 'نتيجة شريط المعادلات (fx)',
      });
      setLastFeedback(`تم حساب المعادلة ${input} بنجاح = ${val}`);
      return;
    }

    // If user edited a selected numeric cell via the formula bar
    if (selectedCell) {
      const targetRow = activeRows.find((r) => r.id === selectedCell.rowId);
      if (targetRow) {
        const num = Number.parseFloat(input);
        const updated = { ...targetRow };
        if (selectedCell.colLetter === 'E' && !Number.isNaN(num)) updated.unitPrice = num;
        else if (selectedCell.colLetter === 'F' && !Number.isNaN(num)) {
          updated.bookQty = num;
          updated.variance = updated.actualQty - num;
        } else if (selectedCell.colLetter === 'G' && !Number.isNaN(num)) {
          updated.actualQty = num;
          updated.variance = num - updated.bookQty;
        } else if (selectedCell.colLetter === 'H' && !Number.isNaN(num)) {
          updated.variance = num;
        } else if (selectedCell.colLetter === 'C') {
          updated.itemName = input;
        }
        handleUpdateSingleRow(updated);
      }
    }
  };

  const handleAddFormulaColumn = (title: string, formulaTemplate: string) => {
    const prevActive = [...activeRows];
    const prevPurged = [...purgedRows];
    const prevCols = [...customColumns];

    const applied = applyStructuredOperations(
      activeRows,
      purgedRows,
      [{ type: 'ADD_FORMULA_COLUMN', columnTitle: title, formula: formulaTemplate }],
      customColumns
    );
    setActiveRows(applied.nextActiveRows);
    setCustomColumns(applied.nextCustomColumns);
    setActiveTab('grid');

    pushAuditLog(
      `إضافة عمود معادلة: ${title}`,
      applied.summaryParts[0] || `تم إضافة عمود ${title}`,
      activeRows.length,
      'manual',
      prevActive,
      prevPurged,
      prevCols
    );
  };

  const handleRemoveCustomColumn = (colLetter: string) => {
    setCustomColumns((prev) => prev.filter((c) => c.colLetter !== colLetter));
  };

  const handleImportUploadedRows = (uploadedRows: InventoryRow[], autoPurge: boolean) => {
    const prevActive = [...activeRows];
    const prevPurged = [...purgedRows];

    if (autoPurge) {
      const organized = organizeSheetPrecisely(uploadedRows, {
        removeZeroVariance: true,
        removeMatchedStatus: true,
        deduplicateSameCompanyItems: true,
        groupByCompany: true,
      }, customColumns);
      setActiveRows(organized.activeRows);
      setPurgedRows(organized.purgedRows);
      pushAuditLog(
        'رفع وتنظيم شيت إكسيل جديد',
        `تم استيراد ${uploadedRows.length} صنف، وحذف ${organized.purgedRows.length} صنف مطابق/صفري، وتنظيم ${organized.activeRows.length} صنف.`,
        uploadedRows.length,
        'manual',
        prevActive,
        prevPurged,
        customColumns
      );
    } else {
      const recalced = recalculateAllCustomColumns(uploadedRows, customColumns);
      setActiveRows(recalced);
      setPurgedRows([]);
      pushAuditLog(
        'رفع شيت إكسيل بدون حذف تلقائي',
        `تم استيراد ${uploadedRows.length} صنف بنجاح.`,
        uploadedRows.length,
        'manual',
        prevActive,
        prevPurged,
        customColumns
      );
    }
    setSelectedCompany('ALL');
    setSelectedStatus('ALL');
  };

  const handleResetDemoSheet = () => {
    const fresh = getInitialMultiCompanyDataset();
    setActiveRows(recalculateAllCustomColumns(fresh, customColumns));
    setPurgedRows([]);
    setSelectedCompany('ALL');
    setSelectedStatus('ALL');
    setSearchQuery('');
    setLastFeedback('تم إعادة تحميل الشيت التجريبي المجمع (4 شركات، يتضمن أصناف مطابقة وفروقات صفرية).');
  };

  return (
    <div id="top" className="min-h-screen flex flex-col bg-slate-50 text-slate-900">
      {/* Top Bar following the strict 3-zone contract */}
      <TopBar
        activeTab={activeTab}
        onSelectTab={setActiveTab}
        onExportExcel={() =>
          exportOrganizedWorkbook(
            visibleRows.length > 0 ? visibleRows : activeRows,
            purgedRows,
            customColumns
          )
        }
        purgedCount={purgedRows.length}
      />

      {/* Mobile Tab Switcher (only visible on small screens, non-sticky to respect 15% mobile sticky cap) */}
      <div className="md:hidden bg-white border-b border-slate-200 px-4 py-2 flex flex-wrap items-center gap-1.5">
        <button
          type="button"
          onClick={() => setActiveTab('grid')}
          className={`px-3 py-1.5 text-xs font-medium rounded-md whitespace-nowrap shrink-0 ${
            activeTab === 'grid' ? 'bg-teal-700 text-white' : 'text-slate-600 bg-slate-100'
          }`}
        >
          الجدول الموحد
        </button>
        <button
          type="button"
          onClick={() => setActiveTab('companies')}
          className={`px-3 py-1.5 text-xs font-medium rounded-md whitespace-nowrap shrink-0 ${
            activeTab === 'companies' ? 'bg-teal-700 text-white' : 'text-slate-600 bg-slate-100'
          }`}
        >
          تقسيم الشركات
        </button>
        <button
          type="button"
          onClick={() => setActiveTab('formulas')}
          className={`px-3 py-1.5 text-xs font-medium rounded-md whitespace-nowrap shrink-0 ${
            activeTab === 'formulas' ? 'bg-teal-700 text-white' : 'text-slate-600 bg-slate-100'
          }`}
        >
          بناء المعادلات
        </button>
        <button
          type="button"
          onClick={() => setActiveTab('purged')}
          className={`px-3 py-1.5 text-xs font-medium rounded-md whitespace-nowrap shrink-0 ${
            activeTab === 'purged' ? 'bg-teal-700 text-white' : 'text-slate-600 bg-slate-100'
          }`}
        >
          المستبعدات ({purgedRows.length})
        </button>
      </div>

      {/* Main Content Container */}
      <main className="flex-1 max-w-[1440px] w-full mx-auto px-4 sm:px-6 py-6 space-y-5">
        {/* Workspace Header & File Upload / Sheet Actions */}
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div>
            <h1 className="text-xl sm:text-2xl font-bold text-slate-900 font-display">
              منظم شيتات الإكسيل المتعدد الشركات ومحرك المعادلات الذكي
            </h1>
            <p className="text-xs sm:text-sm text-slate-600 mt-1">
              ارفع شيت إكسيل واحد لعدة شركات في نفس الصفحة — يقوم البرنامج بحذف المطابق والفارق (0) وترتيب الأصناف بسعرها وحالتها وتنفيذ أوامرك ومعادلاتك فوراً.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2.5 shrink-0">
            <button
              type="button"
              onClick={() => setModalMode('upload')}
              className="inline-flex items-center gap-2 px-4 py-2 text-xs font-semibold text-slate-900 bg-white border border-slate-300 rounded-lg hover:bg-slate-100 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
            >
              <Upload className="w-3.5 h-3.5 text-teal-700" />
              <span>رفع شيت إكسيل مجمع</span>
            </button>

            <button
              type="button"
              onClick={() => setModalMode('add-row')}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-slate-800 bg-white border border-slate-300 rounded-lg hover:bg-slate-100 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
            >
              <Plus className="w-3.5 h-3.5 text-teal-700" />
              <span>إضافة صنف</span>
            </button>

            <button
              type="button"
              onClick={handleResetDemoSheet}
              className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-medium text-slate-600 bg-slate-100 rounded-lg hover:bg-slate-200 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              <span>إعادة ضبط البيانات التجريبية</span>
            </button>
          </div>
        </div>

        {/* 4-Column KPI Summary Strip (Single-level elevation, tabular numerals) */}
        <section
          aria-label="ملخص إحصائيات الشيت والشركات"
          className="grid grid-cols-2 lg:grid-cols-4 gap-4"
        >
          <div className="bg-white border border-slate-200 rounded-xl p-4">
            <div className="text-xs text-slate-500">الشركات والأصناف النشطة</div>
            <div className="mt-1 flex items-baseline gap-2">
              <span className="text-2xl font-bold font-mono tabular-nums text-slate-900">
                {kpiStats.activeCount}
              </span>
              <span className="text-xs text-slate-500">
                صنف في {kpiStats.companiesCount} شركات
              </span>
            </div>
            <div className="mt-1 text-[11px] text-slate-500">
              متبقي {kpiStats.zeroOrMatchedRemaining} صنف مطابق (فارق 0) بالجدول
            </div>
          </div>

          <div className="bg-white border border-slate-200 rounded-xl p-4">
            <div className="text-xs text-slate-500">الأصناف المحذوفة (مطابق / فارق 0)</div>
            <div className="mt-1 flex items-baseline gap-2">
              <span className="text-2xl font-bold font-mono tabular-nums text-teal-700">
                {kpiStats.purgedCount}
              </span>
              <span className="text-xs text-slate-500">صنف مستبعد</span>
            </div>
            <div className="mt-1 text-[11px] text-slate-500">
              {kpiStats.zeroOrMatchedRemaining === 0
                ? 'تم تنظيف الشيت بالكامل من الفروقات الصفرية'
                : 'اضغط "حذف المطابق والفارق (0)" لتنظيفها'}
            </div>
          </div>

          <div className="bg-white border border-slate-200 rounded-xl p-4">
            <div className="text-xs text-slate-500">إجمالي قيمة العجز المالي</div>
            <div className="mt-1 flex items-baseline gap-1.5">
              <span className="text-2xl font-bold font-mono tabular-nums text-rose-700" dir="ltr">
                -{kpiStats.totalShortageVal.toLocaleString('en-US')}
              </span>
              <span className="text-xs text-slate-500">ج.م</span>
            </div>
            <div className="mt-1 text-[11px] font-mono text-slate-400" dir="ltr">
              =SUMIF(J:J, &quot;عجز&quot;, I:I)
            </div>
          </div>

          <div className="bg-white border border-slate-200 rounded-xl p-4">
            <div className="text-xs text-slate-500">إجمالي قيمة الزيادة وصافي الفارق</div>
            <div className="mt-1 flex items-baseline gap-1.5">
              <span className="text-2xl font-bold font-mono tabular-nums text-emerald-700" dir="ltr">
                +{kpiStats.totalSurplusVal.toLocaleString('en-US')}
              </span>
              <span className="text-xs text-slate-500">ج.م</span>
            </div>
            <div className="mt-1 text-[11px] text-slate-500 font-mono tabular-nums" dir="ltr">
              NET: {kpiStats.netVarianceVal.toLocaleString('en-US')}
            </div>
          </div>
        </section>

        {/* Smart Natural-Language Command Bar */}
        <SmartCommandBar
          onExecuteCommand={handleExecuteWrittenCommand}
          onQuickPurgeZero={handleQuickPurgeZero}
          onQuickFullOrganize={handleQuickFullOrganize}
          onUndoLast={handleUndoLast}
          canUndo={auditLogs.length > 0}
          isProcessing={isProcessingCommand}
          lastFeedback={lastFeedback}
          evaluatedBanner={evaluatedBanner}
        />

        {/* Excel Live Formula Bar (fx) */}
        <FormulaBar
          selectedCell={selectedCell}
          onApplyFormulaOrValue={handleApplyFormulaBarInput}
          onOpenFormulaBuilder={() => setActiveTab('formulas')}
        />

        {/* Workspace Active View */}
        {activeTab === 'grid' && (
          <div className="space-y-4">
            {/* Multi-Company & Status Filter Toolbar */}
            <div className="bg-white border border-slate-200 rounded-xl p-4 space-y-3">
              {/* Row 1: Company Segmented Filter Controls */}
              <div className="flex flex-wrap items-center gap-1.5">
                <span className="text-xs font-semibold text-slate-500 ml-1 whitespace-nowrap">
                  الشركة:
                </span>
                <button
                  type="button"
                  onClick={() => setSelectedCompany('ALL')}
                  className={`px-3 py-1.5 text-xs font-medium rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                    selectedCompany === 'ALL'
                      ? 'bg-slate-900 text-white'
                      : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
                  }`}
                >
                  كل الشركات ({activeRows.length})
                </button>
                {allCompanies.map((comp) => {
                  const count = activeRows.filter((r) => r.company === comp).length;
                  return (
                    <button
                      key={comp}
                      type="button"
                      onClick={() => setSelectedCompany(comp)}
                      className={`px-3 py-1.5 text-xs font-medium rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                        selectedCompany === comp
                          ? 'bg-teal-700 text-white'
                          : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
                      }`}
                    >
                      {comp} ({count})
                    </button>
                  );
                })}
              </div>

              {/* Row 2: Status Filter + Live Search */}
              <div className="pt-2.5 border-t border-slate-100 flex flex-wrap items-center justify-between gap-2.5">
                <div className="flex items-center gap-1 bg-slate-100 p-1 rounded-lg">
                  {(['ALL', 'عجز', 'زيادة', 'تالف', 'مرتجع', 'مطابق'] as const).map((st) => (
                    <button
                      key={st}
                      type="button"
                      onClick={() => setSelectedStatus(st)}
                      className={`px-2.5 py-1 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                        selectedStatus === st
                          ? 'bg-white text-slate-900 shadow-xs font-semibold'
                          : 'text-slate-600 hover:text-slate-900'
                      }`}
                    >
                      {st === 'ALL' ? 'كل الحالات' : st}
                    </button>
                  ))}
                </div>

                <div className="relative min-w-[200px]">
                  <Search className="w-3.5 h-3.5 text-slate-400 absolute right-3 top-1/2 -translate-y-1/2 pointer-events-none" />
                  <label htmlFor="grid-search-input" className="sr-only">
                    بحث باسم الصنف أو الكود
                  </label>
                  <input
                    id="grid-search-input"
                    type="text"
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    placeholder="بحث بالصنف أو الكود..."
                    className="w-full pr-8 pl-3 py-1.5 text-xs bg-slate-50 border border-slate-200 rounded-lg focus:outline-none focus:bg-white focus:border-teal-700"
                  />
                </div>
              </div>
            </div>

            {/* Main Interactive Excel Grid */}
            <SpreadsheetGrid
              rows={visibleRows}
              customColumns={customColumns}
              selectedCell={selectedCell}
              onSelectCell={setSelectedCell}
              onUpdateRow={handleUpdateSingleRow}
              onDeleteRow={handleDeleteSingleRow}
              onRemoveCustomColumn={handleRemoveCustomColumn}
              onSortColumn={handleSortColumn}
              sortField={sortField}
              sortDirection={sortDirection}
              onOpenAddRow={() => setModalMode('add-row')}
            />
          </div>
        )}

        {activeTab === 'companies' && (
          <CompanySplitView
            summaries={companySummaries}
            activeRows={activeRows}
            onFilterByCompany={(comp) => {
              setSelectedCompany(comp);
              setActiveTab('grid');
            }}
          />
        )}

        {activeTab === 'formulas' && (
          <FormulaStudioView
            activeRows={activeRows}
            customColumns={customColumns}
            onAddFormulaColumn={handleAddFormulaColumn}
            onRemoveCustomColumn={handleRemoveCustomColumn}
            onEvaluateFormulaInSheet={(formula, label) => {
              setActiveFormulaSpec({ formula, label });
              setActiveTab('grid');
            }}
          />
        )}

        {activeTab === 'purged' && (
          <section
            aria-label="سجل الأصناف المستبعدة والمطابقة (فارق 0)"
            className="bg-white border border-slate-200 rounded-xl p-5"
          >
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-slate-200">
              <div>
                <h2 className="text-base font-bold text-slate-900 font-display">
                  سجل الأصناف المحذوفة والمطابقة (الفارق = 0) ({purgedRows.length} صنف)
                </h2>
                <p className="text-xs text-slate-600 mt-0.5">
                  جميع الأصناف التي تم حذفها لأنها مطابقة أو لأن الفارق بين الدفتر والفعلي يساوي صفر محفوظة هنا للمراجعة أو الاسترجاع.
                </p>
              </div>
              {purgedRows.length > 0 && (
                <button
                  type="button"
                  onClick={() => handleExecuteWrittenCommand('استعادة المحذوف')}
                  className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold text-white bg-teal-700 rounded-lg hover:bg-teal-800 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                >
                  <RotateCcw className="w-3.5 h-3.5" />
                  <span>استعادة جميع الأصناف للجدول ({purgedRows.length})</span>
                </button>
              )}
            </div>

            {purgedRows.length === 0 ? (
              <div className="py-10 text-center">
                <CheckCircle2 className="w-8 h-8 text-slate-300 mx-auto mb-2" />
                <p className="text-sm font-semibold text-slate-700">
                  سجل المستبعدات فارغ حالياً
                </p>
                <p className="text-xs text-slate-500 mt-1">
                  اضغط على زر &quot;حذف المطابق والفارق (0)&quot; في الأعلى لنقل الأصناف الصفرية والمطابقة إلى هذا الأرشيف.
                </p>
              </div>
            ) : (
              <div className="overflow-x-auto mt-4">
                <table className="w-full text-right border-collapse text-xs">
                  <thead>
                    <tr className="bg-slate-50 border-b border-slate-200 text-slate-600">
                      <th className="py-2 px-3">الشركة</th>
                      <th className="py-2 px-3">الكود</th>
                      <th className="py-2 px-3">اسم الصنف</th>
                      <th className="py-2 px-3 text-left">سعر الوحدة</th>
                      <th className="py-2 px-3 text-left">الدفتر</th>
                      <th className="py-2 px-3 text-left">الفعلي</th>
                      <th className="py-2 px-3 text-left">الفارق</th>
                      <th className="py-2 px-3">الحالة</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {purgedRows.map((row) => (
                      <tr key={row.id} className="hover:bg-slate-50">
                        <td className="py-2 px-3 font-medium text-slate-800">{row.company}</td>
                        <td className="py-2 px-3 font-mono text-slate-500" dir="ltr">
                          {row.code}
                        </td>
                        <td className="py-2 px-3 font-semibold text-slate-900">{row.itemName}</td>
                        <td className="py-2 px-3 text-left font-mono tabular-nums" dir="ltr">
                          {row.unitPrice.toLocaleString('en-US')}
                        </td>
                        <td className="py-2 px-3 text-left font-mono tabular-nums" dir="ltr">
                          {row.bookQty}
                        </td>
                        <td className="py-2 px-3 text-left font-mono tabular-nums" dir="ltr">
                          {row.actualQty}
                        </td>
                        <td className="py-2 px-3 text-left font-mono tabular-nums" dir="ltr">
                          {row.variance}
                        </td>
                        <td className="py-2 px-3 text-slate-500">{row.status}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </section>
        )}

        {/* Operation & Command Audit History */}
        {auditLogs.length > 0 && (
          <section
            aria-label="سجل الأوامر والعمليات المنفذة"
            className="bg-white border border-slate-200 rounded-xl p-4"
          >
            <div className="flex items-center justify-between mb-3">
              <div className="flex items-center gap-2">
                <History className="w-4 h-4 text-slate-500" />
                <h3 className="text-xs font-bold text-slate-800">
                  سجل الأوامر الكتابية والعمليات المنفذة ({auditLogs.length})
                </h3>
              </div>
              <button
                type="button"
                onClick={handleUndoLast}
                className="text-xs font-semibold text-teal-700 hover:underline cursor-pointer"
              >
                تراجع عن آخر عملية
              </button>
            </div>
            <div className="space-y-2">
              {auditLogs.slice(0, 4).map((log) => (
                <div
                  key={log.id}
                  className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 text-xs py-2 px-3 bg-slate-50 rounded-lg border border-slate-100"
                >
                  <div className="flex items-center gap-2">
                    <span className="font-mono text-slate-400 tabular-nums">{log.timestamp}</span>
                    <span aria-hidden="true">·</span>
                    <span className="font-semibold text-slate-900">{log.commandText}</span>
                    <span aria-hidden="true">·</span>
                    <span className="text-slate-600">{log.summaryText}</span>
                  </div>
                  <span className="text-[11px] text-slate-400 whitespace-nowrap">
                    {log.source === 'gemini-ai'
                      ? 'عبر الذكاء الاصطناعي'
                      : log.source === 'local-smart'
                      ? 'محرك الأوامر الفوري'
                      : 'إجراء مباشر'}
                  </span>
                </div>
              ))}
            </div>
          </section>
        )}
      </main>

      {/* Clean Quiet Footer */}
      <footer className="bg-white border-t border-slate-200 py-4 px-6 mt-8">
        <div className="max-w-[1440px] mx-auto flex flex-col sm:flex-row items-center justify-between gap-2 text-xs text-slate-500">
          <div className="flex items-center gap-2">
            <FileSpreadsheet className="w-4 h-4 text-teal-700" />
            <span>ميزان إكسيل الذكي — لتنظيم شيتات الشركات وحذف المطابق وتنفيذ المعادلات والأوامر الكتابية</span>
          </div>
          <div>يدعم ملفات Excel (.xlsx / .xls / .csv) والتصدير متعدد الأوراق</div>
        </div>
      </footer>

      {/* Upload / Add Row Modal */}
      <UploadModal
        isOpen={modalMode !== null}
        mode={modalMode || 'upload'}
        onClose={() => setModalMode(null)}
        onImportRows={handleImportUploadedRows}
        onAddSingleRow={(newRow) => {
          const prevActive = [...activeRows];
          const prevPurged = [...purgedRows];
          const next = recalculateAllCustomColumns([newRow, ...activeRows], customColumns);
          setActiveRows(next);
          pushAuditLog(
            `إضافة صنف: ${newRow.itemName}`,
            `تمت إضافة "${newRow.itemName}" لشركة "${newRow.company}" وحساب معادلاته تلقائياً.`,
            1,
            'manual',
            prevActive,
            prevPurged,
            customColumns
          );
        }}
        existingCompanies={allCompanies}
      />
    </div>
  );
}
