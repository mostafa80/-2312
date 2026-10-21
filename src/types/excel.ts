export type ItemStatus = 'عجز' | 'زيادة' | 'مطابق' | 'تالف' | 'مرتجع';

export interface CustomColumn {
  id: string;
  colLetter: string; // e.g. 'K', 'L', 'M'
  title: string;
  formulaTemplate: string; // e.g. '=E{row}*1.14' or '=G{row}*E{row}'
}

export interface InventoryRow {
  id: string;
  company: string; // Col A
  code: string; // Col B
  itemName: string; // Col C
  category: string; // Col D
  unitPrice: number; // Col E
  bookQty: number; // Col F
  actualQty: number; // Col G
  variance: number; // Col H (=G-F)
  varianceValue: number; // Col I (=H*E)
  status: ItemStatus; // Col J (=IF(H<0,"عجز",IF(H>0,"زيادة","مطابق")))
  notes?: string;
  customValues?: Record<string, number | string>;
  customFormulas?: Record<string, string>; // cell-specific custom formula override
}

export type SortField =
  | 'company'
  | 'itemName'
  | 'unitPrice'
  | 'bookQty'
  | 'actualQty'
  | 'variance'
  | 'absVariance'
  | 'varianceValue'
  | 'status';

export type SortDirection = 'asc' | 'desc';

export interface OrganizeConfig {
  removeZeroVariance: boolean;
  removeMatchedStatus: boolean;
  deduplicateSameCompanyItems: boolean;
  autoRecalculateStatus: boolean;
  primarySort: SortField;
  primaryDirection: SortDirection;
  secondarySort: SortField;
  secondaryDirection: SortDirection;
  groupByCompany: boolean;
}

export type OperationType =
  | 'PURGE_ZERO_AND_MATCHED'
  | 'ORGANIZE_PRECISE'
  | 'SORT_ROWS'
  | 'FILTER_KEEP_COMPANY'
  | 'DELETE_COMPANY'
  | 'DELETE_BY_STATUS'
  | 'KEEP_ONLY_STATUS'
  | 'DELETE_BY_PRICE'
  | 'KEEP_BY_PRICE'
  | 'ADJUST_PRICE_PERCENT'
  | 'DEDUPLICATE_ITEMS'
  | 'RECALCULATE_STATUSES'
  | 'RESTORE_PURGED'
  | 'ADD_FORMULA_COLUMN'
  | 'EVALUATE_FORMULA';

export interface StructuredOperation {
  type: OperationType;
  field?: SortField;
  direction?: SortDirection;
  companyName?: string;
  status?: ItemStatus;
  operator?: 'gt' | 'lt' | 'eq' | 'gte' | 'lte';
  value?: number;
  columnTitle?: string;
  formula?: string;
  explanation?: string;
}

export interface EvaluatedFormulaBanner {
  formula: string;
  label: string;
  result: string | number;
}

export interface AuditLogEntry {
  id: string;
  timestamp: string;
  commandText: string;
  summaryText: string;
  affectedCount: number;
  formulaUsed?: string;
  source: 'local-smart' | 'gemini-ai' | 'manual';
  previousActiveRows: InventoryRow[];
  previousPurgedRows: InventoryRow[];
  previousCustomColumns: CustomColumn[];
}

export interface CompanySummary {
  company: string;
  activeCount: number;
  purgedCount: number;
  shortageCount: number;
  surplusCount: number;
  damagedCount: number;
  returnedCount: number;
  totalShortageValue: number;
  totalSurplusValue: number;
  netVarianceValue: number;
  totalInventoryValue: number;
}

export interface SelectedCellInfo {
  rowIndex: number; // 1-based Excel row index
  rowId: string;
  colLetter: string; // 'A' .. 'J' or custom
  colLabel: string;
  rawValue: string | number;
  formula: string;
}
