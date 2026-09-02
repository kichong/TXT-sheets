export type CellScalar = string | number | boolean | null;
export type CellValueType = 'blank' | 'text' | 'number' | 'boolean' | 'date' | 'error';
export type HorizontalAlignment = 'left' | 'center' | 'right';

export interface CellBorderSide {
  style?: 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double';
  color?: string;
}

export interface CellStyle {
  fontFamily?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  textColor?: string;
  fillColor?: string;
  horizontal?: HorizontalAlignment;
  vertical?: 'top' | 'middle' | 'bottom';
  wrapText?: boolean;
  numberFormat?: string;
  border?: {
    top?: CellBorderSide;
    right?: CellBorderSide;
    bottom?: CellBorderSide;
    left?: CellBorderSide;
  };
}

export interface CellData {
  value: CellScalar;
  valueType?: CellValueType;
  formula?: string;
  cachedValue?: CellScalar;
  hyperlink?: string;
  style?: CellStyle;
}

export interface SheetDocument {
  id: string;
  name: string;
  rowCount: number;
  columnCount: number;
  cells: Record<string, CellData>;
  merges: string[];
  columnWidths: Record<string, number>;
  rowHeights: Record<string, number>;
  frozenRows?: number;
  frozenColumns?: number;
  hiddenGridlines?: boolean;
}

export type WorkbookFormat = 'xlsx' | 'csv' | 'tsv';

export interface WorkbookSource {
  id: string;
  displayName: string;
  format: WorkbookFormat;
}

export interface CompatibilityIssue {
  feature: string;
  detail: string;
}

export interface CompatibilityReportRequest {
  sourceFormat: WorkbookFormat | 'unsaved';
  issues: CompatibilityIssue[];
}

export interface WorkbookDocument {
  schemaVersion: 1;
  title: string;
  activeSheetId: string;
  sheets: SheetDocument[];
  source?: WorkbookSource;
  compatibilityIssues: CompatibilityIssue[];
}

export interface OpenResult {
  workbook: WorkbookDocument;
  recentFiles: RecentFile[];
}

export interface SaveResult {
  status: 'saved' | 'canceled';
  source?: WorkbookSource;
  recentFiles: RecentFile[];
}

export interface RecentFile {
  id: string;
  displayName: string;
  format: WorkbookFormat;
  lastOpenedAt: string;
}

export type UpdatePhase =
  | 'idle'
  | 'checking'
  | 'available'
  | 'downloading'
  | 'downloaded'
  | 'up-to-date'
  | 'error'
  | 'unavailable';

export interface AppUpdateState {
  currentVersion: string;
  phase: UpdatePhase;
  canCheck: boolean;
  availableVersion?: string;
  downloadPercent?: number;
  lastCheckedAt?: string;
  message?: string;
}

export type AppCommand = 'new' | 'open' | 'save' | 'save-as' | 'undo' | 'redo' | 'find';

export interface SpreadsheetApi {
  open(): Promise<OpenResult | null>;
  openExternal(): Promise<OpenResult | null>;
  cancelExternal(): void;
  openRecent(id: string): Promise<OpenResult>;
  save(workbook: WorkbookDocument): Promise<SaveResult>;
  saveAs(workbook: WorkbookDocument): Promise<SaveResult>;
  reportCompatibility(request: CompatibilityReportRequest): Promise<void>;
  getRecentFiles(): Promise<RecentFile[]>;
  getRecovery(): Promise<WorkbookDocument | null>;
  writeRecovery(workbook: WorkbookDocument): Promise<void>;
  clearRecovery(): Promise<void>;
  getUpdateState(): Promise<AppUpdateState>;
  checkForUpdates(): Promise<AppUpdateState>;
  downloadUpdate(): Promise<AppUpdateState>;
  installUpdate(): Promise<void>;
  setDirty(dirty: boolean): void;
  onCommand(callback: (command: AppCommand) => void): () => void;
  onExternalFile(callback: () => void): () => void;
  onUpdateState(callback: (state: AppUpdateState) => void): () => void;
}

export const DEFAULT_CELL_STYLE: CellStyle = {
  fontFamily: 'Aptos',
  fontSize: 11,
  textColor: '#202124',
  horizontal: 'left',
  vertical: 'middle',
  numberFormat: 'General',
};

export function createBlankWorkbook(): WorkbookDocument {
  const id = crypto.randomUUID();
  return {
    schemaVersion: 1,
    title: 'Untitled workbook',
    activeSheetId: id,
    sheets: [{
      id,
      name: 'Sheet1',
      rowCount: 200,
      columnCount: 26,
      cells: {},
      merges: [],
      columnWidths: {},
      rowHeights: {},
    }],
    compatibilityIssues: [],
  };
}
