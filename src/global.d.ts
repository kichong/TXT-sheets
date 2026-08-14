import type { SpreadsheetApi } from './shared/types';

declare global {
  interface Window { spreadsheet: SpreadsheetApi; }
}

export {};
