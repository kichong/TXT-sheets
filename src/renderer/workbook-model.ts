import type { CellData, CellStyle, SheetDocument, WorkbookDocument } from '../shared/types';
import { cellKey, columnName, parseCellAddress } from './formulas';

export interface CellPoint { row: number; column: number; }
export interface Selection { anchor: CellPoint; focus: CellPoint; }

export const MIN_COLUMN_WIDTH = 40;
export const MAX_COLUMN_WIDTH = 640;
export const MIN_ROW_HEIGHT = 20;
export const MAX_ROW_HEIGHT = 240;

function boundedDimension(value: number, minimum: number, maximum: number): number {
  return Math.round(Math.min(maximum, Math.max(minimum, Number.isFinite(value) ? value : minimum)));
}

export function resizeColumn(sheet: SheetDocument, index: number, width: number): void {
  sheet.columnWidths[String(index)] = boundedDimension(width, MIN_COLUMN_WIDTH, MAX_COLUMN_WIDTH);
}

export function resizeRow(sheet: SheetDocument, index: number, height: number): void {
  sheet.rowHeights[String(index)] = boundedDimension(height, MIN_ROW_HEIGHT, MAX_ROW_HEIGHT);
}

export function selectionBounds(selection: Selection) {
  return {
    top: Math.min(selection.anchor.row, selection.focus.row),
    bottom: Math.max(selection.anchor.row, selection.focus.row),
    left: Math.min(selection.anchor.column, selection.focus.column),
    right: Math.max(selection.anchor.column, selection.focus.column),
  };
}

export function selectionLabel(selection: Selection): string {
  const bounds = selectionBounds(selection);
  const start = `${columnName(bounds.left)}${bounds.top + 1}`;
  const end = `${columnName(bounds.right)}${bounds.bottom + 1}`;
  return start === end ? start : `${start}:${end}`;
}

export function cloneWorkbook(workbook: WorkbookDocument): WorkbookDocument {
  return structuredClone(workbook);
}

export function selectedCells(sheet: SheetDocument, selection: Selection): Array<{ row: number; column: number; cell?: CellData }> {
  const bounds = selectionBounds(selection);
  const cells = [];
  for (let row = bounds.top; row <= bounds.bottom; row += 1) {
    for (let column = bounds.left; column <= bounds.right; column += 1) {
      cells.push({ row, column, cell: sheet.cells[cellKey(row, column)] });
    }
  }
  return cells;
}

export function applyStyle(sheet: SheetDocument, selection: Selection, patch: Partial<CellStyle>): void {
  for (const { row, column, cell } of selectedCells(sheet, selection)) {
    sheet.cells[cellKey(row, column)] = {
      value: cell?.value ?? null,
      valueType: cell?.valueType ?? 'blank',
      ...cell,
      style: { ...cell?.style, ...patch },
    };
  }
}

function shiftCells(sheet: SheetDocument, axis: 'row' | 'column', index: number, delta: 1 | -1): void {
  const next: Record<string, CellData> = {};
  for (const [key, cell] of Object.entries(sheet.cells)) {
    let [row, column] = key.split(':').map(Number);
    const coordinate = axis === 'row' ? row : column;
    if (delta === -1 && coordinate === index) continue;
    if (coordinate >= index) {
      if (axis === 'row') row += delta;
      else column += delta;
    }
    if (row >= 0 && column >= 0) next[cellKey(row, column)] = cell;
  }
  sheet.cells = next;
  const dimensions = axis === 'row' ? sheet.rowHeights : sheet.columnWidths;
  const shifted: Record<string, number> = {};
  for (const [key, value] of Object.entries(dimensions)) {
    let coordinate = Number(key);
    if (delta === -1 && coordinate === index) continue;
    if (coordinate >= index) coordinate += delta;
    if (coordinate >= 0) shifted[String(coordinate)] = value;
  }
  if (axis === 'row') sheet.rowHeights = shifted;
  else sheet.columnWidths = shifted;
  sheet.merges = sheet.merges.filter((range) => {
    const [start, end] = range.split(':').map(parseCellAddress);
    if (!start || !end) return false;
    const startCoordinate = axis === 'row' ? start.row : start.column;
    const endCoordinate = axis === 'row' ? end.row : end.column;
    return index < startCoordinate || index > endCoordinate;
  });
}

export function insertRow(sheet: SheetDocument, index: number): void {
  shiftCells(sheet, 'row', index, 1);
  sheet.rowCount += 1;
}

export function deleteRow(sheet: SheetDocument, index: number): void {
  if (sheet.rowCount <= 1) return;
  shiftCells(sheet, 'row', index, -1);
  sheet.rowCount -= 1;
}

export function insertColumn(sheet: SheetDocument, index: number): void {
  shiftCells(sheet, 'column', index, 1);
  sheet.columnCount += 1;
}

export function deleteColumn(sheet: SheetDocument, index: number): void {
  if (sheet.columnCount <= 1) return;
  shiftCells(sheet, 'column', index, -1);
  sheet.columnCount -= 1;
}

export function uniqueSheetName(workbook: WorkbookDocument, preferred = 'Sheet'): string {
  const names = new Set(workbook.sheets.map((sheet) => sheet.name.toLocaleLowerCase()));
  let index = 1;
  while (names.has(`${preferred}${index}`.toLocaleLowerCase())) index += 1;
  return `${preferred}${index}`;
}
