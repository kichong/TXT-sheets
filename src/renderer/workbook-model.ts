import type { CellData, CellStyle, SheetDocument, WorkbookDocument } from '../shared/types';
import { cellKey, columnName, normalizeCellInput, parseCellAddress } from './formulas';
import { rewriteStructuralReferences, shiftCopiedReferences, shiftReferenceInterval, type SheetAxis } from './formula-references';

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

export function nearestCellTemplate(sheet: SheetDocument, row: number, column: number): CellData | undefined {
  for (let candidate = row - 1; candidate >= 0; candidate -= 1) {
    const cell = sheet.cells[cellKey(candidate, column)];
    if (cell?.style || cell?.valueType === 'date') return cell;
  }
  for (let candidate = row + 1; candidate < sheet.rowCount; candidate += 1) {
    const cell = sheet.cells[cellKey(candidate, column)];
    if (cell?.style || cell?.valueType === 'date') return cell;
  }
  return undefined;
}

export function setCellInput(sheet: SheetDocument, row: number, column: number, input: string): void {
  const key = cellKey(row, column);
  const existing = sheet.cells[key];
  const template = existing?.style ? existing : nearestCellTemplate(sheet, row, column) ?? existing;
  const normalized = normalizeCellInput(input, template);
  const inheritedStyle = existing?.style ?? template?.style;
  if (normalized) sheet.cells[key] = { ...normalized, style: inheritedStyle ? structuredClone(inheritedStyle) : undefined };
  else if (inheritedStyle) sheet.cells[key] = { value: null, valueType: 'blank', style: structuredClone(inheritedStyle) };
  else delete sheet.cells[key];
}

function modulo(value: number, divisor: number): number {
  return ((value % divisor) + divisor) % divisor;
}

export function fillSelection(sheet: SheetDocument, source: Selection, target: Selection): void {
  const sourceBounds = selectionBounds(source);
  const targetBounds = selectionBounds(target);
  const sourceHeight = sourceBounds.bottom - sourceBounds.top + 1;
  const sourceWidth = sourceBounds.right - sourceBounds.left + 1;
  const originals = new Map<string, CellData | undefined>();
  for (let row = sourceBounds.top; row <= sourceBounds.bottom; row += 1) {
    for (let column = sourceBounds.left; column <= sourceBounds.right; column += 1) {
      originals.set(cellKey(row, column), structuredClone(sheet.cells[cellKey(row, column)]));
    }
  }
  for (let row = targetBounds.top; row <= targetBounds.bottom; row += 1) {
    for (let column = targetBounds.left; column <= targetBounds.right; column += 1) {
      if (row >= sourceBounds.top && row <= sourceBounds.bottom && column >= sourceBounds.left && column <= sourceBounds.right) continue;
      const sourceRow = sourceBounds.top + modulo(row - sourceBounds.top, sourceHeight);
      const sourceColumn = sourceBounds.left + modulo(column - sourceBounds.left, sourceWidth);
      const sourceCell = originals.get(cellKey(sourceRow, sourceColumn));
      const targetKey = cellKey(row, column);
      if (!sourceCell) {
        delete sheet.cells[targetKey];
        continue;
      }
      const next = structuredClone(sourceCell);
      if (next.formula) {
        const copied = shiftCopiedReferences(next.formula, row - sourceRow, column - sourceColumn);
        if (copied !== next.formula) delete next.cachedValue;
        next.formula = copied;
      }
      sheet.cells[targetKey] = next;
    }
  }
}

function shiftCells(sheet: SheetDocument, axis: SheetAxis, index: number, delta: 1 | -1): void {
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
  sheet.merges = sheet.merges.flatMap((range) => {
    const [start, end] = range.split(':').map(parseCellAddress);
    if (!start || !end) return [];
    const coordinates = shiftReferenceInterval(start[axis], end[axis], index, delta);
    if (!coordinates) return [];
    start[axis] = coordinates[0];
    end[axis] = coordinates[1];
    if (start.row === end.row && start.column === end.column) return [];
    return [`${columnName(start.column)}${start.row + 1}:${columnName(end.column)}${end.row + 1}`];
  });
  const frozen = axis === 'row' ? 'frozenRows' : 'frozenColumns';
  if (sheet[frozen] && index < sheet[frozen]!) sheet[frozen] = Math.max(0, sheet[frozen]! + delta);
}

function changeStructure(workbook: WorkbookDocument, sheetId: string, axis: SheetAxis, index: number, delta: 1 | -1): boolean {
  const sheet = workbook.sheets.find((item) => item.id === sheetId);
  if (!sheet) return false;
  const count = axis === 'row' ? 'rowCount' : 'columnCount';
  if (!Number.isInteger(index) || index < 0 || index >= sheet[count] || (delta === -1 && sheet[count] <= 1)) return false;
  shiftCells(sheet, axis, index, delta);
  sheet[count] += delta;
  for (const owner of workbook.sheets) {
    for (const cell of Object.values(owner.cells)) {
      if (!cell.formula) continue;
      const next = rewriteStructuralReferences(cell.formula, owner.name, sheet.name, axis, index, delta);
      if (next !== cell.formula) {
        cell.formula = next;
        delete cell.cachedValue;
      }
    }
  }
  return true;
}

export function insertRow(workbook: WorkbookDocument, sheetId: string, index: number): boolean {
  return changeStructure(workbook, sheetId, 'row', index, 1);
}

export function deleteRow(workbook: WorkbookDocument, sheetId: string, index: number): boolean {
  return changeStructure(workbook, sheetId, 'row', index, -1);
}

export function insertColumn(workbook: WorkbookDocument, sheetId: string, index: number): boolean {
  return changeStructure(workbook, sheetId, 'column', index, 1);
}

export function deleteColumn(workbook: WorkbookDocument, sheetId: string, index: number): boolean {
  return changeStructure(workbook, sheetId, 'column', index, -1);
}

export function uniqueSheetName(workbook: WorkbookDocument, preferred = 'Sheet'): string {
  const names = new Set(workbook.sheets.map((sheet) => sheet.name.toLocaleLowerCase()));
  let index = 1;
  while (names.has(`${preferred}${index}`.toLocaleLowerCase())) index += 1;
  return `${preferred}${index}`;
}
