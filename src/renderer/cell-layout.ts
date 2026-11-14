import type { SheetDocument } from '../shared/types';
import { addressForCell, cellKey, parseCellAddress } from './formulas';
import { resizeColumn, resizeRow, selectionBounds, type Selection, type CellPoint } from './workbook-model';

function rectangle(range: string) {
  const [first, last] = range.split(':');
  const start = parseCellAddress(first), end = parseCellAddress(last ?? first);
  return start && end ? { top: start.row, left: start.column, bottom: end.row, right: end.column } : null;
}

function overlaps(a: ReturnType<typeof selectionBounds>, b: ReturnType<typeof selectionBounds>) {
  return a.top <= b.bottom && a.bottom >= b.top && a.left <= b.right && a.right >= b.left;
}

export function resizedIndices(selection: Selection, axis: 'row' | 'column', index: number): number[] {
  const bounds = selectionBounds(selection);
  const first = axis === 'row' ? bounds.top : bounds.left, last = axis === 'row' ? bounds.bottom : bounds.right;
  return index >= first && index <= last ? Array.from({ length: last - first + 1 }, (_, offset) => first + offset) : [index];
}

export function resizeSelection(sheet: SheetDocument, selection: Selection, axis: 'row' | 'column', index: number, size: number): void {
  for (const target of resizedIndices(selection, axis, index)) {
    if (axis === 'row') resizeRow(sheet, target, size); else resizeColumn(sheet, target, size);
  }
}

export function mergeProblem(sheet: SheetDocument, selection: Selection): { error?: string; discardsContent: boolean } {
  const bounds = selectionBounds(selection);
  if (bounds.top === bounds.bottom && bounds.left === bounds.right) return { error: 'Select two or more cells to merge.', discardsContent: false };
  if ((bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1) > 100_000) return { error: 'Select a smaller range to merge.', discardsContent: false };
  for (const table of sheet.tables ?? []) {
    const area = rectangle(table.range);
    if (area && overlaps(bounds, area)) return { error: 'Cells inside an Excel table cannot be merged. Select cells outside the table.', discardsContent: false };
  }
  for (const range of sheet.merges) {
    const area = rectangle(range);
    if (area && overlaps(bounds, area) && (area.top < bounds.top || area.left < bounds.left || area.bottom > bounds.bottom || area.right > bounds.right)) {
      return { error: 'Select the entire existing merged cell before merging it with other cells.', discardsContent: false };
    }
  }
  const discardsContent = Object.entries(sheet.cells).some(([key, cell]) => {
    const [row, column] = key.split(':').map(Number);
    return row >= bounds.top && row <= bounds.bottom && column >= bounds.left && column <= bounds.right
      && (row !== bounds.top || column !== bounds.left)
      && (cell.formula || (cell.value !== null && cell.value !== '') || cell.note || cell.validation);
  });
  return { discardsContent };
}

export function mergeSelection(sheet: SheetDocument, selection: Selection, allowDiscard = false): void {
  const problem = mergeProblem(sheet, selection);
  if (problem.error) throw new Error(problem.error);
  if (problem.discardsContent && !allowDiscard) throw new Error('Merging keeps only the upper-left cell. Confirm before discarding other contents.');
  const bounds = selectionBounds(selection);
  sheet.merges = sheet.merges.filter((range) => { const area = rectangle(range); return !area || !overlaps(bounds, area); });
  for (const key of Object.keys(sheet.cells)) {
    const [row, column] = key.split(':').map(Number);
    if (row >= bounds.top && row <= bounds.bottom && column >= bounds.left && column <= bounds.right && (row !== bounds.top || column !== bounds.left)) delete sheet.cells[key];
  }
  sheet.cells[cellKey(bounds.top, bounds.left)] ??= { value: null, valueType: 'blank' };
  sheet.merges.push(`${addressForCell(bounds.top, bounds.left)}:${addressForCell(bounds.bottom, bounds.right)}`);
}

export function unmergeSelection(sheet: SheetDocument, selection: Selection): void {
  const bounds = selectionBounds(selection);
  sheet.merges = sheet.merges.filter((range) => { const area = rectangle(range); return !area || !overlaps(bounds, area); });
}

export function nextGridPoint(sheet: SheetDocument, point: CellPoint, rowDelta: number, columnDelta: number): CellPoint {
  let { row, column } = point;
  for (const range of sheet.merges) {
    const area = rectangle(range);
    if (area && row >= area.top && row <= area.bottom && column >= area.left && column <= area.right) {
      if (rowDelta > 0) row = area.bottom;
      if (columnDelta > 0) column = area.right;
      if (rowDelta < 0) row = area.top;
      if (columnDelta < 0) column = area.left;
      break;
    }
  }
  row = Math.max(0, Math.min(sheet.rowCount - 1, row + rowDelta));
  column = Math.max(0, Math.min(sheet.columnCount - 1, column + columnDelta));
  for (const range of sheet.merges) {
    const area = rectangle(range);
    if (area && row >= area.top && row <= area.bottom && column >= area.left && column <= area.right) return { row: area.top, column: area.left };
  }
  return { row, column };
}
