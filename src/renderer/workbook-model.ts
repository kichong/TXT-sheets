import type { CellBorderSide, CellData, CellStyle, SheetDocument, WorkbookDocument } from '../shared/types';
import { cellKey, columnName, normalizeCellInput, parseCellAddress } from './formulas';
import { rewriteCellShiftReferences, rewriteStructuralReferences, shiftCopiedReferences, shiftReferenceInterval, type SheetAxis } from './formula-references';

export interface CellPoint { row: number; column: number; }
export interface Selection { anchor: CellPoint; focus: CellPoint; kind?: 'cells' | 'rows' | 'columns'; }

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
  else if (existing?.note !== undefined || existing?.validation) sheet.cells[key] = { value: null, valueType: 'blank' };
  else delete sheet.cells[key];
  if (existing?.note !== undefined) sheet.cells[key].note = existing.note;
  if (existing?.validation) sheet.cells[key].validation = structuredClone(existing.validation);
}

export function applyBorders(sheet: SheetDocument, selection: Selection, preset: 'all' | 'outside' | 'horizontal' | 'vertical' | 'bottom' | 'none', side: CellBorderSide): void {
  const bounds = selectionBounds(selection);
  for (const { row, column, cell } of selectedCells(sheet, selection)) {
    const border = preset === 'none' ? {} : { ...cell?.style?.border };
    if (preset === 'all' || (preset === 'outside' && row === bounds.top) || (preset === 'horizontal' && row > bounds.top)) border.top = { ...side };
    if (preset === 'all' || (preset === 'outside' && row === bounds.bottom) || (preset === 'horizontal' && row < bounds.bottom) || (preset === 'bottom' && row === bounds.bottom)) border.bottom = { ...side };
    if (preset === 'all' || (preset === 'outside' && column === bounds.left) || (preset === 'vertical' && column > bounds.left)) border.left = { ...side };
    if (preset === 'all' || (preset === 'outside' && column === bounds.right) || (preset === 'vertical' && column < bounds.right)) border.right = { ...side };
    sheet.cells[cellKey(row, column)] = { value: null, ...cell, style: { ...cell?.style, border } };
  }
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
  if (delta === 1 && sheet[count] >= (axis === 'row' ? 1_048_576 : 16_384)) return false;
  if (!Number.isInteger(index) || index < 0 || index > sheet[count] || (delta === -1 && index === sheet[count]) || (delta === -1 && sheet[count] <= 1)) return false;
  sheet.tables = (sheet.tables ?? []).flatMap((table) => {
    const [start, end] = table.range.split(':').map(parseCellAddress);
    if (!start || !end) return [];
    const coordinates = shiftReferenceInterval(start[axis], end[axis], index, delta);
    if (!coordinates) return [];
    if (axis === 'column' && index >= start.column && index <= end.column) {
      if (delta === -1) table.columns.splice(index - start.column, 1);
      else if (index > start.column) {
        let suffix = 1;
        while (table.columns.some((column) => column.name === `Column ${suffix}`)) suffix++;
        table.columns.splice(index - start.column, 0, { name: `Column ${suffix}`, filterButton: true });
      }
    }
    if (axis === 'row' && delta === -1 && index === start.row) table.headerRow = false;
    if (axis === 'row' && delta === -1 && index === end.row) table.totalsRow = false;
    start[axis] = coordinates[0]; end[axis] = coordinates[1];
    table.range = `${columnName(start.column)}${start.row + 1}:${columnName(end.column)}${end.row + 1}`;
    table.ref = table.range.split(':')[0];
    return [table];
  });
  for (const owner of workbook.sheets) {
    owner.conditionalFormats = (owner.conditionalFormats ?? []).flatMap((group) => {
      if (owner.id === sheet.id) group.ref = group.ref.split(/\s+/u).map((range) => rewriteStructuralReferences(range, owner.name, sheet.name, axis, index, delta)).filter((range) => !range.includes('#REF!')).join(' ');
      if (!group.ref) return [];
      for (const rule of group.rules) if ('formulae' in rule && rule.formulae) {
        rule.formulae = rule.formulae.map((formula) => typeof formula === 'string' ? rewriteStructuralReferences(formula, owner.name, sheet.name, axis, index, delta) : formula);
      }
      return [group];
    });
    for (const cell of Object.values(owner.cells)) if (cell.validation?.formulae) {
      cell.validation.formulae = cell.validation.formulae.map((formula) => typeof formula === 'string' ? rewriteStructuralReferences(formula, owner.name, sheet.name, axis, index, delta) : formula);
    }
  }
  shiftCells(sheet, axis, index, delta);
  // A new row inside an inline-list range inherits its neighboring validation.
  if (delta === 1) {
    const limit = axis === 'row' ? sheet.columnCount : sheet.rowCount;
    for (let other = 0; other < limit; other++) {
      const row = axis === 'row' ? index : other, column = axis === 'column' ? index : other;
      const before = sheet.cells[cellKey(row - (axis === 'row' ? 1 : 0), column - (axis === 'column' ? 1 : 0))]?.validation;
      const after = sheet.cells[cellKey(row + (axis === 'row' ? 1 : 0), column + (axis === 'column' ? 1 : 0))]?.validation;
      if (before && after && JSON.stringify(before) === JSON.stringify(after)) {
        sheet.cells[cellKey(row, column)] = { value: null, validation: structuredClone(before) };
      }
    }
  }
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

export function cellShiftProblem(workbook: WorkbookDocument, sheetId: string, selection: Selection, axis: SheetAxis, delta: 1 | -1): string | null {
  const sheet = workbook.sheets.find((item) => item.id === sheetId)!;
  const bounds = selectionBounds(selection);
  const managedRanges = [...(sheet.tables ?? []).map((table) => table.range), ...(sheet.conditionalFormats ?? []).flatMap((group) => group.ref.split(/\s+/u))];
  for (const range of managedRanges) {
    const [startText, endText] = range.split(':');
    const start = parseCellAddress(startText), end = parseCellAddress(endText ?? startText);
    if (!start || !end) continue;
    const affected = axis === 'row' ? end.row >= bounds.top && end.column >= bounds.left && start.column <= bounds.right
      : end.column >= bounds.left && end.row >= bounds.top && start.row <= bounds.bottom;
    if (affected) return 'Insert or delete whole rows or columns to keep tables and conditional formatting intact.';
  }
  if (delta === 1) {
    const amount = axis === 'row' ? bounds.bottom - bounds.top + 1 : bounds.right - bounds.left + 1;
    const limit = axis === 'row' ? 1_048_576 : 16_384;
    for (const key of Object.keys(sheet.cells)) {
      const [row, column] = key.split(':').map(Number);
      const affected = axis === 'row' ? column >= bounds.left && column <= bounds.right && row >= bounds.top : row >= bounds.top && row <= bounds.bottom && column >= bounds.left;
      if (affected && (axis === 'row' ? row : column) + amount >= limit) return 'There is no room to shift these cells at the edge of the worksheet.';
    }
  }
  for (const range of sheet.merges) {
    const [start, end] = range.split(':').map(parseCellAddress);
    if (!start || !end) continue;
    const affected = axis === 'row' ? end.row >= bounds.top && end.column >= bounds.left && start.column <= bounds.right
      : end.column >= bounds.left && end.row >= bounds.top && start.row <= bounds.bottom;
    if (affected) return 'Unmerge cells in the affected rows or columns before shifting cells.';
  }
  try {
    for (const owner of workbook.sheets) for (const cell of Object.values(owner.cells)) {
      if (cell.formula) rewriteCellShiftReferences(cell.formula, owner.name, sheet.name, { ...bounds, axis, delta });
    }
  } catch (error) { return (error as Error).message; }
  return null;
}

export function shiftSelectedCells(workbook: WorkbookDocument, sheetId: string, selection: Selection, axis: SheetAxis, delta: 1 | -1): void {
  const problem = cellShiftProblem(workbook, sheetId, selection, axis, delta);
  if (problem) throw new Error(problem);
  const sheet = workbook.sheets.find((item) => item.id === sheetId)!;
  const bounds = selectionBounds(selection);
  const count = axis === 'row' ? bounds.bottom - bounds.top + 1 : bounds.right - bounds.left + 1;
  const index = axis === 'row' ? bounds.top : bounds.left;
  const next: Record<string, CellData> = {};
  for (const [key, cell] of Object.entries(sheet.cells)) {
    let [row, column] = key.split(':').map(Number);
    const inBand = axis === 'row' ? column >= bounds.left && column <= bounds.right : row >= bounds.top && row <= bounds.bottom;
    const coordinate = axis === 'row' ? row : column;
    if (inBand && coordinate >= index) {
      if (delta === -1 && coordinate < index + count) continue;
      if (axis === 'row') row += count * delta;
      else column += count * delta;
      if (delta === 1) {
        sheet.rowCount = Math.max(sheet.rowCount, row + 1);
        sheet.columnCount = Math.max(sheet.columnCount, column + 1);
      }
    }
    next[cellKey(row, column)] = cell;
  }
  sheet.cells = next;
  for (const owner of workbook.sheets) for (const cell of Object.values(owner.cells)) {
    if (!cell.formula) continue;
    const nextFormula = rewriteCellShiftReferences(cell.formula, owner.name, sheet.name, { ...bounds, axis, delta });
    if (nextFormula !== cell.formula) { cell.formula = nextFormula; delete cell.cachedValue; }
  }
}
