import { describe, expect, it } from 'vitest';
import { createBlankWorkbook } from '../src/shared/types';
import { cellShiftProblem, insertColumn, insertRow, shiftSelectedCells } from '../src/renderer/workbook-model';
import { rewriteCellShiftReferences } from '../src/renderer/formula-references';

const selection = { anchor: { row: 1, column: 0 }, focus: { row: 2, column: 0 } };
describe('selected cell shifts', () => {
  it('inserts a block without changing adjacent columns, styles, or unrelated references', () => {
    const workbook = createBlankWorkbook();
    const sheet = workbook.sheets[0];
    sheet.cells = { '1:0': { value: 7, style: { bold: true } }, '2:0': { value: 9 }, '1:1': { value: 12 }, '0:2': { value: null, formula: '=$A$2+A3+B2', cachedValue: 28 } };
    shiftSelectedCells(workbook, sheet.id, selection, 'row', 1);
    expect(sheet.cells['3:0']).toEqual({ value: 7, style: { bold: true } });
    expect(sheet.cells['4:0'].value).toBe(9);
    expect(sheet.cells['1:1'].value).toBe(12);
    expect(sheet.cells['0:2'].formula).toBe('=$A$4+A5+B2');
    expect(sheet.cells['0:2'].cachedValue).toBeUndefined();
  });
  it('deletes a block, shrinks contained ranges, and marks deleted references', () => {
    const workbook = createBlankWorkbook();
    const sheet = workbook.sheets[0];
    sheet.cells = { '1:0': { value: 7 }, '3:0': { value: 12 }, '0:1': { value: null, formula: '=SUM(A1:A5)+A2' } };
    shiftSelectedCells(workbook, sheet.id, selection, 'row', -1);
    expect(sheet.cells['1:0'].value).toBe(12);
    expect(sheet.cells['0:1'].formula).toBe('=SUM(A1:A3)+#REF!');
  });
  it.each([1, -1] as const)('shifts cells horizontally and updates cross-sheet formulas (%i)', (delta) => {
    const workbook = createBlankWorkbook();
    const sheet = workbook.sheets[0];
    const other = { ...structuredClone(sheet), id: 'other', name: 'Other', cells: { '0:0': { value: null, formula: `='${sheet.name}'!B2+'${sheet.name}'!C3` } } };
    workbook.sheets.push(other);
    sheet.cells['1:2'] = { value: 17 };
    shiftSelectedCells(workbook, sheet.id, { anchor: { row: 1, column: 1 }, focus: { row: 1, column: 1 } }, 'column', delta);
    expect(sheet.cells[delta === 1 ? '1:3' : '1:1'].value).toBe(17);
    expect(other.cells['0:0'].formula).toBe(delta === 1 ? `='${sheet.name}'!C2+'${sheet.name}'!C3` : `='${sheet.name}'!#REF!+'${sheet.name}'!C3`);
  });
  it('rejects partial merged cells and split formula ranges without changing the workbook', () => {
    const workbook = createBlankWorkbook();
    const sheet = workbook.sheets[0];
    sheet.merges = ['A2:B2'];
    const before = structuredClone(workbook);
    expect(cellShiftProblem(workbook, sheet.id, selection, 'row', 1)).toMatch(/Unmerge/);
    expect(() => shiftSelectedCells(workbook, sheet.id, selection, 'row', 1)).toThrow(/Unmerge/);
    expect(workbook).toEqual(before);
    sheet.merges = [];
    sheet.cells['0:2'] = { value: null, formula: '=SUM(A2:B5)' };
    expect(cellShiftProblem(workbook, sheet.id, selection, 'row', 1)).toMatch(/split/);
  });
  it('ignores strings and unrelated worksheets', () => {
    expect(rewriteCellShiftReferences('=A2+Other!A2+"A2"', 'Data', 'Data', { top: 1, bottom: 2, left: 0, right: 0, axis: 'row', delta: 1 })).toBe('=A4+Other!A2+"A2"');
  });
  it('marks shifted references beyond worksheet limits and rejects moving stored cells off the sheet', () => {
    const change = { top: 1, bottom: 2, left: 0, right: 0, axis: 'row' as const, delta: 1 as const };
    expect(rewriteCellShiftReferences('=A1048576', 'Data', 'Data', change)).toBe('=#REF!');
    const workbook = createBlankWorkbook();
    const sheet = workbook.sheets[0];
    sheet.cells['1048575:0'] = { value: 7 };
    expect(cellShiftProblem(workbook, sheet.id, selection, 'row', 1)).toMatch(/edge/);
  });
  it('supports inserting after the last row or column', () => {
    const workbook = createBlankWorkbook();
    const sheet = workbook.sheets[0];
    const rows = sheet.rowCount, columns = sheet.columnCount;
    expect(insertRow(workbook, sheet.id, rows)).toBe(true);
    expect(insertColumn(workbook, sheet.id, columns)).toBe(true);
    expect([sheet.rowCount, sheet.columnCount]).toEqual([rows + 1, columns + 1]);
  });
});
