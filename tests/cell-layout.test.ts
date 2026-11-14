import { describe, expect, it } from 'vitest';
import { createBlankWorkbook } from '../src/shared/types';
import { mergeSelection, mergeProblem, unmergeSelection, resizeSelection, nextGridPoint } from '../src/renderer/cell-layout';
import { exportWorkbook, importWorkbook } from '../src/main/workbook-io';

const range = { anchor: { row: 1, column: 1 }, focus: { row: 3, column: 3 } };
describe('selected cell layout', () => {
  it('uniformly resizes selected dimensions, bounds the size, and leaves other dimensions alone', () => {
    const sheet = createBlankWorkbook().sheets[0];
    sheet.columnWidths = { '0': 88, '1': 99, '2': 120, '3': 200 };
    resizeSelection(sheet, range, 'column', 2, 180);
    expect(sheet.columnWidths).toEqual({ '0': 88, '1': 180, '2': 180, '3': 180 });
    resizeSelection(sheet, range, 'row', 2, 60);
    expect(sheet.rowHeights).toEqual({ '1': 60, '2': 60, '3': 60 });
    resizeSelection(sheet, range, 'row', 0, 1);
    expect(sheet.rowHeights['0']).toBe(20);
    expect(sheet.rowHeights['1']).toBe(60);
  });
  it('requires confirmation before loss, supports vertical/horizontal merges, navigation, unmerge, and XLSX reopen', async () => {
    const w = createBlankWorkbook(), sheet = w.sheets[0];
    sheet.cells['1:1'] = { value: 'Heading', style: { bold: true } };
    sheet.cells['2:1'] = { value: 'Other contents' };
    expect(mergeProblem(sheet, range).discardsContent).toBe(true);
    expect(() => mergeSelection(sheet, range)).toThrow(/Confirm/u);
    expect(sheet.cells['2:1'].value).toBe('Other contents');
    mergeSelection(sheet, range, true);
    expect(sheet.merges).toEqual(['B2:D4']);
    expect(sheet.cells['2:1']).toBeUndefined();
    expect(nextGridPoint(sheet, { row: 1, column: 1 }, 0, 1)).toEqual({ row: 1, column: 4 });
    expect(nextGridPoint(sheet, { row: 1, column: 4 }, 0, -1)).toEqual({ row: 1, column: 1 });
    expect(nextGridPoint(sheet, { row: 1, column: 1 }, 1, 0)).toEqual({ row: 4, column: 1 });
    resizeSelection(sheet, range, 'row', 2, 55);
    const copy = await importWorkbook(await exportWorkbook(w, 'layout.xlsx'), 'layout.xlsx', { id: 'test', displayName: 'layout.xlsx', format: 'xlsx' });
    expect(copy.sheets[0].merges).toEqual(['B2:D4']);
    expect(copy.sheets[0].cells['1:1'].value).toBe('Heading');
    expect(copy.sheets[0].rowHeights['2']).toBe(55);
    unmergeSelection(sheet, { anchor: { row: 1, column: 1 }, focus: { row: 1, column: 1 } });
    expect(sheet.merges).toEqual([]);
    mergeSelection(sheet, { anchor: { row: 5, column: 0 }, focus: { row: 5, column: 3 } });
    mergeSelection(sheet, { anchor: { row: 6, column: 0 }, focus: { row: 10, column: 0 } });
    expect(sheet.merges).toEqual(['A6:D6', 'A7:A11']);
  });
  it('rejects table intersections and partially selected existing merged cells', () => {
    const sheet = createBlankWorkbook().sheets[0];
    sheet.tables = [{ name: 'Table1', ref: 'B2', range: 'B2:D9', columns: [{ name: 'A' }, { name: 'B' }, { name: 'C' }] }];
    expect(mergeProblem(sheet, range).error).toMatch(/Excel table/u);
    sheet.tables = []; sheet.merges = ['A1:C3'];
    expect(mergeProblem(sheet, range).error).toMatch(/entire existing/u);
  });
});
