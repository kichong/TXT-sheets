import { describe, expect, it } from 'vitest';
import { createBlankWorkbook } from '../src/shared/types';
import { renameWorksheet, reorderWorksheet } from '../src/renderer/sheet-actions';

describe('worksheet reordering', () => {
  it('moves sheets to a drop boundary without changing the active sheet or contents', () => {
    const workbook = createBlankWorkbook();
    const first = workbook.sheets[0];
    first.cells['0:0'] = { value: 42, valueType: 'number' };
    workbook.sheets.push({ ...structuredClone(first), id: 'second', name: 'Second' });
    workbook.sheets.push({ ...structuredClone(first), id: 'third', name: 'Third' });

    expect(reorderWorksheet(workbook, first.id, 3)).toBe(true);
    expect(workbook.sheets.map((sheet) => sheet.id)).toEqual(['second', 'third', first.id]);
    expect(workbook.activeSheetId).toBe(first.id);
    expect(workbook.sheets[2].cells['0:0'].value).toBe(42);
    expect(reorderWorksheet(workbook, 'third', 0)).toBe(true);
    expect(workbook.sheets.map((sheet) => sheet.id)).toEqual(['third', 'second', first.id]);
  });

  it('ignores drops at the current position or with an invalid target', () => {
    const workbook = createBlankWorkbook();
    const id = workbook.sheets[0].id;
    expect(reorderWorksheet(workbook, id, 0)).toBe(false);
    expect(reorderWorksheet(workbook, id, 1)).toBe(false);
    expect(reorderWorksheet(workbook, id, 2)).toBe(false);
    expect(reorderWorksheet(workbook, 'missing', 0)).toBe(false);
  });
});

describe('worksheet renaming', () => {
  it('updates sheet references while preserving strings and external references', () => {
    const workbook = createBlankWorkbook();
    const sheet = workbook.sheets[0];
    sheet.cells['0:0'] = { value: null, valueType: 'blank', formula: '=Sheet1!A2+\'Sheet1\'!A3+"Sheet1!A4"+[Other.xlsx]Sheet1!A5' };
    expect(renameWorksheet(workbook, sheet.id, "Sam's budget")).toBeNull();
    expect(sheet.cells['0:0'].formula).toBe('=\'Sam\'\'s budget\'!A2+\'Sam\'\'s budget\'!A3+"Sheet1!A4"+[Other.xlsx]Sheet1!A5');
    expect(renameWorksheet(workbook, sheet.id, 'Budget')).toBeNull();
    expect(sheet.cells['0:0'].formula).toContain("'Budget'!A2");
  });

  it('rejects invalid and duplicate names without changing the workbook', () => {
    const workbook = createBlankWorkbook();
    const sheet = workbook.sheets[0];
    workbook.sheets.push({ ...structuredClone(sheet), id: 'other', name: 'Budget' });
    const before = structuredClone(workbook);
    for (const name of ['', 'x'.repeat(32), 'a/b', '[data]', "'name", 'budget']) {
      expect(renameWorksheet(workbook, sheet.id, name)).toBeTruthy();
      expect(workbook).toEqual(before);
    }
  });
});
