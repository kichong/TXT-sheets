import { describe, expect, it } from 'vitest';
import { createBlankWorkbook } from '../src/shared/types';
import { renameWorksheet } from '../src/renderer/sheet-actions';

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
