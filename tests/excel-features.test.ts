import { describe, expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { importWorkbook, exportWorkbook } from '../src/main/workbook-io';
import { conditionalStyle } from '../src/renderer/conditional-formatting';
import { createFormulaEvaluator } from '../src/renderer/formulas';
import { applyBorders, setCellInput, insertRow, insertColumn, deleteColumn, cellShiftProblem } from '../src/renderer/workbook-model';
import { listOptions } from '../src/shared/excel-features';

const source = { id: 'test', displayName: 'fixture.xlsx', format: 'xlsx' as const };

async function fixture() {
  const excel = new ExcelJS.Workbook();
  const sheet = excel.addWorksheet('Actions');
  sheet.addTable({ name: 'Actions', ref: 'B3', headerRow: true, totalsRow: false,
    columns: [{ name: 'Action' }, { name: 'Status' }], rows: [['First', 'Draft'], ['Second', 'Done']],
    style: { theme: 'TableStyleMedium2', showRowStripes: true } });
  sheet.getCell('C4').dataValidation = { type: 'list', formulae: ['"Draft,Done,Not filed,Filed"'] };
  sheet.getCell('C5').dataValidation = { type: 'list', formulae: ['"Draft,Done,Not filed,Filed"'] };
  sheet.addConditionalFormatting({ ref: 'C4:C5', rules: [
    { type: 'containsText', operator: 'containsText', text: 'Done', priority: 1, style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFDCFCE7' } }, font: { color: { argb: 'FF166534' }, bold: true } } },
    { type: 'expression', formulae: ['$C4="Not filed"'], priority: 2, style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE4E6' } } } },
  ] });
  return importWorkbook(new Uint8Array(await excel.xlsx.writeBuffer()), 'fixture.xlsx', source);
}

describe('native Excel features', () => {
  it('round-trips tables, inline validation, and conditional styles while keeping edits authoritative', async () => {
    let workbook = await fixture();
    expect(workbook.compatibilityIssues).toEqual([]);
    const selected = { anchor: { row: 3, column: 1 }, focus: { row: 4, column: 2 } };
    applyBorders(workbook.sheets[0], selected, 'all', { style: 'dashed', color: '#334455' });
    setCellInput(workbook.sheets[0], 3, 2, 'Done');
    for (let pass = 0; pass < 2; pass++) {
      const bytes = await exportWorkbook(workbook, 'saved.xlsx');
      const zip = await JSZip.loadAsync(bytes);
      expect(Object.keys(zip.files)).toContain('xl/tables/table1.xml');
      const native = new ExcelJS.Workbook();
      await native.xlsx.load(Buffer.from(bytes) as never);
      expect(native.worksheets[0].getTable('Actions').name).toBe('Actions');
      expect(native.worksheets[0].getCell('C4').value).toBe('Done');
      expect(native.worksheets[0].getCell('C4').dataValidation.formulae).toEqual(['"Draft,Done,Not filed,Filed"']);
      expect(native.worksheets[0].getCell('C4').border.bottom?.style).toBe('dashed');
      workbook = await importWorkbook(bytes, 'saved.xlsx', source);
      expect(workbook.compatibilityIssues).toEqual([]);
      expect(workbook.sheets[0].tables?.[0]).toMatchObject({ name: 'Actions', range: 'B3:C5', headerRow: true, style: { showRowStripes: true } });
      expect(workbook.sheets[0].conditionalFormats?.[0].rules).toHaveLength(2);
      expect(listOptions(workbook.sheets[0].cells['3:2'].validation)).toContain('Done');
    }
  });

  it('reacts to edits, relative expression references, priority, clear contents, and border presets', async () => {
    const w = await fixture(), sheet = w.sheets[0];
    const appearance = (row: number) => conditionalStyle(sheet, row, 2, createFormulaEvaluator(w));
    expect(appearance(3).fillColor).toBeUndefined();
    expect(appearance(4)).toMatchObject({ fillColor: '#DCFCE7', textColor: '#166534', bold: true });
    setCellInput(sheet, 4, 2, 'Not filed');
    expect(appearance(4).fillColor).toBe('#FFE4E6');
    setCellInput(sheet, 4, 2, 'Filed');
    expect(appearance(4).fillColor).toBeUndefined();
    setCellInput(sheet, 4, 2, '');
    expect(listOptions(sheet.cells['4:2'].validation)).toContain('Done');
    const range = { anchor: { row: 7, column: 1 }, focus: { row: 8, column: 2 } };
    applyBorders(sheet, range, 'outside', { style: 'thin', color: '#123456' });
    expect(sheet.cells['7:1'].style?.border).toEqual({ top: { style: 'thin', color: '#123456' }, left: { style: 'thin', color: '#123456' } });
    applyBorders(sheet, range, 'horizontal', { style: 'dotted' });
    expect(sheet.cells['7:1'].style?.border?.bottom?.style).toBe('dotted');
    applyBorders(sheet, range, 'none', {});
    expect(sheet.cells['7:1'].style?.border).toEqual({});
  });

  it('keeps table and rule ranges aligned through row/column edits and rejects unsafe partial shifts', async () => {
    const w = await fixture(), sheet = w.sheets[0];
    insertRow(w, sheet.id, 4);
    expect(sheet.tables?.[0].range).toBe('B3:C6');
    expect(sheet.conditionalFormats?.[0].ref).toBe('C4:C6');
    expect(listOptions(sheet.cells['4:2'].validation)).toContain('Done');
    insertColumn(w, sheet.id, 2);
    expect(sheet.tables?.[0].range).toBe('B3:D6');
    expect(sheet.tables?.[0].columns.map((column) => column.name)).toEqual(['Action', 'Column 1', 'Status']);
    expect(sheet.conditionalFormats?.[0].ref).toBe('D4:D6');
    expect(conditionalStyle(sheet, 5, 3, createFormulaEvaluator(w)).fillColor).toBe('#DCFCE7');
    expect(cellShiftProblem(w, sheet.id, { anchor: { row: 3, column: 1 }, focus: { row: 3, column: 1 } }, 'row', 1)).toMatch(/whole rows/u);
    deleteColumn(w, sheet.id, 2);
    expect(sheet.tables?.[0].range).toBe('B3:C6');
    const reopened = await importWorkbook(await exportWorkbook(w, 'edited.xlsx'), 'edited.xlsx', source);
    expect(reopened.sheets[0].tables?.[0].range).toBe('B3:C6');
    expect(reopened.sheets[0].cells['5:2'].value).toBe('Done');
  });

  it.skipIf(!process.env.TXT_SHEETS_WORKBOOK)('checks a supplied local workbook without retaining its content in fixtures', async () => {
    const path = process.env.TXT_SHEETS_WORKBOOK!;
    let workbook = await importWorkbook(await readFile(path), 'local.xlsx', source);
    expect(workbook.compatibilityIssues).toEqual([]);
    const expected = workbook.sheets.map((sheet) => ({ tables: sheet.tables, rules: sheet.conditionalFormats,
      validations: Object.entries(sheet.cells).filter(([, cell]) => cell.validation).map(([key, cell]) => [key, cell.validation]) }));
    for (let pass = 0; pass < 2; pass++) {
      workbook = await importWorkbook(await exportWorkbook(workbook, 'copy.xlsx'), 'copy.xlsx', source);
      expect(workbook.compatibilityIssues).toEqual([]);
      workbook.sheets.forEach((sheet, index) => {
        expect(sheet.tables).toEqual(expected[index].tables);
        expect(sheet.conditionalFormats).toEqual(expected[index].rules);
        expect(Object.entries(sheet.cells).filter(([, cell]) => cell.validation).map(([key, cell]) => [key, cell.validation])).toEqual(expected[index].validations);
      });
    }
  });
});
