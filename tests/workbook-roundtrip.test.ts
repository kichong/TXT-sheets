import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { exportWorkbook, importWorkbook } from '../src/main/workbook-io';
import { createFormulaEvaluator } from '../src/renderer/formulas';
import { reorderWorksheet } from '../src/renderer/sheet-actions';
import { createBlankWorkbook } from '../src/shared/types';
import { insertRow, deleteColumn } from '../src/renderer/workbook-model';

const fixturePath = resolve('tests/fixtures/fidelity-fixture.xlsx');

describe('Excel workbook fidelity', () => {
  it('preserves structural reference edits, merged ranges, frozen panes, and deleted-reference errors across XLSX saves', async () => {
    const workbook = createBlankWorkbook();
    const data = workbook.sheets[0];
    data.name = 'Data';
    data.cells = { '1:0': { value: 7 }, '2:0': { value: 11 }, '0:2': { value: null, formula: '=$A$2+A3' } };
    data.merges = ['D2:F3'];
    data.frozenRows = 2;
    data.frozenColumns = 1;
    const summary = { ...createBlankWorkbook().sheets[0], name: 'Summary' };
    summary.cells = { '0:0': { value: null, formula: '=SUM(Data!A2:A3)' }, '0:1': { value: null, formula: '=Data!B2' } };
    workbook.sheets.push(summary);
    insertRow(workbook, data.id, 1);
    deleteColumn(workbook, data.id, 1);
    let reopened = workbook;
    for (let pass = 0; pass < 2; pass++) {
      const bytes = await exportWorkbook(reopened, 'structure.xlsx');
      const zip = await JSZip.loadAsync(bytes);
      expect(await zip.file('xl/workbook.xml')!.async('string')).toContain('fullCalcOnLoad="1"');
      reopened = await importWorkbook(bytes, 'structure.xlsx', { id: 'structure', displayName: 'structure.xlsx', format: 'xlsx' });
      const main = reopened.sheets.find((sheet) => sheet.name === 'Data')!;
      const other = reopened.sheets.find((sheet) => sheet.name === 'Summary')!;
      expect(main.cells['0:1'].formula).toBe('=$A$3+A4');
      expect(other.cells['0:0'].formula).toBe('=SUM(Data!A3:A4)');
      expect(other.cells['0:1'].formula).toBe('=Data!#REF!');
      expect(main.merges).toContain('C3:E4');
      expect(main.frozenRows).toBe(3);
      expect(main.frozenColumns).toBe(1);
      expect(createFormulaEvaluator(reopened).evaluateCell(other.id, 0, 0)).toBe(18);
      expect(createFormulaEvaluator(reopened).evaluateCell(other.id, 0, 1)).toBe('#REF!');
    }
  });
  it('exports CSV with many populated cells without exceeding the argument limit', async () => {
    const workbook = createBlankWorkbook();
    const sheet = workbook.sheets[0];
    for (let row = 0; row < 400; row++) {
      for (let column = 0; column < 400; column++) {
        sheet.cells[`${row}:${column}`] = { value: row * 400 + column, valueType: 'number' };
      }
    }
    const csv = new TextDecoder().decode(await exportWorkbook(workbook, 'large.csv'));
    const rows = csv.split('\r\n');
    expect(rows).toHaveLength(400);
    expect(rows[0].split(',')).toHaveLength(400);
    expect(rows[0].startsWith('0,1,2,')).toBe(true);
    expect(rows.at(-1)?.endsWith(',159999')).toBe(true);
  });
  it('imports formatting from blank continuation rows', async () => {
    const excel = new ExcelJS.Workbook();
    const sheet = excel.addWorksheet('Sheet1');
    sheet.getCell('A1').value = 'Date';
    sheet.getCell('A2').value = new Date('2026-01-02T00:00:00Z');
    sheet.getCell('A2').numFmt = 'm/d/yy';
    sheet.getCell('A3').numFmt = 'm/d/yy';
    const bytes = new Uint8Array(await excel.xlsx.writeBuffer());
    const imported = await importWorkbook(bytes, 'formatted.xlsx', { id: 'formatted', displayName: 'formatted.xlsx', format: 'xlsx' });
    expect(imported.sheets[0].cells['2:0']).toMatchObject({ value: null, valueType: 'blank', style: { numberFormat: 'm/d/yy' } });
  });

  it('imports and exports core workbook structure, formulas, and styling', async () => {
    const bytes = new Uint8Array(await readFile(fixturePath));
    const source = { id: 'fixture', displayName: 'fidelity-fixture.xlsx', format: 'xlsx' as const };
    const workbook = await importWorkbook(bytes, fixturePath, source);
    const overview = workbook.sheets.find((sheet) => sheet.name === 'Overview')!;

    expect(workbook.sheets.map((sheet) => sheet.name)).toEqual(['Overview', 'Inputs']);
    expect(overview.merges).toEqual(expect.arrayContaining(['A1:D1', 'A9:C9']));
    expect(overview.cells['0:0'].value).toBe('Office order');
    expect(overview.cells['0:0'].style?.fillColor).toBe('#1F7A59');
    expect(overview.cells['0:0'].style?.bold).toBe(true);
    expect(overview.cells['8:3'].formula).toBe('=SUM(D4:D7)');
    expect(overview.columnWidths['0']).toBeGreaterThan(100);

    const evaluator = createFormulaEvaluator(workbook);
    expect(evaluator.evaluateCell(overview.id, 3, 3)).toBe(54);
    expect(evaluator.evaluateCell(overview.id, 8, 3)).toBe(208);

    const directory = await mkdtemp(join(tmpdir(), 'txt-sheets-roundtrip-'));
    try {
      overview.frozenRows = 3;
      expect(reorderWorksheet(workbook, overview.id, workbook.sheets.length)).toBe(true);
      const output = join(directory, 'roundtrip.xlsx');
      await writeFile(output, await exportWorkbook(workbook, output));
      const reopened = await importWorkbook(new Uint8Array(await readFile(output)), output, source);
      expect(reopened.sheets.map((sheet) => sheet.name)).toEqual(['Inputs', 'Overview']);
      const reopenedOverview = reopened.sheets.find((sheet) => sheet.name === 'Overview')!;
      expect(reopenedOverview.merges).toEqual(expect.arrayContaining(['A1:D1', 'A9:C9']));
      expect(reopenedOverview.cells['8:3'].formula).toBe('=SUM(D4:D7)');
      expect(reopenedOverview.cells['0:0'].style?.fillColor).toBe('#1F7A59');
      expect(reopenedOverview.frozenRows).toBe(3);
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });
});
