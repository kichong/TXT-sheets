import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { exportWorkbook, importWorkbook } from '../src/main/workbook-io';
import { createFormulaEvaluator } from '../src/renderer/formulas';

const fixturePath = resolve('outputs/019ff27f-24bb-7ae0-8017-7cbf746a0b6f/fidelity-fixture.xlsx');

describe('Excel workbook fidelity', () => {
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
      const output = join(directory, 'roundtrip.xlsx');
      await writeFile(output, await exportWorkbook(workbook, output));
      const reopened = await importWorkbook(new Uint8Array(await readFile(output)), output, source);
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
