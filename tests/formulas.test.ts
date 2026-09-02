import { describe, expect, it } from 'vitest';
import { createBlankWorkbook } from '../src/shared/types';
import { cellKey, createFormulaEvaluator, normalizeCellInput } from '../src/renderer/formulas';
import { fillSelection } from '../src/renderer/workbook-model';
import { FORMULA_FUNCTIONS } from '../src/renderer/formula-functions';

describe('formula evaluator', () => {
  it('keeps named formulas in an explicit extensible registry', () => {
    expect(Object.keys(FORMULA_FUNCTIONS)).toEqual(['SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT']);
  });
  it('evaluates arithmetic with Excel precedence', () => {
    const workbook = createBlankWorkbook();
    const evaluator = createFormulaEvaluator(workbook);
    expect(evaluator.evaluateFormula(workbook.activeSheetId, '=2+3*4')).toBe(14);
    expect(evaluator.evaluateFormula(workbook.activeSheetId, '=(2+3)^2')).toBe(25);
  });

  it('supports subtraction, multiplication, and division with cell references', () => {
    const workbook = createBlankWorkbook();
    const sheet = workbook.sheets[0];
    sheet.cells[cellKey(0, 0)] = { value: 24, valueType: 'number' };
    sheet.cells[cellKey(0, 1)] = { value: 6, valueType: 'number' };
    const evaluator = createFormulaEvaluator(workbook);
    expect(evaluator.evaluateFormula(sheet.id, '=A1-B1')).toBe(18);
    expect(evaluator.evaluateFormula(sheet.id, '=A1*B1')).toBe(144);
    expect(evaluator.evaluateFormula(sheet.id, '=A1/B1')).toBe(4);
    expect(evaluator.evaluateFormula(sheet.id, '=$A$1/$B$1')).toBe(4);
    expect(evaluator.evaluateFormula(sheet.id, '= ( A1 ) - ( B1 )')).toBe(18);
  });

  it('treats referenced blank cells as zero in arithmetic', () => {
    const workbook = createBlankWorkbook();
    const sheet = workbook.sheets[0];
    sheet.cells[cellKey(0, 1)] = { value: 19.99, valueType: 'number' };
    expect(createFormulaEvaluator(workbook).evaluateFormula(sheet.id, '=A1-B1')).toBe(-19.99);
  });

  it('adjusts relative references when a formula is dragged to new rows', () => {
    const workbook = createBlankWorkbook();
    const sheet = workbook.sheets[0];
    sheet.cells[cellKey(28, 5)] = { value: 100, valueType: 'number' };
    sheet.cells[cellKey(28, 6)] = { value: 25, valueType: 'number' };
    sheet.cells[cellKey(29, 5)] = { value: 80, valueType: 'number' };
    sheet.cells[cellKey(29, 6)] = { value: 30, valueType: 'number' };
    sheet.cells[cellKey(28, 7)] = { value: null, formula: '=F29-G29', valueType: 'number', style: { numberFormat: '$#,##0.00' } };
    fillSelection(sheet, { anchor: { row: 28, column: 7 }, focus: { row: 28, column: 7 } }, { anchor: { row: 28, column: 7 }, focus: { row: 29, column: 7 } });

    expect(sheet.cells[cellKey(29, 7)].formula).toBe('=F30-G30');
    expect(sheet.cells[cellKey(29, 7)].style?.numberFormat).toBe('$#,##0.00');
    expect(createFormulaEvaluator(workbook).evaluateCell(sheet.id, 29, 7)).toBe(50);
  });

  it('preserves function names and absolute references during drag fill', () => {
    const workbook = createBlankWorkbook();
    const sheet = workbook.sheets[0];
    sheet.cells[cellKey(0, 2)] = { value: null, formula: '=LOG10(A1)+$B$2', cachedValue: 2, valueType: 'number' };
    fillSelection(sheet, { anchor: { row: 0, column: 2 }, focus: { row: 0, column: 2 } }, { anchor: { row: 0, column: 2 }, focus: { row: 0, column: 3 } });
    expect(sheet.cells[cellKey(0, 3)].formula).toBe('=LOG10(B1)+$B$2');
  });

  it('parses continued dates and formatted currency using the prior cell format', () => {
    expect(normalizeCellInput('8/24', { value: '2026-07-24', valueType: 'date', style: { numberFormat: 'm/d/yy' } })).toEqual({
      value: '2026-08-24', valueType: 'date',
    });
    expect(normalizeCellInput('$21.66', { value: 40.47, valueType: 'number', style: { numberFormat: '$#,##0.00' } })).toEqual({
      value: 21.66, valueType: 'number',
    });
  });

  it('shows the cached result for imported formulas outside the first registry', () => {
    const workbook = createBlankWorkbook();
    const sheet = workbook.sheets[0];
    sheet.cells[cellKey(0, 0)] = { value: null, formula: '=IF(B1>0,1,0)', cachedValue: 1 };
    expect(createFormulaEvaluator(workbook).evaluateCell(sheet.id, 0, 0)).toBe(1);
  });

  it('evaluates ranges and basic aggregate functions', () => {
    const workbook = createBlankWorkbook();
    const sheet = workbook.sheets[0];
    sheet.cells[cellKey(0, 0)] = { value: 10, valueType: 'number' };
    sheet.cells[cellKey(1, 0)] = { value: 20, valueType: 'number' };
    sheet.cells[cellKey(2, 0)] = { value: 30, valueType: 'number' };
    const evaluator = createFormulaEvaluator(workbook);
    expect(evaluator.evaluateFormula(sheet.id, '=SUM(A1:A3)')).toBe(60);
    expect(evaluator.evaluateFormula(sheet.id, '=AVERAGE(A1:A3)')).toBe(20);
    expect(evaluator.evaluateFormula(sheet.id, '=MIN(A1:A3)+MAX(A1:A3)')).toBe(40);
    expect(evaluator.evaluateFormula(sheet.id, '=COUNT(A1:A3)')).toBe(3);
  });

  it('supports cross-sheet references and detects circular references', () => {
    const workbook = createBlankWorkbook();
    const first = workbook.sheets[0];
    const second: typeof first = { ...first, id: 'second', name: 'Annual Plan', cells: {}, merges: [], columnWidths: {}, rowHeights: {} };
    workbook.sheets.push(second);
    second.cells[cellKey(0, 0)] = { value: 8, valueType: 'number' };
    first.cells[cellKey(0, 0)] = { value: null, formula: "='Annual Plan'!A1*2" };
    first.cells[cellKey(1, 0)] = { value: null, formula: '=A3' };
    first.cells[cellKey(2, 0)] = { value: null, formula: '=A2' };
    const evaluator = createFormulaEvaluator(workbook);
    expect(evaluator.evaluateCell(first.id, 0, 0)).toBe(16);
    expect(evaluator.evaluateCell(first.id, 1, 0)).toBe('#CIRC!');
  });

  it('returns spreadsheet errors safely', () => {
    const workbook = createBlankWorkbook();
    const evaluator = createFormulaEvaluator(workbook);
    expect(evaluator.evaluateFormula(workbook.activeSheetId, '=4/0')).toBe('#DIV/0!');
    expect(evaluator.evaluateFormula(workbook.activeSheetId, '=MYSTERY(1)')).toBe('#NAME?');
  });
});
