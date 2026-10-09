import { describe, expect, it } from 'vitest';
import { createBlankWorkbook } from '../src/shared/types';
import { createFormulaEvaluator } from '../src/renderer/formulas';
import { rewriteStructuralReferences, shiftCopiedReferences } from '../src/renderer/formula-references';
import { deleteColumn, deleteRow, insertColumn, insertRow } from '../src/renderer/workbook-model';

describe('structural formula references', () => {
  it.each([
    ['=A1+A2+$A$2+A$2+$A2', 'row', 1, 1, '=A1+A3+$A$3+A$3+$A3'],
    ['=A1+B1+$B$1+B$1+$B1', 'column', 1, 1, '=A1+C1+$C$1+C$1+$C1'],
    ['=A1+A2+A3', 'row', 1, -1, '=A1+#REF!+A2'],
    ['=A1+B1+C1', 'column', 1, -1, '=A1+#REF!+B1'],
    ['=SUM(A1:A5)', 'row', 2, 1, '=SUM(A1:A6)'],
    ['=SUM(A1:A5)', 'row', 0, 1, '=SUM(A2:A6)'],
    ['=SUM(A1:A5)', 'row', 5, 1, '=SUM(A1:A5)'],
    ['=SUM(A1:A5)', 'row', 0, -1, '=SUM(A1:A4)'],
    ['=SUM(A1:A5)', 'row', 4, -1, '=SUM(A1:A4)'],
    ['=SUM(A2:A5)', 'row', 2, -1, '=SUM(A2:A4)'],
    ['=SUM(A2:A5)', 'row', 0, -1, '=SUM(A1:A4)'],
    ['=SUM(A5:A1)', 'row', 2, -1, '=SUM(A4:A1)'],
    ['=SUM(B2:D2)', 'row', 1, -1, '=SUM(#REF!)'],
    ['=SUM($B$2 : $D$5)', 'column', 1, -1, '=SUM($B$2 : $C$5)'],
    ['=SUM(A:C)+SUM(2:5)', 'column', 1, 1, '=SUM(A:D)+SUM(2:5)'],
    ['=SUM($A:$C)+SUM($2:$5)', 'row', 2, -1, '=SUM($A:$C)+SUM($2:$4)'],
    ['=A1048576', 'row', 1, 1, '=#REF!'],
    ['=XFD1', 'column', 1, 1, '=#REF!'],
  ] as const)('rewrites %s for a %s edit at %i (%i)', (formula, axis, index, delta, expected) => {
    expect(rewriteStructuralReferences(formula, 'Data', 'Data', axis, index, delta)).toBe(expected);
  });

  it('updates only the edited sheet and handles quoted names, escaped apostrophes, and qualified range endpoints', () => {
    const formula = "=A2+Data!$B$2+'Data'!C2+'Other'!D2+SUM(Data!A2:Data!A5)";
    expect(rewriteStructuralReferences(formula, 'Other', 'Data', 'row', 1, 1)).toBe(
      "=A2+Data!$B$3+'Data'!C3+'Other'!D2+SUM(Data!A3:Data!A6)",
    );
    expect(rewriteStructuralReferences("='Bob''s Plan'!A2+'Bob''s Plan'!A3", 'Other', "Bob's Plan", 'row', 1, -1)).toBe(
      "='Bob''s Plan'!#REF!+'Bob''s Plan'!A2",
    );
    expect(rewriteStructuralReferences('=данные!A2+DATA!B2', 'Other', 'Данные', 'row', 1, 1)).toBe('=данные!A3+DATA!B2');
    expect(rewriteStructuralReferences("=Data ! A2+'Data'! A3", 'Other', 'Data', 'row', 1, 1)).toBe("=Data ! A3+'Data'! A4");
  });

  it('does not change text, named functions, structured fields, external links, 3-D references, or partial names', () => {
    const formula = '=IF(A2="A2 ""B2""",LOG10(A2),Table1[[A2]:[B2]])+[Book.xlsx]Data!A2+\'Data:Other\'!A2+Data:Other!A2+_A2+A2_name+A20.thing';
    expect(rewriteStructuralReferences(formula, 'Data', 'Data', 'row', 1, 1)).toBe(
      '=IF(A3="A2 ""B2""",LOG10(A3),Table1[[A2]:[B2]])+[Book.xlsx]Data!A2+\'Data:Other\'!A2+Data:Other!A2+_A2+A2_name+A20.thing',
    );
    expect(shiftCopiedReferences('=LOG10(A1)+$B$2+IF(A1="A1",1,0)', 1, 1)).toBe('=LOG10(B2)+$B$2+IF(B2="A1",1,0)');
    expect(shiftCopiedReferences('=A1+$B$2', -1, 0)).toBe('=#REF!+$B$2');
  });

  it('moves cells and formulas throughout the workbook, invalidates changed caches, and keeps calculated values correct', () => {
    const workbook = createBlankWorkbook();
    const data = workbook.sheets[0];
    data.name = 'Data';
    data.cells = { '1:0': { value: 7 }, '2:0': { value: 11 }, '0:1': { value: null, formula: '=$A$2+A3', cachedValue: 18 } };
    const summary = { ...data, id: 'summary', name: 'Summary', cells: {
      '0:0': { value: null, formula: '=Data!A2+Data!A3', cachedValue: 18 },
      '0:1': { value: 4 },
      '0:2': { value: null, formula: '=B1*2', cachedValue: 8 },
    } };
    workbook.sheets.push(summary);
    expect(insertRow(workbook, data.id, 1)).toBe(true);
    expect(data.cells['2:0'].value).toBe(7);
    expect(data.cells['3:0'].value).toBe(11);
    expect(data.cells['0:1'].formula).toBe('=$A$3+A4');
    expect(data.cells['0:1'].cachedValue).toBeUndefined();
    expect(summary.cells['0:0'].formula).toBe('=Data!A3+Data!A4');
    expect(summary.cells['0:2'].formula).toBe('=B1*2');
    expect(summary.cells['0:2'].cachedValue).toBe(8);
    expect(createFormulaEvaluator(workbook).evaluateCell(summary.id, 0, 0)).toBe(18);
    deleteRow(workbook, data.id, 2);
    expect(summary.cells['0:0'].formula).toBe('=Data!#REF!+Data!A3');
    expect(createFormulaEvaluator(workbook).evaluateCell(summary.id, 0, 0)).toBe('#REF!');
    expect(createFormulaEvaluator(workbook).evaluateFormula(data.id, '=SUM(1,#REF!,3)')).toBe('#REF!');
  });

  it('moves, expands, and shrinks merged ranges, dimensions, and frozen panes along with cells', () => {
    const workbook = createBlankWorkbook();
    const sheet = workbook.sheets[0];
    sheet.merges = ['B2:D4', 'F6:G6'];
    sheet.rowHeights = { '1': 40, '4': 50 };
    sheet.columnWidths = { '1': 120, '4': 140 };
    sheet.frozenRows = 2;
    sheet.frozenColumns = 2;
    insertRow(workbook, sheet.id, 0);
    expect(sheet.merges).toEqual(['B3:D5', 'F7:G7']);
    expect(sheet.rowHeights).toEqual({ '2': 40, '5': 50 });
    expect(sheet.frozenRows).toBe(3);
    insertColumn(workbook, sheet.id, 2);
    expect(sheet.merges).toEqual(['B3:E5', 'G7:H7']);
    expect(sheet.columnWidths).toEqual({ '1': 120, '5': 140 });
    deleteRow(workbook, sheet.id, 2);
    expect(sheet.merges).toEqual(['B3:E4', 'G6:H6']);
    expect(sheet.frozenRows).toBe(2);
    deleteColumn(workbook, sheet.id, 1);
    expect(sheet.merges).toEqual(['B3:D4', 'F6:G6']);
    expect(sheet.columnWidths).toEqual({ '4': 140 });
    expect(sheet.frozenColumns).toBe(1);
    deleteRow(workbook, sheet.id, 5);
    expect(sheet.merges).toEqual(['B3:D4']);
  });

  it('rejects invalid edits and preserves the final row/column', () => {
    const workbook = createBlankWorkbook();
    const sheet = workbook.sheets[0];
    sheet.rowCount = 1;
    sheet.columnCount = 1;
    const original = structuredClone(workbook);
    expect(deleteRow(workbook, sheet.id, 0)).toBe(false);
    expect(deleteColumn(workbook, sheet.id, 0)).toBe(false);
    expect(insertRow(workbook, sheet.id, -1)).toBe(false);
    expect(insertColumn(workbook, sheet.id, 0.5)).toBe(false);
    expect(insertRow(workbook, 'missing', 0)).toBe(false);
    expect(workbook).toEqual(original);
  });
});
