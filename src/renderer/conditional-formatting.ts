import type { CellStyle, SheetDocument } from '../shared/types';
import { comparisonExpression, conditionalCellStyle, supportedConditionalRule } from '../shared/excel-features';
import { parseCellAddress } from './formulas';
import type { FormulaEvaluator } from './formulas';
import { shiftCopiedReferences } from './formula-references';

export function conditionalStyle(sheet: SheetDocument, row: number, column: number, evaluator: FormulaEvaluator): CellStyle {
  let result: CellStyle = {};
  const matches = (sheet.conditionalFormats ?? []).flatMap((group) => group.ref.split(/\s+/u).flatMap((range) => {
    const [start, last] = range.split(':');
    const first = parseCellAddress(start), end = parseCellAddress(last ?? start);
    if (!first || !end || row < first.row || row > end.row || column < first.column || column > end.column) return [];
    return group.rules.map((rule) => ({ rule, first }));
  })).sort((a, b) => a.rule.priority - b.rule.priority);
  for (const { rule, first } of matches) {
    if (!supportedConditionalRule(rule)) continue;
    let applies = false;
    if (rule.type === 'containsText') {
      const text = String(evaluator.evaluateCell(sheet.id, row, column) ?? '').toLowerCase();
      const expected = (rule.text ?? '').toLowerCase();
      applies = text.includes(expected);
    } else if (rule.type === 'expression') {
      const formula = shiftCopiedReferences(String(rule.formulae?.[0] ?? ''), row - first.row, column - first.column).replace(/^=/u, '');
      const match = comparisonExpression.exec(formula);
      if (match) {
        const address = parseCellAddress(match[1])!;
        const actual = evaluator.evaluateCell(sheet.id, address.row, address.column);
        const raw = match[3];
        const expected = raw.startsWith('"') ? raw.slice(1, -1).replace(/""/gu, '"') : /^(TRUE|FALSE)$/iu.test(raw) ? raw.toUpperCase() === 'TRUE' : Number(raw);
        const left = typeof expected === 'string' ? String(actual ?? '').toLowerCase() : actual;
        const right = typeof expected === 'string' ? expected.toLowerCase() : expected;
        if (left !== null && right !== null) {
          applies = match[2] === '=' ? left === right : match[2] === '<>' ? left !== right
            : match[2] === '<' ? left < right : match[2] === '>' ? left > right : match[2] === '<=' ? left <= right : left >= right;
        }
      }
    }
    if (applies) {
      result = { ...conditionalCellStyle(rule.style), ...result };
      if ((rule as typeof rule & { stopIfTrue?: boolean }).stopIfTrue) break;
    }
  }
  return result;
}
