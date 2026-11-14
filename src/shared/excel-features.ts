import type { ConditionalFormattingRule, DataValidation, Style } from 'exceljs';
import type { CellStyle } from './types';

export function listOptions(validation?: DataValidation): string[] | null {
  if (validation?.type !== 'list' || validation.formulae?.length !== 1) return null;
  const formula = validation.formulae[0];
  if (typeof formula !== 'string' || !/^".*"$/su.test(formula)) return null;
  return formula.slice(1, -1).split(',');
}

export const comparisonExpression = /^\s*(\$?[A-Z]+\$?\d+)\s*(=|<>|<=|>=|<|>)\s*("(?:[^"]|"")*"|-?\d+(?:\.\d+)?|TRUE|FALSE)\s*$/iu;

export function supportedConditionalRule(rule: ConditionalFormattingRule): boolean {
  if (rule.type === 'containsText') return rule.operator === undefined || rule.operator === 'containsText';
  return rule.type === 'expression' && rule.formulae?.length === 1 && comparisonExpression.test(String(rule.formulae[0]).replace(/^=/u, ''));
}

export function conditionalCellStyle(style?: Partial<Style>): CellStyle {
  const color = (value?: { argb?: string }) => value?.argb ? `#${value.argb.slice(-6)}` : undefined;
  const fill = style?.fill;
  return Object.fromEntries(Object.entries({
    bold: style?.font?.bold, italic: style?.font?.italic,
    underline: style?.font?.underline === undefined ? undefined : Boolean(style.font.underline),
    textColor: color(style?.font?.color),
    fillColor: fill?.type === 'pattern' ? color(fill.fgColor ?? fill.bgColor) : undefined,
    fontFamily: style?.font?.name, fontSize: style?.font?.size,
    numberFormat: style?.numFmt,
  }).filter(([, value]) => value !== undefined));
}
