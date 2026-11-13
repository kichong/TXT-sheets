import type { CellScalar } from '../shared/types';

export type FormulaError = '#DIV/0!' | '#VALUE!' | '#NAME?' | '#REF!' | '#CIRC!';
export type FormulaAtom = CellScalar | FormulaError;
export type FormulaFunction = (values: FormulaAtom[]) => FormulaAtom;

export function isFormulaError(value: unknown): value is FormulaError {
  return typeof value === 'string' && value.startsWith('#');
}

function numericValues(values: FormulaAtom[]): number[] {
  return values.filter((value): value is number => typeof value === 'number' && Number.isFinite(value));
}

/**
 * The formula registry is the only place that named functions are wired in.
 * Future functions can live in separate modules and be added here without
 * changing tokenization, references, dependency evaluation, or the grid.
 */
export const FORMULA_FUNCTIONS: Readonly<Record<string, FormulaFunction>> = Object.freeze({
  SUM: (values) => numericValues(values).reduce((total, value) => total + value, 0),
  AVERAGE: (values) => {
    const numbers = numericValues(values);
    return numbers.length ? numbers.reduce((total, value) => total + value, 0) / numbers.length : '#DIV/0!';
  },
  MIN: (values) => {
    const numbers = numericValues(values);
    return numbers.length ? numbers.reduce((minimum, value) => Math.min(minimum, value), Infinity) : 0;
  },
  MAX: (values) => {
    const numbers = numericValues(values);
    return numbers.length ? numbers.reduce((maximum, value) => Math.max(maximum, value), -Infinity) : 0;
  },
  COUNT: (values) => numericValues(values).length,
});

export function evaluateNamedFunction(name: string, values: FormulaAtom[]): FormulaAtom {
  const error = values.find(isFormulaError);
  if (error) return error;
  return FORMULA_FUNCTIONS[name.toUpperCase()]?.(values) ?? '#NAME?';
}

export function isSupportedFunction(name: string): boolean {
  name = name.toUpperCase();
  return name === 'COUNTIFS' || name === 'HYPERLINK' || Boolean(FORMULA_FUNCTIONS[name]);
}

function matchesCriterion(value: FormulaAtom, criterion: FormulaAtom): boolean {
  if (typeof criterion !== 'string') return value === criterion;
  const match = /^(<=|>=|<>|=|<|>)(.*)$/u.exec(criterion);
  const operator = match?.[1] ?? '=';
  const expected = match?.[2] ?? criterion;
  const numeric = expected.trim() !== '' && Number.isFinite(Number(expected));
  let comparison: number;
  if (numeric) {
    if (typeof value !== 'number') return operator === '<>';
    comparison = value - Number(expected);
  } else {
    const actual = String(value ?? '');
    if ((operator === '=' || operator === '<>') && /[*?~]/u.test(expected)) {
      let pattern = '';
      for (let i = 0; i < expected.length; i++) {
        let char = expected[i];
        if (char === '~' && i + 1 < expected.length) char = expected[++i];
        else if (char === '*') { pattern += '.*'; continue; }
        else if (char === '?') { pattern += '.'; continue; }
        pattern += char.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&');
      }
      const matches = new RegExp(`^${pattern}$`, 'iu').test(actual);
      return operator === '=' ? matches : !matches;
    }
    comparison = actual.toLowerCase().localeCompare(expected.toLowerCase());
  }
  return operator === '=' ? comparison === 0 : operator === '<>' ? comparison !== 0
    : operator === '<' ? comparison < 0 : operator === '>' ? comparison > 0
      : operator === '<=' ? comparison <= 0 : comparison >= 0;
}

export function evaluateFunctionArguments(name: string, args: (FormulaAtom | FormulaAtom[])[]): FormulaAtom {
  name = name.toUpperCase();
  if (name === 'HYPERLINK') {
    if (args.length < 1 || args.length > 2 || args.some(Array.isArray)) return '#VALUE!';
    return args[1] as FormulaAtom ?? args[0] as FormulaAtom;
  }
  if (name !== 'COUNTIFS') return evaluateNamedFunction(name, args.flat());
  if (!args.length || args.length % 2) return '#VALUE!';
  const ranges: FormulaAtom[][] = [];
  const criteria: FormulaAtom[] = [];
  for (let i = 0; i < args.length; i += 2) {
    const range = args[i];
    const criterion = args[i + 1];
    if (!Array.isArray(range) || Array.isArray(criterion)) return '#VALUE!';
    if (ranges.length && range.length !== ranges[0].length) return '#VALUE!';
    if (isFormulaError(criterion)) return criterion;
    ranges.push(range); criteria.push(criterion);
  }
  return ranges[0].reduce<number>((count, _, index) => count + Number(ranges.every((range, i) => matchesCriterion(range[index], criteria[i]))), 0);
}
