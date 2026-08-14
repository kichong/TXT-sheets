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
    return numbers.length ? Math.min(...numbers) : 0;
  },
  MAX: (values) => {
    const numbers = numericValues(values);
    return numbers.length ? Math.max(...numbers) : 0;
  },
  COUNT: (values) => numericValues(values).length,
});

export function evaluateNamedFunction(name: string, values: FormulaAtom[]): FormulaAtom {
  const error = values.find(isFormulaError);
  if (error) return error;
  return FORMULA_FUNCTIONS[name.toUpperCase()]?.(values) ?? '#NAME?';
}
