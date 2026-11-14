import type { CellData, CellScalar } from './types';

function displayNumber(value: number, format?: string): string {
  if (!format || format === 'General') return Number.isInteger(value) ? String(value) : String(Number(value.toPrecision(12)));
  const primary = format.split(';')[0];
  const decimals = Math.min(6, (primary.split('.')[1]?.match(/[0#]/gu) ?? []).length);
  if (format.includes('%')) return new Intl.NumberFormat(undefined, { style: 'percent', minimumFractionDigits: decimals, maximumFractionDigits: decimals }).format(value);
  if (format.includes('$')) return new Intl.NumberFormat(undefined, { style: 'currency', currency: 'USD', currencySign: format.includes('(') ? 'accounting' : 'standard', minimumFractionDigits: decimals, maximumFractionDigits: decimals }).format(value);
  return new Intl.NumberFormat(undefined, { minimumFractionDigits: decimals, maximumFractionDigits: decimals }).format(value);
}

export function displayValue(cell: CellData | undefined, evaluated: CellScalar): string {
  const value = cell?.formula ? evaluated : cell?.value;
  if (value === null || value === undefined) return '';
  if (cell?.valueType === 'date' && typeof value === 'string') {
    const date = new Date(value);
    return Number.isNaN(date.valueOf()) ? value : date.toLocaleDateString(undefined, { timeZone: 'UTC' });
  }
  return typeof value === 'number' ? displayNumber(value, cell?.style?.numberFormat) : String(value);
}
