import { columnIndex, columnName } from './formulas';

export type SheetAxis = 'row' | 'column';
interface Address {
  row?: number;
  column?: number;
  absoluteRow: boolean;
  absoluteColumn: boolean;
}

const SHEET_NAME = String.raw`(?:'(?:[^']|'')+'|(?:\[[^\]]+\])?[\p{L}_\\][\p{L}\p{N}_.\\]*)`;
const QUALIFIER = `${SHEET_NAME}(?::${SHEET_NAME})?`;
const ADDRESS = String.raw`(?:\$?[A-Z]{1,3}\$?[1-9]\d*|\$?[A-Z]{1,3}|\$?[1-9]\d*)`;

function parseAddress(text: string): Address | null {
  const cell = /^(\$?)([A-Z]{1,3})(\$?)([1-9]\d*)$/iu.exec(text);
  const column = /^(\$?)([A-Z]{1,3})$/iu.exec(text);
  const row = /^(\$?)([1-9]\d*)$/u.exec(text);
  const result: Address | null = cell
    ? { column: columnIndex(cell[2]), row: Number(cell[4]) - 1, absoluteColumn: !!cell[1], absoluteRow: !!cell[3] }
    : column ? { column: columnIndex(column[2]), absoluteColumn: !!column[1], absoluteRow: false }
      : row ? { row: Number(row[2]) - 1, absoluteRow: !!row[1], absoluteColumn: false } : null;
  if (!result || (result.column !== undefined && result.column > 16_383) || (result.row !== undefined && result.row > 1_048_575)) return null;
  return result;
}

function formatAddress(address: Address): string {
  return (address.column === undefined ? '' : `${address.absoluteColumn ? '$' : ''}${columnName(address.column)}`) +
    (address.row === undefined ? '' : `${address.absoluteRow ? '$' : ''}${address.row + 1}`);
}

function sheetName(prefix: string): string | undefined {
  if (!prefix) return undefined;
  const name = prefix.replace(/\s*!\s*$/u, '').trim();
  return name.startsWith("'") ? name.slice(1, -1).replace(/''/gu, "'") : name;
}

/** Change A1 references without interpreting string literals, table fields, or external/3-D links. */
function mapReferences(formula: string, transform: (first: Address, last: Address | undefined, sheet: string | undefined) => string | null): string {
  const reference = new RegExp(`(?<prefix>(?:${QUALIFIER}\\s*!\\s*)?)(?<first>${ADDRESS})(?:(?<join>\\s*:\\s*(?:${QUALIFIER}\\s*!\\s*)?)(?<last>${ADDRESS}))?`, 'iyu');
  let output = '';
  let index = 0;
  while (index < formula.length) {
    const char = formula[index];
    // Consume Excel's escaped string literals as one token.
    if (char === '"') {
      const start = index++;
      while (index < formula.length) {
        if (formula[index++] !== '"') continue;
        if (formula[index] === '"') { index++; continue; }
        break;
      }
      output += formula.slice(start, index);
      continue;
    }
    const previous = formula[index - 1] ?? '';
    reference.lastIndex = index;
    const match = !/[\p{L}\p{N}_.$!\]]/u.test(previous) ? reference.exec(formula) : null;
    if (match) {
      const next = formula.slice(reference.lastIndex);
      const groups = match.groups!;
      const first = parseAddress(groups.first);
      const last = groups.last ? parseAddress(groups.last) : undefined;
      const name = sheetName(groups.prefix);
      const join = groups.join ?? '';
      const lastPrefix = join.replace(/^\s*:\s*/u, '');
      const lastName = sheetName(lastPrefix) ?? name;
      const valid = first && (!groups.last || last) &&
        (!!last || (first.row !== undefined && first.column !== undefined)) &&
        (!last || ((first.row === undefined) === (last.row === undefined) && (first.column === undefined) === (last.column === undefined))) &&
        !/^[\p{L}\p{N}_.\[]|^\s*\(/u.test(next) &&
        !groups.prefix.includes('[') && !groups.prefix.includes(':') &&
        !lastPrefix.includes('[') && !lastPrefix.includes(':') &&
        name?.toLocaleLowerCase() === lastName?.toLocaleLowerCase();
      if (valid) {
        const changed = transform(first, last ?? undefined, name);
        output += changed === '#REF!' ? `${groups.prefix}#REF!` : changed === null ? match[0] :
          `${groups.prefix}${formatAddress(first)}${last ? `${join}${formatAddress(last)}` : ''}`;
      } else output += match[0];
      index = reference.lastIndex;
      continue;
    }
    // Structured references can contain cell-looking column labels and nested brackets.
    if (char === '[') {
      const start = index;
      let depth = 0;
      do {
        if (formula[index] === '[') depth++;
        else if (formula[index] === ']') depth--;
        index++;
      } while (index < formula.length && depth > 0);
      output += formula.slice(start, index);
      continue;
    }
    // A standalone quoted qualifier is not a local cell reference.
    if (char === "'") {
      const start = index++;
      while (index < formula.length) {
        if (formula[index++] !== "'") continue;
        if (formula[index] === "'") { index++; continue; }
        break;
      }
      output += formula.slice(start, index);
      continue;
    }
    output += char;
    index++;
  }
  return output;
}

/** Shift an interval on insert/delete, shrinking a range rather than breaking its endpoints. */
export function shiftReferenceInterval(first: number, last: number, index: number, delta: 1 | -1): [number, number] | null {
  const low = Math.min(first, last);
  const high = Math.max(first, last);
  if (delta === -1 && low === high && index === low) return null;
  const nextLow = low >= index && (delta === 1 || low > index) ? low + delta : low;
  const nextHigh = high >= index ? high + delta : high;
  return first <= last ? [nextLow, nextHigh] : [nextHigh, nextLow];
}

export function rewriteStructuralReferences(formula: string, ownerSheet: string, editedSheet: string, axis: SheetAxis, index: number, delta: 1 | -1): string {
  return mapReferences(formula, (first, last, qualifier) => {
    if ((qualifier ?? ownerSheet).toLocaleLowerCase() !== editedSheet.toLocaleLowerCase() || first[axis] === undefined) return null;
    const coordinates = shiftReferenceInterval(first[axis]!, (last ?? first)[axis]!, index, delta);
    if (!coordinates) return '#REF!';
    if (coordinates[0] === first[axis] && coordinates[1] === (last ?? first)[axis]) return null;
    first[axis] = coordinates[0];
    if (last) last[axis] = coordinates[1];
    if ((first.column ?? 0) > 16_383 || (last?.column ?? 0) > 16_383 || (first.row ?? 0) > 1_048_575 || (last?.row ?? 0) > 1_048_575) return '#REF!';
    return '';
  });
}

export function shiftCopiedReferences(formula: string, rowDelta: number, columnDelta: number): string {
  return mapReferences(formula, (first, last) => {
    let changed = false;
    for (const address of last ? [first, last] : [first]) {
      if (address.row !== undefined && !address.absoluteRow && rowDelta !== 0) { address.row += rowDelta; changed = true; }
      if (address.column !== undefined && !address.absoluteColumn && columnDelta !== 0) { address.column += columnDelta; changed = true; }
      if ((address.row ?? 0) < 0 || (address.column ?? 0) < 0 || (address.row ?? 0) > 1_048_575 || (address.column ?? 0) > 16_383) return '#REF!';
    }
    return changed ? '' : null;
  });
}
