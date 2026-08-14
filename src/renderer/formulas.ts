import type { CellData, CellScalar, SheetDocument, WorkbookDocument } from '../shared/types';
import { evaluateNamedFunction, FORMULA_FUNCTIONS, isFormulaError } from './formula-functions';
import type { FormulaAtom, FormulaError } from './formula-functions';

export type FormulaResult = FormulaAtom;
type ExpressionValue = FormulaResult | FormulaResult[];

interface Token {
  type: 'number' | 'word' | 'quoted' | 'operator' | 'eof';
  value: string;
}

export function cellKey(row: number, column: number): string {
  return `${row}:${column}`;
}

export function columnName(index: number): string {
  let value = index + 1;
  let result = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

export function columnIndex(name: string): number {
  return name.toUpperCase().split('').reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0) - 1;
}

export function parseCellAddress(address: string): { row: number; column: number } | null {
  const match = /^\$?([A-Z]+)\$?(\d+)$/iu.exec(address.trim());
  if (!match) return null;
  return { column: columnIndex(match[1]), row: Number(match[2]) - 1 };
}

export function addressForCell(row: number, column: number): string {
  return `${columnName(column)}${row + 1}`;
}

function isError(value: ExpressionValue): value is FormulaError {
  return isFormulaError(value);
}

function scalarNumber(value: ExpressionValue): number | FormulaError {
  if (isError(value)) return value;
  if (Array.isArray(value)) return '#VALUE!';
  if (typeof value === 'number') return value;
  if (typeof value === 'boolean') return value ? 1 : 0;
  if (value === null || value === '') return 0;
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : '#VALUE!';
}

function flatten(values: ExpressionValue[]): FormulaResult[] {
  return values.flatMap((value) => Array.isArray(value) ? value : [value]);
}

function tokenize(source: string): Token[] {
  const tokens: Token[] = [];
  let index = 0;
  while (index < source.length) {
    const char = source[index];
    if (/\s/u.test(char)) { index += 1; continue; }
    if (char === "'") {
      let value = '';
      index += 1;
      while (index < source.length) {
        if (source[index] === "'" && source[index + 1] === "'") { value += "'"; index += 2; continue; }
        if (source[index] === "'") { index += 1; break; }
        value += source[index];
        index += 1;
      }
      tokens.push({ type: 'quoted', value });
      continue;
    }
    const number = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/u.exec(source.slice(index));
    if (number) { tokens.push({ type: 'number', value: number[0] }); index += number[0].length; continue; }
    const cellReference = /^\$?[A-Z]+\$?\d+/iu.exec(source.slice(index));
    if (cellReference) { tokens.push({ type: 'word', value: cellReference[0] }); index += cellReference[0].length; continue; }
    const word = /^[A-Z_][A-Z0-9_.]*/iu.exec(source.slice(index));
    if (word) { tokens.push({ type: 'word', value: word[0] }); index += word[0].length; continue; }
    if ('+-*/^(),:!'.includes(char)) { tokens.push({ type: 'operator', value: char }); index += 1; continue; }
    tokens.push({ type: 'word', value: char });
    index += 1;
  }
  tokens.push({ type: 'eof', value: '' });
  return tokens;
}

export interface FormulaEvaluator {
  evaluateCell(sheetId: string, row: number, column: number): FormulaResult;
  evaluateFormula(sheetId: string, formula: string): FormulaResult;
}

export function createFormulaEvaluator(workbook: WorkbookDocument): FormulaEvaluator {
  const memo = new Map<string, FormulaResult>();
  const active = new Set<string>();

  const findSheet = (idOrName: string): SheetDocument | undefined =>
    workbook.sheets.find((sheet) => sheet.id === idOrName || sheet.name.toLocaleLowerCase() === idOrName.toLocaleLowerCase());

  const evaluateCell = (sheetId: string, row: number, column: number): FormulaResult => {
    const sheet = findSheet(sheetId);
    if (!sheet || row < 0 || column < 0) return '#REF!';
    const memoKey = `${sheet.id}:${row}:${column}`;
    if (memo.has(memoKey)) return memo.get(memoKey)!;
    if (active.has(memoKey)) return '#CIRC!';
    const cell = sheet.cells[cellKey(row, column)];
    if (!cell) return null;
    if (!cell.formula) return cell.value;
    const functions = [...cell.formula.matchAll(/\b([A-Z][A-Z0-9.]*)\s*\(/giu)].map((match) => match[1].toUpperCase());
    if (cell.cachedValue !== undefined && functions.some((name) => !FORMULA_FUNCTIONS[name])) return cell.cachedValue;
    active.add(memoKey);
    const result = evaluateFormula(sheet.id, cell.formula);
    active.delete(memoKey);
    const resolved = result === '#NAME?' && cell.cachedValue !== undefined ? cell.cachedValue : result;
    memo.set(memoKey, resolved);
    return resolved;
  };

  const evaluateFormula = (sheetId: string, formula: string): FormulaResult => {
    const tokens = tokenize(formula.replace(/^=/u, ''));
    let position = 0;
    const peek = (offset = 0) => tokens[position + offset] ?? tokens[tokens.length - 1];
    const consume = () => tokens[position++];
    const accept = (value: string) => peek().value === value ? (consume(), true) : false;

    const parseReference = (): ExpressionValue | null => {
      let referenceSheet = sheetId;
      let addressToken: Token;
      if ((peek().type === 'word' || peek().type === 'quoted') && peek(1).value === '!') {
        const qualifier = consume().value;
        consume();
        const target = findSheet(qualifier);
        if (!target) return '#REF!';
        referenceSheet = target.id;
        addressToken = consume();
      } else {
        addressToken = consume();
      }
      const start = parseCellAddress(addressToken.value);
      if (!start) { position -= 1; return null; }
      if (!accept(':')) return evaluateCell(referenceSheet, start.row, start.column);
      const endToken = consume();
      const end = parseCellAddress(endToken.value);
      if (!end) return '#REF!';
      const values: FormulaResult[] = [];
      const rowStart = Math.min(start.row, end.row);
      const rowEnd = Math.max(start.row, end.row);
      const columnStart = Math.min(start.column, end.column);
      const columnEnd = Math.max(start.column, end.column);
      for (let row = rowStart; row <= rowEnd; row += 1) {
        for (let column = columnStart; column <= columnEnd; column += 1) {
          values.push(evaluateCell(referenceSheet, row, column));
        }
      }
      return values;
    };

    const parseExpression = (): ExpressionValue => {
      let left = parseTerm();
      while (peek().value === '+' || peek().value === '-') {
        const operator = consume().value;
        const right = parseTerm();
        const leftNumber = scalarNumber(left);
        const rightNumber = scalarNumber(right);
        if (isError(leftNumber)) { left = leftNumber; continue; }
        if (isError(rightNumber)) { left = rightNumber; continue; }
        left = operator === '+' ? leftNumber + rightNumber : leftNumber - rightNumber;
      }
      return left;
    };

    const parseFunction = (name: string): ExpressionValue => {
      const args: ExpressionValue[] = [];
      if (!accept(')')) {
        do { args.push(parseExpression()); } while (accept(','));
        if (!accept(')')) return '#VALUE!';
      }
      return evaluateNamedFunction(name, flatten(args));
    };

    const parsePrimary = (): ExpressionValue => {
      if (accept('(')) {
        const value = parseExpression();
        return accept(')') ? value : '#VALUE!';
      }
      if (peek().type === 'number') return Number(consume().value);
      if (peek().type === 'word' || peek().type === 'quoted') {
        if (peek().type === 'word' && peek(1).value === '(' && !parseCellAddress(peek().value)) {
          const name = consume().value;
          consume();
          return parseFunction(name);
        }
        const reference = parseReference();
        if (reference !== null) return reference;
      }
      consume();
      return '#NAME?';
    };

    const parseUnary = (): ExpressionValue => {
      if (accept('+')) return parseUnary();
      if (accept('-')) {
        const value = scalarNumber(parseUnary());
        return isError(value) ? value : -value;
      }
      return parsePrimary();
    };

    const parsePower = (): ExpressionValue => {
      const left = parseUnary();
      if (!accept('^')) return left;
      const right = parsePower();
      const leftNumber = scalarNumber(left);
      const rightNumber = scalarNumber(right);
      if (isError(leftNumber)) return leftNumber;
      if (isError(rightNumber)) return rightNumber;
      return leftNumber ** rightNumber;
    };

    const parseTerm = (): ExpressionValue => {
      let left = parsePower();
      while (peek().value === '*' || peek().value === '/') {
        const operator = consume().value;
        const right = parsePower();
        const leftNumber = scalarNumber(left);
        const rightNumber = scalarNumber(right);
        if (isError(leftNumber)) { left = leftNumber; continue; }
        if (isError(rightNumber)) { left = rightNumber; continue; }
        if (operator === '/' && rightNumber === 0) { left = '#DIV/0!'; continue; }
        left = operator === '*' ? leftNumber * rightNumber : leftNumber / rightNumber;
      }
      return left;
    };

    const result = parseExpression();
    if (Array.isArray(result)) return '#VALUE!';
    if (peek().type !== 'eof') return '#VALUE!';
    return result;
  };

  return { evaluateCell, evaluateFormula };
}

export function editableCellText(cell: CellData | undefined): string {
  if (!cell) return '';
  if (cell.formula) return cell.formula;
  if (cell.value === null) return '';
  return String(cell.value);
}

export function normalizeCellInput(input: string): CellData | undefined {
  if (input === '') return undefined;
  if (input.startsWith('=')) return { value: null, formula: input, valueType: 'number' };
  const trimmed = input.trim();
  if (trimmed !== '' && Number.isFinite(Number(trimmed))) return { value: Number(trimmed), valueType: 'number' };
  if (/^(true|false)$/iu.test(trimmed)) return { value: trimmed.toLowerCase() === 'true', valueType: 'boolean' };
  return { value: input, valueType: 'text' };
}
