import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import Papa from 'papaparse';
import { isSupportedFunction } from '../renderer/formula-functions';
import { addressForCell, parseCellAddress } from '../renderer/formulas';
import { listOptions, supportedConditionalRule } from '../shared/excel-features';
import { randomUUID } from 'node:crypto';
import { extname, parse, posix } from 'node:path';
import type {
  CellBorderSide,
  CellData,
  CellScalar,
  CellStyle,
  CompatibilityIssue,
  SheetDocument,
  WorkbookDocument,
  WorkbookFormat,
  WorkbookSource,
} from '../shared/types';

function formatForPath(path: string): WorkbookFormat {
  const extension = extname(path).toLowerCase();
  if (extension === '.csv') return 'csv';
  if (extension === '.tsv') return 'tsv';
  if (extension === '.xlsx') return 'xlsx';
  throw new Error('Choose an .xlsx, .csv, or .tsv file.');
}

function normalizeColor(color: Partial<ExcelJS.Color> | undefined): string | undefined {
  if (!color) return undefined;
  if ('argb' in color && typeof color.argb === 'string') return `#${color.argb.slice(-6)}`;
  return undefined;
}

function normalizeBorder(side: Partial<ExcelJS.Border> | undefined): CellBorderSide | undefined {
  if (!side?.style && !side?.color) return undefined;
  const supported = new Set(['thin', 'medium', 'thick', 'dashed', 'dotted', 'double']);
  return {
    style: supported.has(side.style ?? '') ? side.style as CellBorderSide['style'] : 'thin',
    color: normalizeColor(side.color),
  };
}

function importStyle(cell: ExcelJS.Cell): CellStyle | undefined {
  const font = cell.font ?? {};
  const fill = cell.fill as ExcelJS.FillPattern | undefined;
  const alignment = cell.alignment ?? {};
  const border = cell.border ?? {};
  const style: CellStyle = {
    fontFamily: font.name,
    fontSize: font.size,
    bold: font.bold,
    italic: font.italic,
    underline: Boolean(font.underline),
    textColor: normalizeColor(font.color),
    fillColor: fill?.type === 'pattern' ? normalizeColor(fill.fgColor) : undefined,
    horizontal: alignment.horizontal === 'center' || alignment.horizontal === 'right' ? alignment.horizontal : alignment.horizontal === 'left' ? 'left' : undefined,
    vertical: alignment.vertical === 'top' || alignment.vertical === 'bottom' ? alignment.vertical : alignment.vertical === 'middle' ? 'middle' : undefined,
    wrapText: alignment.wrapText,
    numberFormat: cell.numFmt && cell.numFmt !== 'General' ? cell.numFmt : undefined,
    border: {
      top: normalizeBorder(border.top), right: normalizeBorder(border.right),
      bottom: normalizeBorder(border.bottom), left: normalizeBorder(border.left),
    },
  };
  const hasValue = Object.entries(style).some(([key, value]) => key === 'border'
    ? Object.values(value as object).some(Boolean)
    : value !== undefined && value !== false);
  return hasValue ? style : undefined;
}

function scalarFromExcel(value: ExcelJS.CellValue): { value: CellScalar; valueType?: CellData['valueType']; hyperlink?: string } {
  if (value === null || value === undefined) return { value: null, valueType: 'blank' };
  if (typeof value === 'string') return { value, valueType: 'text' };
  if (typeof value === 'number') return { value, valueType: 'number' };
  if (typeof value === 'boolean') return { value, valueType: 'boolean' };
  if (value instanceof Date) return { value: value.toISOString().slice(0, 10), valueType: 'date' };
  if ('richText' in value) return { value: value.richText.map((part) => part.text).join(''), valueType: 'text' };
  if ('hyperlink' in value) return { value: value.text, valueType: 'text', hyperlink: value.hyperlink };
  if ('error' in value) return { value: value.error, valueType: 'error' };
  return { value: String(value), valueType: 'text' };
}

function formulaCell(cell: ExcelJS.Cell): CellData | null {
  const formula = cell.formula;
  if (!formula) return null;
  const raw = cell.value as ExcelJS.CellFormulaValue;
  const cached = raw && typeof raw === 'object' && 'result' in raw ? scalarFromExcel(raw.result ?? null).value : null;
  return { value: null, valueType: typeof cached === 'number' ? 'number' : 'text', formula: `=${formula}`, cachedValue: cached, style: importStyle(cell) };
}

function compatibilityIssuesFor(worksheet: ExcelJS.Worksheet): CompatibilityIssue[] {
  const issues: CompatibilityIssue[] = [];
  if (worksheet.getImages().length) issues.push({ feature: 'Images', detail: `${worksheet.name} contains images. They remain in the source file, but this editor does not display or preserve them on save.` });
  const model = worksheet.model as ExcelJS.WorksheetModel & { tables?: unknown[]; conditionalFormattings?: unknown[]; dataValidations?: { model?: object } };
  if ((model.conditionalFormattings as ExcelJS.ConditionalFormattingOptions[] | undefined)?.some((group) => group.rules.some((rule) => !supportedConditionalRule(rule)))) issues.push({ feature: 'Conditional formatting', detail: `${worksheet.name} contains advanced conditional rules that are preserved on save but not evaluated here.` });
  if (Object.values(model.dataValidations ?? {}).some((rule) => listOptions(rule as ExcelJS.DataValidation) === null)) issues.push({ feature: 'Data validation', detail: `${worksheet.name} contains advanced validation rules that are preserved on save but not enforced here.` });
  if (worksheet.state !== 'visible') issues.push({ feature: 'Hidden sheets', detail: `${worksheet.name} was hidden and is visible while editing here.` });
  const unsupportedFunctions = new Set<string>();
  worksheet.eachRow({ includeEmpty: false }, (row) => row.eachCell({ includeEmpty: false }, (cell) => {
    if (!cell.formula) return;
    for (const match of cell.formula.matchAll(/\b([A-Z][A-Z0-9.]*)\s*\(/giu)) {
      const name = match[1].toUpperCase();
      if (!isSupportedFunction(name)) unsupportedFunctions.add(name);
    }
  }));
  if (unsupportedFunctions.size) issues.push({
    feature: 'Advanced formulas',
    detail: `${worksheet.name} uses ${[...unsupportedFunctions].sort().join(', ')}. Cached results are displayed, but these functions do not recalculate in this first release.`,
  });
  return issues;
}

function importWorksheet(worksheet: ExcelJS.Worksheet): SheetDocument {
  const cells: Record<string, CellData> = {};
  worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    row.eachCell({ includeEmpty: true }, (cell, columnNumber) => {
      if (cell.type === ExcelJS.ValueType.Merge) return;
      const style = importStyle(cell);
      if ((cell.value === null || cell.value === undefined) && !style && !cell.note) return;
      const imported = formulaCell(cell) ?? { ...scalarFromExcel(cell.value), style };
      if (cell.note) imported.note = typeof cell.note === 'string' ? cell.note : (cell.note.texts ?? []).map((part) => part.text).join('');
      if (cell.dataValidation?.type) imported.validation = structuredClone(cell.dataValidation);
      cells[`${rowNumber - 1}:${columnNumber - 1}`] = imported;
    });
  });
  const columnWidths: Record<string, number> = {};
  for (let column = 1; column <= Math.max(worksheet.columnCount, 26); column += 1) {
    const width = worksheet.getColumn(column).width;
    if (width) columnWidths[String(column - 1)] = Math.max(34, Math.round(width * 7.2));
  }
  const rowHeights: Record<string, number> = {};
  worksheet.eachRow({ includeEmpty: true }, (row, number) => {
    if (row.height) rowHeights[String(number - 1)] = Math.max(20, Math.round(row.height * 1.34));
  });
  const model = worksheet.model as ExcelJS.WorksheetModel & { merges?: string[]; dataValidations?: Record<string, ExcelJS.DataValidation>; tables?: unknown[]; conditionalFormattings?: ExcelJS.ConditionalFormattingOptions[] };
  for (const [address, validation] of Object.entries(model.dataValidations ?? {})) {
    const point = parseCellAddress(address);
    if (!point) continue;
    const key = `${point.row}:${point.column}`;
    cells[key] ??= { value: null, valueType: 'blank' };
    cells[key].validation = structuredClone(validation as ExcelJS.DataValidation);
  }
  const tables = (model.tables ?? []).map((raw) => {
    const table = raw as unknown as ExcelJS.TableProperties & { tableRef: string };
    return { name: table.name, displayName: table.displayName, range: table.tableRef, ref: table.tableRef.split(':')[0],
      headerRow: table.headerRow, totalsRow: table.totalsRow,
      columns: structuredClone(table.columns).map((column, index) => ({ ...column, filterButton: column.filterButton ?? true,
        totalsRowLabel: index === 0 ? column.totalsRowLabel ?? 'Total' : column.totalsRowLabel,
        totalsRowFunction: index > 0 ? column.totalsRowFunction ?? 'none' : column.totalsRowFunction })),
      style: structuredClone(table.style) };
  });
  const conditionalFormats = structuredClone(model.conditionalFormattings ?? []);
  for (const group of conditionalFormats) for (const rule of group.rules) {
    if (rule.type === 'containsText' && rule.text !== undefined && !('formulae' in rule)) {
      (rule as typeof rule & { formulae: string[] }).formulae = [`NOT(ISERROR(SEARCH("${rule.text.replace(/"/gu, '""')}",${group.ref.split(/[ :]/u)[0]})))`];
    }
  }
  const views = Array.isArray(worksheet.views) ? worksheet.views[0] : undefined;
  return {
    id: randomUUID(),
    name: worksheet.name,
    rowCount: Math.max(200, worksheet.rowCount + 20),
    columnCount: Math.max(26, worksheet.columnCount + 5),
    cells,
    merges: [...(model.merges ?? [])],
    columnWidths,
    rowHeights,
    frozenRows: views?.state === 'frozen' ? views.ySplit : undefined,
    frozenColumns: views?.state === 'frozen' ? views.xSplit : undefined,
    hiddenGridlines: views?.showGridLines === false,
    tables,
    conditionalFormats,
  };
}

export async function importWorkbook(bytes: Uint8Array, path: string, source: WorkbookSource): Promise<WorkbookDocument> {
  const format = formatForPath(path);
  if (format === 'csv' || format === 'tsv') {
    const text = new TextDecoder().decode(bytes);
    const parsed = Papa.parse<string[]>(text, { delimiter: format === 'csv' ? ',' : '\t', skipEmptyLines: false });
    if (parsed.errors.length) throw new Error(parsed.errors[0].message);
    const id = randomUUID();
    const cells: Record<string, CellData> = {};
    parsed.data.forEach((row, rowIndex) => row.forEach((value, columnIndex) => {
      if (value === '') return;
      const numeric = Number(value);
      cells[`${rowIndex}:${columnIndex}`] = Number.isFinite(numeric) && value.trim() !== ''
        ? { value: numeric, valueType: 'number' }
        : { value, valueType: 'text' };
    }));
    return {
      schemaVersion: 1,
      title: parse(path).name,
      activeSheetId: id,
      sheets: [{ id, name: 'Sheet1', rowCount: Math.max(200, parsed.data.length + 20), columnCount: Math.max(26, ...parsed.data.map((row) => row.length + 5)), cells, merges: [], columnWidths: {}, rowHeights: {} }],
      source,
      compatibilityIssues: [],
    };
  }
  const excel = new ExcelJS.Workbook();
  const input = Buffer.from(bytes) as unknown as Parameters<typeof excel.xlsx.load>[0];
  try {
    await excel.xlsx.load(input);
  } catch (error) {
    // Some valid OOXML producers use an `x:` prefix for the main spreadsheet
    // namespace. Excel accepts it, while ExcelJS expects that namespace as the
    // default. Normalize only that equivalent XML representation and retry.
    const zip = await JSZip.loadAsync(bytes);
    // ExcelJS also assumes relative relationship targets and numbered,
    // case-sensitive VML filenames. These are packaging conventions, not OOXML
    // requirements. Normalize aliases in memory without touching the source.
    const aliases = new Map<string, string>();
    let drawingNumber = 1;
    for (const entry of Object.values(zip.files)) {
      if (/^xl\/drawings\/.*\.vml$/iu.test(entry.name)) {
        let target: string;
        do { target = `xl/drawings/vmlDrawing${drawingNumber++}.vml`; }
        while (zip.file(target) || [...aliases.values()].includes(target));
        aliases.set(entry.name, target);
        zip.file(target, await entry.async('uint8array'));
      }
    }
    const entries = Object.values(zip.files).filter((entry) => !entry.dir && entry.name.startsWith('xl/') && /\.(xml|rels)$/u.test(entry.name));
    let normalized = false;
    for (const entry of entries) {
      const xml = await entry.async('string');
      let next = xml;
      for (const declaration of xml.matchAll(/xmlns:([\w.-]+)=(["'])http:\/\/schemas\.openxmlformats\.org\/spreadsheetml\/2006\/main\2/gu)) {
        const prefix = declaration[1].replace(/[.*+?^${}()|[\]\\]/gu, '\\$&');
        next = next.replace(declaration[0], 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"')
          .replace(new RegExp(`<(\\/?)${prefix}:`, 'gu'), '<$1');
      }
      if (entry.name.endsWith('.rels')) {
        const base = posix.dirname(posix.dirname(entry.name));
        next = next.replace(/<(?:\w+:)?Relationship\b[^>]*>/gu, (relationship) => {
          if (/\bTargetMode="External"/iu.test(relationship)) return relationship;
          return relationship.replace(/\bTarget="([^"]+)"/gu, (attribute: string, target: string) => {
            if (/^[a-z][a-z\d+.-]*:/iu.test(target) || target.startsWith('#')) return attribute;
            const absolute = target.startsWith('/') ? target.slice(1) : posix.normalize(posix.join(base, target));
            return `Target="${posix.relative(base, aliases.get(absolute) ?? absolute)}"`;
          });
        });
      }
      if (next === xml) continue;
      zip.file(entry.name, next);
      normalized = true;
    }
    if (!normalized) throw error;
    const repaired = await zip.generateAsync({ type: 'uint8array' });
    const repairedInput = Buffer.from(repaired) as unknown as Parameters<typeof excel.xlsx.load>[0];
    await excel.xlsx.load(repairedInput);
  }
  const sheets = excel.worksheets.map(importWorksheet);
  if (!sheets.length) throw new Error('This workbook does not contain any worksheets.');
  return {
    schemaVersion: 1,
    title: parse(path).name,
    activeSheetId: sheets[0].id,
    sheets,
    source,
    compatibilityIssues: excel.worksheets.flatMap(compatibilityIssuesFor),
  };
}

function excelColor(color: string | undefined): Partial<ExcelJS.Color> | undefined {
  return color ? { argb: `FF${color.replace('#', '').toUpperCase()}` } : undefined;
}

function exportStyle(cell: ExcelJS.Cell, style: CellStyle | undefined): void {
  if (!style) return;
  cell.font = {
    name: style.fontFamily, size: style.fontSize, bold: style.bold, italic: style.italic,
    underline: style.underline, color: excelColor(style.textColor),
  };
  if (style.fillColor) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: excelColor(style.fillColor)! };
  cell.alignment = { horizontal: style.horizontal, vertical: style.vertical, wrapText: style.wrapText };
  if (style.numberFormat) cell.numFmt = style.numberFormat;
  const borderSide = (side: CellBorderSide | undefined): Partial<ExcelJS.Border> | undefined => side ? { style: side.style, color: excelColor(side.color) } : undefined;
  cell.border = {
    top: borderSide(style.border?.top), right: borderSide(style.border?.right),
    bottom: borderSide(style.border?.bottom), left: borderSide(style.border?.left),
  };
}

function exportCellValue(cell: ExcelJS.Cell, data: CellData): void {
  if (data.note !== undefined) cell.note = data.note;
  if (data.validation) cell.dataValidation = structuredClone(data.validation);
  if (data.formula) {
    cell.value = { formula: data.formula.replace(/^=/u, ''), result: data.cachedValue ?? undefined } as ExcelJS.CellFormulaValue;
  } else if (data.valueType === 'date' && typeof data.value === 'string') {
    cell.value = new Date(data.value);
  } else if (data.hyperlink && typeof data.value === 'string') {
    cell.value = { text: data.value, hyperlink: data.hyperlink };
  } else {
    cell.value = data.value;
  }
  exportStyle(cell, data.style);
}

export async function exportWorkbook(workbook: WorkbookDocument, path: string): Promise<Uint8Array> {
  const format = formatForPath(path);
  const active = workbook.sheets.find((sheet) => sheet.id === workbook.activeSheetId) ?? workbook.sheets[0];
  if (format === 'csv' || format === 'tsv') {
    let maxRow = 0;
    let maxColumn = 0;
    for (const key of Object.keys(active.cells)) {
      const [row, column] = key.split(':').map(Number);
      maxRow = Math.max(maxRow, row);
      maxColumn = Math.max(maxColumn, column);
    }
    const rows: CellScalar[][] = Array.from({ length: maxRow + 1 }, (_, row) =>
      Array.from({ length: maxColumn + 1 }, (_, column) => {
        const cell = active.cells[`${row}:${column}`];
        return cell?.formula ? cell.cachedValue ?? '' : cell?.value ?? '';
      }));
    return new TextEncoder().encode(Papa.unparse(rows, { delimiter: format === 'csv' ? ',' : '\t', newline: '\r\n' }));
  }
  const excel = new ExcelJS.Workbook();
  excel.creator = 'TXT Sheets';
  excel.calcProperties.fullCalcOnLoad = true;
  excel.modified = new Date();
  for (const sheet of workbook.sheets) {
    const worksheet = excel.addWorksheet(sheet.name, {
      views: [{ state: sheet.frozenRows || sheet.frozenColumns ? 'frozen' : 'normal', ySplit: sheet.frozenRows, xSplit: sheet.frozenColumns, showGridLines: !sheet.hiddenGridlines }],
    });
    Object.entries(sheet.columnWidths).forEach(([index, width]) => { worksheet.getColumn(Number(index) + 1).width = width / 7.2; });
    Object.entries(sheet.rowHeights).forEach(([index, height]) => { worksheet.getRow(Number(index) + 1).height = height / 1.34; });
    for (const table of sheet.tables ?? []) {
      const [start, last] = table.range.split(':').map(parseCellAddress);
      if (!start || !last) throw new Error('A table has an invalid range.');
      const rows: ExcelJS.CellValue[][] = [];
      for (let row = start.row + (table.headerRow ? 1 : 0); row <= last.row - (table.totalsRow ? 1 : 0); row++) {
        rows.push(Array.from({ length: table.columns.length }, (_, index) => {
          const data = sheet.cells[`${row}:${start.column + index}`];
          return data?.formula ? { formula: data.formula.slice(1), result: data.cachedValue ?? undefined } as ExcelJS.CellFormulaValue : data?.value ?? null;
        }));
      }
      const columns = table.columns.map((column, index) => ({ ...column,
        name: table.headerRow ? String(sheet.cells[`${start.row}:${start.column + index}`]?.value ?? column.name) : column.name,
      }));
      worksheet.addTable({ ...structuredClone(table), ref: addressForCell(start.row, start.column), columns, rows });
    }
    Object.entries(sheet.cells).forEach(([key, data]) => {
      const [row, column] = key.split(':').map(Number);
      exportCellValue(worksheet.getCell(row + 1, column + 1), data);
    });
    sheet.merges.forEach((range) => { try { worksheet.mergeCells(range); } catch { /* invalid imported merge */ } });
    for (const group of sheet.conditionalFormats ?? []) worksheet.addConditionalFormatting(structuredClone(group));
  }
  return new Uint8Array(await excel.xlsx.writeBuffer());
}

export { formatForPath };
