import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import Papa from 'papaparse';
import { randomUUID } from 'node:crypto';
import { extname, parse } from 'node:path';
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
  if (model.tables?.length) issues.push({ feature: 'Excel tables', detail: `${worksheet.name} contains structured tables. Values and styling are shown, but table metadata may be simplified on save.` });
  if (model.conditionalFormattings?.length) issues.push({ feature: 'Conditional formatting', detail: `${worksheet.name} contains conditional formatting that is not evaluated by this basic editor.` });
  if (model.dataValidations?.model && Object.keys(model.dataValidations.model).length) issues.push({ feature: 'Data validation', detail: `${worksheet.name} contains validation rules that are not editable here.` });
  if (worksheet.state !== 'visible') issues.push({ feature: 'Hidden sheets', detail: `${worksheet.name} was hidden and is visible while editing here.` });
  const supportedFunctions = new Set(['SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT']);
  const unsupportedFunctions = new Set<string>();
  worksheet.eachRow({ includeEmpty: false }, (row) => row.eachCell({ includeEmpty: false }, (cell) => {
    if (!cell.formula) return;
    for (const match of cell.formula.matchAll(/\b([A-Z][A-Z0-9.]*)\s*\(/giu)) {
      const name = match[1].toUpperCase();
      if (!supportedFunctions.has(name)) unsupportedFunctions.add(name);
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
      if ((cell.value === null || cell.value === undefined) && !style) return;
      const imported = formulaCell(cell) ?? { ...scalarFromExcel(cell.value), style };
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
  const model = worksheet.model as ExcelJS.WorksheetModel & { merges?: string[] };
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
    const entries = Object.values(zip.files).filter((entry) => !entry.dir && entry.name.startsWith('xl/') && entry.name.endsWith('.xml'));
    let normalized = false;
    for (const entry of entries) {
      const xml = await entry.async('string');
      if (!xml.includes('xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"')) continue;
      const next = xml
        .replace('xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"', 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"')
        .replace(/<(\/?)x:/gu, '<$1');
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
    const used = Object.keys(active.cells).map((key) => key.split(':').map(Number));
    const maxRow = Math.max(0, ...used.map(([row]) => row));
    const maxColumn = Math.max(0, ...used.map(([, column]) => column));
    const rows: CellScalar[][] = Array.from({ length: maxRow + 1 }, (_, row) =>
      Array.from({ length: maxColumn + 1 }, (_, column) => {
        const cell = active.cells[`${row}:${column}`];
        return cell?.formula ? cell.cachedValue ?? '' : cell?.value ?? '';
      }));
    return new TextEncoder().encode(Papa.unparse(rows, { delimiter: format === 'csv' ? ',' : '\t', newline: '\r\n' }));
  }
  const excel = new ExcelJS.Workbook();
  excel.creator = 'TXT Sheets';
  excel.modified = new Date();
  for (const sheet of workbook.sheets) {
    const worksheet = excel.addWorksheet(sheet.name, {
      views: [{ state: sheet.frozenRows || sheet.frozenColumns ? 'frozen' : 'normal', ySplit: sheet.frozenRows, xSplit: sheet.frozenColumns, showGridLines: !sheet.hiddenGridlines }],
    });
    Object.entries(sheet.columnWidths).forEach(([index, width]) => { worksheet.getColumn(Number(index) + 1).width = width / 7.2; });
    Object.entries(sheet.rowHeights).forEach(([index, height]) => { worksheet.getRow(Number(index) + 1).height = height / 1.34; });
    Object.entries(sheet.cells).forEach(([key, data]) => {
      const [row, column] = key.split(':').map(Number);
      exportCellValue(worksheet.getCell(row + 1, column + 1), data);
    });
    sheet.merges.forEach((range) => { try { worksheet.mergeCells(range); } catch { /* invalid imported merge */ } });
  }
  return new Uint8Array(await excel.xlsx.writeBuffer());
}

export { formatForPath };
