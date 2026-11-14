import type { CellBorderSide, CellStyle, SheetDocument, WorkbookPrintRequest } from '../shared/types';
import { displayValue } from '../shared/cell-display';
import { cellKey, createFormulaEvaluator, parseCellAddress } from '../renderer/formulas';
import { conditionalStyle } from '../renderer/conditional-formatting';

function escape(value: string): string {
  return value.replace(/[&<>"']/gu, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[char]!);
}

function bounds(sheet: SheetDocument) {
  let top = Infinity, left = Infinity, bottom = -1, right = -1;
  for (const [key, cell] of Object.entries(sheet.cells)) {
    const visible = cell.formula || (cell.value !== null && cell.value !== '') || Object.values(cell.style?.border ?? {}).some((side) => side?.style)
      || (cell.style?.fillColor && !/^#(?:fff|ffffff)$/iu.test(cell.style.fillColor));
    if (!visible) continue;
    const [row, column] = key.split(':').map(Number);
    top = Math.min(top, row); left = Math.min(left, column); bottom = Math.max(bottom, row); right = Math.max(right, column);
  }
  if (bottom < 0) return { top: 0, left: 0, bottom: 0, right: 0 };
  for (const merge of sheet.merges) {
    const [first, last] = merge.split(':').map(parseCellAddress);
    if (first && last && first.row <= bottom && first.column <= right && last.row >= top && last.column >= left) {
      top = Math.min(top, first.row); left = Math.min(left, first.column); bottom = Math.max(bottom, last.row); right = Math.max(right, last.column);
    }
  }
  return { top, left, bottom, right };
}

function cellCss(style: CellStyle, scale: number, gridlines: boolean): string {
  const side = (value?: CellBorderSide) => value?.style ? `${(value.style === 'thick' || value.style === 'double' ? 3 : value.style === 'medium' ? 2 : 1) * scale}px ${['dotted', 'dashed', 'double'].includes(value.style) ? value.style : 'solid'} ${value.color ?? '#aab0b8'}` : gridlines ? `${scale}px solid #d9dde2` : '0';
  return `font-family:${JSON.stringify(style.fontFamily ?? 'Calibri')},"Segoe UI",sans-serif;font-size:${(style.fontSize ?? 11) * 4 / 3 * scale}px;font-weight:${style.bold ? 700 : 400};font-style:${style.italic ? 'italic' : 'normal'};text-decoration:${style.underline ? 'underline' : 'none'};color:${style.textColor ?? '#202124'};background:${style.fillColor ?? '#fff'};text-align:${style.horizontal ?? 'left'};vertical-align:${style.vertical === 'middle' ? 'middle' : style.vertical ?? 'middle'};white-space:${style.wrapText ? 'pre-wrap' : 'nowrap'};padding:${3 * scale}px ${4 * scale}px;border-top:${side(style.border?.top)};border-right:${side(style.border?.right)};border-bottom:${side(style.border?.bottom)};border-left:${side(style.border?.left)};`;
}

export function buildWorkbookPrintHtml(request: WorkbookPrintRequest): string {
  const evaluator = createFormulaEvaluator(request.workbook);
  const sheets = request.scope === 'all' ? request.workbook.sheets : request.workbook.sheets.filter((sheet) => sheet.id === request.workbook.activeSheetId);
  if (!sheets.length) throw new Error('There is no worksheet to print.');
  const availableWidth = ((request.landscape ? 11 : 8.5) - 0.8) * 96;
  const sections = sheets.map((sheet) => {
    const area = bounds(sheet);
    if ((area.bottom - area.top + 1) * (area.right - area.left + 1) > 250_000) throw new Error('This worksheet is too large to print at once. Copy the needed range into a smaller sheet.');
    const widths = Array.from({ length: area.right - area.left + 1 }, (_, offset) => sheet.columnWidths[String(area.left + offset)] ?? 96);
    const scale = Math.min(1, availableWidth / widths.reduce((sum, width) => sum + width, 0));
    const merged = new Map<string, { top: number; left: number; bottom: number; right: number }>();
    for (const range of sheet.merges) {
      const [first, last] = range.split(':').map(parseCellAddress);
      if (!first || !last) continue;
      const rect = { top: first.row, left: first.column, bottom: last.row, right: last.column };
      for (let row = Math.max(first.row, area.top); row <= Math.min(last.row, area.bottom); row++) for (let column = Math.max(first.column, area.left); column <= Math.min(last.column, area.right); column++) merged.set(cellKey(row, column), rect);
    }
    const rows: string[] = [];
    for (let row = area.top; row <= area.bottom; row++) {
      const cells: string[] = [];
      for (let column = area.left; column <= area.right; column++) {
        const merge = merged.get(cellKey(row, column));
        if (merge && (row !== merge.top || column !== merge.left)) continue;
        const cell = sheet.cells[cellKey(row, column)];
        const style = { ...cell?.style, ...conditionalStyle(sheet, row, column, evaluator) };
        const text = displayValue(cell, evaluator.evaluateCell(sheet.id, row, column));
        let spillWidth = widths[column - area.left];
        if (!merge && !style.wrapText && (!style.horizontal || style.horizontal === 'left') && typeof cell?.value === 'string' && !cell.formula) {
          let blocked = false;
          for (let next = column + 1; next <= area.right; next++) {
            const neighbor = sheet.cells[cellKey(row, next)];
            if (merged.has(cellKey(row, next)) || neighbor?.formula || (neighbor?.value !== null && neighbor?.value !== undefined && neighbor.value !== '')) { blocked = true; break; }
            spillWidth += widths[next - area.left];
          }
          if (!blocked) spillWidth = Math.max(spillWidth, availableWidth / scale - widths.slice(0, column - area.left).reduce((sum, width) => sum + width, 0));
        }
        const span = merge ? ` colspan="${merge.right - merge.left + 1}" rowspan="${merge.bottom - merge.top + 1}"` : '';
        const content = spillWidth > widths[column - area.left] ? `<span class="spill" style="width:${(spillWidth - 8) * scale}px">${escape(text)}</span>` : `<span class="value">${escape(text)}</span>`;
        cells.push(`<td${span} style="${escape(cellCss(style, scale, !sheet.hiddenGridlines))}">${content}</td>`);
      }
      rows.push(`<tr style="height:${(sheet.rowHeights[String(row)] ?? 25) * scale}px">${cells.join('')}</tr>`);
    }
    const columns = `<colgroup>${widths.map((width) => `<col style="width:${width * scale}px">`).join('')}</colgroup>`;
    const table = (body: string) => `<table style="width:${widths.reduce((sum, width) => sum + width, 0) * scale}px">${columns}${body}</table>`;
    // Repeat just a native table's column headings, not the summary above it.
    const header = (sheet.tables ?? []).filter((table) => table.headerRow).map((table) => parseCellAddress(table.ref)?.row).find((row) => row !== undefined && row >= area.top && row <= area.bottom);
    const split = header === undefined || sheet.merges.some((range) => { const [first, last] = range.split(':').map(parseCellAddress); return first && last && first.row <= header && last.row > header; }) ? -1 : header - area.top;
    const body = split < 0 ? table(`<tbody>${rows.join('')}</tbody>`) : (split ? table(`<tbody>${rows.slice(0, split).join('')}</tbody>`) : '') + table(`<thead>${rows[split]}</thead><tbody>${rows.slice(split + 1).join('')}</tbody>`);
    return `<section>${sheets.length > 1 ? `<h1>${escape(sheet.name)}</h1>` : ''}${body}</section>`;
  });
  return `<!doctype html><html><head><meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'"><title>${escape(request.workbook.title)}</title><style>
    @page { size: Letter ${request.landscape ? 'landscape' : 'portrait'}; margin: .4in; }
    * { box-sizing:border-box; print-color-adjust:exact; -webkit-print-color-adjust:exact; }
    html,body { margin:0; padding:0; color:#202124; background:#fff; font-family:Calibri,"Segoe UI",sans-serif; }
    section + section { break-before:page; } h1 { font-size:12pt; margin:0 0 10px; }
    table { border-collapse:collapse; table-layout:fixed; } td { position:relative; overflow:visible; line-height:1.2; }
    .value { display:block; overflow:hidden; overflow-wrap:anywhere; } .spill { display:block; position:relative; z-index:1; overflow:hidden; white-space:nowrap; }
    tr { break-inside:avoid; } thead { display:table-header-group; }
    </style></head><body>${sections.join('')}</body></html>`;
}
