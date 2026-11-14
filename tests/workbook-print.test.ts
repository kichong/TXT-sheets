import { describe, expect, it } from 'vitest';
import { createBlankWorkbook } from '../src/shared/types';
import { buildWorkbookPrintHtml } from '../src/main/workbook-print';

describe('workbook print surface', () => {
  it('prints all used rows, computed results, styles, and merges without editor chrome or HTML injection', () => {
    const w = createBlankWorkbook(), sheet = w.sheets[0];
    sheet.cells = { '0:0': { value: '<script>private & literal</script>', style: { bold: true } },
      '300:0': { value: 3 }, '301:0': { value: null, formula: '=A301*2' }, '2:0': { value: 'Merged' } };
    sheet.merges = ['A3:C4'];
    const html = buildWorkbookPrintHtml({ workbook: w, scope: 'active', landscape: true });
    expect(html).toContain('Letter landscape');
    expect(html).toContain('&lt;script&gt;private &amp; literal&lt;/script&gt;');
    expect(html).not.toContain('<script>');
    expect(html).toContain('colspan="3" rowspan="2"');
    expect(html).toContain('>6</span>');
    expect(html).not.toContain('role="grid"');
    expect(html).not.toContain('cell-editor');
  });
  it('scopes output to the active sheet or the entire workbook and preserves conditional colors', () => {
    const w = createBlankWorkbook(), sheet = w.sheets[0];
    sheet.cells['0:0'] = { value: 'Done' };
    sheet.conditionalFormats = [{ ref: 'A1', rules: [{ type: 'containsText', operator: 'containsText', text: 'Done', priority: 1, style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFDCFCE7' } } } }] }];
    const other = createBlankWorkbook().sheets[0]; other.name = 'Other'; other.cells['0:0'] = { value: 'Second worksheet' }; w.sheets.push(other);
    const active = buildWorkbookPrintHtml({ workbook: w, scope: 'active', landscape: false });
    expect(active).toContain('Letter portrait');
    expect(active).toContain('#DCFCE7');
    expect(active).not.toContain('Second worksheet');
    expect(buildWorkbookPrintHtml({ workbook: w, scope: 'all', landscape: true })).toContain('Second worksheet');
  });
});
