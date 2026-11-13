import assert from 'node:assert/strict';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import ExcelJS from 'exceljs';
import { _electron as electron } from 'playwright';

const directory = await mkdtemp(join(tmpdir(), 'txt-sheets-compatibility-'));
const input = join(directory, 'fixture.xlsx');
const output = join(directory, 'saved.xlsx');
const excel = new ExcelJS.Workbook();
const sheet = excel.addWorksheet('Actions', { views: [{ showGridLines: false }] });
sheet.getColumn('A').width = 6;
sheet.getColumn('B').width = 30;
sheet.getColumn('C').width = 18;
sheet.getCell('A2').value = 'A heading that extends into empty cells';
sheet.getCell('A2').font = { name: 'Unavailable Test Font', size: 18, bold: true };
sheet.getRow(2).height = 28;
sheet.getCell('A4').value = 'Done';
sheet.getCell('A4').note = 'Source note';
sheet.getCell('A5').value = 'Draft';
sheet.getCell('C2').value = { formula: 'COUNTIFS(A4:A5,"Done")', result: 99 };
sheet.getCell('B4').value = 'First line\nSecond line';
sheet.getCell('B4').alignment = { wrapText: true };
sheet.getRow(4).height = 40;
await excel.xlsx.writeFile(input);

const packaged = process.argv.includes('--packaged');
const app = await electron.launch({
  ...(packaged ? { executablePath: resolve('release/win-unpacked/TXT Sheets.exe') } : {}),
  args: [...(packaged ? [] : [resolve('.')]), `--user-data-dir=${join(directory, 'profile')}`, input],
});
const errors = [];
try {
  const page = await app.firstWindow();
  page.on('pageerror', (error) => errors.push(error.message));
  const cell = (address) => page.getByRole('gridcell', { name: address, exact: true });
  await cell('A2').filter({ hasText: 'A heading' }).waitFor();
  assert.equal(await cell('C2').innerText(), '1');
  assert.equal(await cell('A4').getAttribute('title'), 'Source note');
  assert.equal(await cell('A2').evaluate((node) => {
    const span = node.querySelector('span');
    return span.clientWidth > node.clientWidth && span.clientWidth < node.clientWidth + 300;
  }), true);
  assert.match(await cell('A2').evaluate((node) => getComputedStyle(node).fontFamily), /sans-serif/u);
  assert.equal(await cell('B4').evaluate((node) => getComputedStyle(node).whiteSpace), 'pre-wrap');
  assert.equal(await cell('B4').evaluate((node) => getComputedStyle(node).borderRightColor), 'rgba(0, 0, 0, 0)');
  await cell('A5').dblclick();
  await page.getByRole('textbox', { name: 'Edit A5', exact: true }).fill('Done');
  await page.keyboard.press('Enter');
  assert.equal(await cell('C2').innerText(), '2');
  await page.getByRole('grid').focus();
  await page.keyboard.press('Control+z');
  assert.equal(await cell('C2').innerText(), '1');
  await page.screenshot({ path: join(directory, 'render.png') });
  await app.evaluate(({ dialog }, filePath) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath });
  }, output);
  await page.keyboard.press('Control+Shift+s');
  await page.getByText('Saved as saved.xlsx', { exact: true }).waitFor();
  const saved = new ExcelJS.Workbook();
  await saved.xlsx.readFile(output);
  assert.equal(saved.worksheets[0].getCell('A4').note, 'Source note');
  assert.equal(saved.worksheets[0].getCell('C2').result, 1);
  assert.equal(saved.worksheets[0].views[0].showGridLines, false);
  assert.deepEqual(errors, []);
  console.log('Workbook compatibility UI checks passed.');
} finally {
  await app.close();
}
