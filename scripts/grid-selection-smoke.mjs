import assert from 'node:assert/strict';
import { mkdtemp, mkdir } from 'node:fs/promises';
import { resolve } from 'node:path';
import ExcelJS from 'exceljs';
import { _electron as electron } from 'playwright';

await mkdir('output/playwright', { recursive: true });
const profile = await mkdtemp(resolve('output/playwright/grid-selection-'));
const path = resolve(profile, 'selection.xlsx');
const fixture = new ExcelJS.Workbook();
const sheet = fixture.addWorksheet('Data');
for (let row = 1; row <= 20; row++) { sheet.getCell(`A${row}`).value = row; sheet.getCell(`B${row}`).value = 100 + row; }
sheet.mergeCells('C12:D13');
sheet.getCell('C12').value = 'Merged';
sheet.getCell('E1').value = { formula: 'A2+A4', result: 6 };
const summary = fixture.addWorksheet('Summary');
summary.getCell('A1').value = { formula: 'Data!A2', result: 2 };
await fixture.xlsx.writeFile(path);
let app;
let page;
const packaged = process.argv.includes('--packaged');
const errors = [];
async function launch() {
  app = await electron.launch({ ...(packaged ? { executablePath: resolve('release/win-unpacked/TXT Sheets.exe') } : {}), args: [...(packaged ? [] : [resolve('.')]), `--user-data-dir=${profile}`, path] });
  page = await app.firstWindow();
  page.on('pageerror', (error) => errors.push(error.message));
  await page.getByRole('gridcell', { name: 'B1', exact: true }).filter({ hasText: '101' }).waitFor();
  await app.evaluate(({ dialog }) => { dialog.showMessageBox = async () => ({ response: 1 }); });
}
const cell = (name) => page.getByRole('gridcell', { name, exact: true });
const header = (kind, name) => page.getByRole(kind + 'header', { name, exact: true }).locator('button');
async function drag(from, to) {
  const start = await from.boundingBox(), end = await to.boundingBox();
  assert.ok(start && end);
  await page.mouse.move(start.x + start.width / 2, start.y + start.height / 2);
  await page.mouse.down();
  await page.mouse.move(end.x + end.width / 2, end.y + end.height / 2, { steps: 12 });
  await page.mouse.up();
}
async function label(expected) { await page.waitForFunction((expected) => document.querySelector('.name-box').textContent === expected, expected); }
async function menu(at, action) {
  await at.click({ button: 'right' });
  await page.getByRole('menu', { name: 'Selection actions' }).waitFor();
  if (action) await page.getByRole('menuitem', { name: action, exact: true }).click();
}
async function undo() { await page.getByRole('grid').focus(); await page.keyboard.press('Control+z'); }
try {
  await launch();
  await drag(cell('A1'), cell('C3')); await label('A1:C3');
  assert.equal(await page.locator('[role="gridcell"][aria-selected="true"]').count(), 9);
  await menu(cell('B2')); await label('A1:C3');
  await page.screenshot({ path: resolve(profile, 'context-menu.png') });
  const box = await page.getByRole('menu').boundingBox();
  const view = await page.evaluate(() => ({ width: innerWidth, height: innerHeight }));
  assert.ok(box.x >= 0 && box.y >= 0 && box.x + box.width <= view.width && box.y + box.height <= view.height);
  await page.keyboard.press('Escape');
  assert.equal(await page.getByRole('grid').evaluate((node) => node === document.activeElement), true);
  await page.getByRole('button', { name: 'Use dark theme', exact: true }).click();
  await menu(cell('B2'));
  await page.screenshot({ path: resolve(profile, 'context-menu-dark.png') });
  await page.keyboard.press('Escape');
  await page.getByRole('button', { name: 'Use light theme', exact: true }).click();
  await drag(cell('C3'), cell('A1')); await label('A1:C3');
  await drag(cell('A1'), cell('C3'));
  await menu(cell('B2'), 'Copy');
  await menu(cell('F5'), 'Paste');
  await page.waitForFunction(() => document.querySelector('[aria-label="F5"]').textContent === '1');
  await menu(cell('F5'), 'Clear contents');
  await page.waitForFunction(() => document.querySelector('[aria-label="F5"]').textContent === '');
  await menu(cell('E5')); await label('E5'); await page.keyboard.press('Escape');
  await cell('C12').dblclick();
  assert.equal(await page.getByRole('textbox', { name: 'Edit C12', exact: true }).inputValue(), 'Merged');
  await page.keyboard.press('Escape');
  await cell('C10').dblclick();
  await page.getByRole('textbox', { name: 'Edit C10', exact: true }).fill('unfinished edit');
  await cell('A1').click();
  assert.equal(await cell('C10').innerText(), 'unfinished edit');
  await drag(header('row', 'Row 2'), header('row', 'Row 4')); await label('A2:Z4');
  await menu(header('row', 'Row 3'), 'Insert rows above');
  await page.waitForFunction(() => document.querySelector('[aria-label="A5"]').textContent === '2');
  await undo(); await page.waitForFunction(() => document.querySelector('[aria-label="A2"]').textContent === '2');
  await drag(header('column', 'Column B'), header('column', 'Column D')); await label('B1:D200');
  await menu(header('column', 'Column C'), 'Insert columns left');
  await page.waitForFunction(() => document.querySelector('[aria-label="E1"]').textContent === '101');
  await undo();
  await drag(cell('A1'), cell('A2'));
  await drag(page.getByRole('button', { name: 'Drag to fill selected cells', exact: true }), cell('A4'));
  await page.waitForFunction(() => document.querySelector('[aria-label="A3"]').textContent === '1');
  await undo();
  await drag(cell('A2'), cell('A3')); await label('A2:A3');
  await menu(cell('A2'), 'Delete cells, shift up');
  await page.waitForFunction(() => document.querySelector('[aria-label="A2"]').textContent === '4');
  await undo();
  await drag(cell('A2'), cell('A3')); await menu(cell('A2'), 'Insert cells, shift down');
  await page.waitForFunction(() => document.querySelector('[aria-label="A4"]').textContent === '2');
  await page.getByRole('grid').focus(); await page.keyboard.press('Shift+F10');
  await page.getByRole('menu', { name: 'Selection actions' }).waitFor();
  await page.keyboard.press('Escape');
  await page.getByRole('button', { name: 'Save', exact: true }).click();
  await page.getByText('Saved selection.xlsx', { exact: true }).waitFor();
  const saved = new ExcelJS.Workbook(); await saved.xlsx.readFile(path);
  assert.equal(saved.getWorksheet('Data').getCell('A4').value, 2);
  assert.equal(saved.getWorksheet('Data').getCell('E1').value.formula, 'A4+A6');
  assert.equal(saved.getWorksheet('Summary').getCell('A1').value.formula, 'Data!A4');
  await app.close(); await launch();
  await page.waitForFunction(() => document.querySelector('[aria-label="A4"]').textContent === '2');
  // Auto-scroll while dragging past the visible bottom edge.
  const start = await cell('A4').boundingBox(), grid = await page.getByRole('grid').boundingBox();
  await page.mouse.move(start.x + 20, start.y + start.height / 2); await page.mouse.down();
  await page.mouse.move(start.x + 20, grid.y + grid.height - 2, { steps: 8 });
  await page.waitForTimeout(600);
  assert.ok(await page.getByRole('grid').evaluate((node) => node.scrollTop) > 0);
  await page.mouse.move(start.x + 20, grid.y + grid.height + 20);
  await page.mouse.up();
  const stopped = await page.locator('.name-box').innerText();
  await page.mouse.move(grid.x + 150, grid.y + 100);
  assert.equal(await page.locator('.name-box').innerText(), stopped);
  assert.deepEqual(errors, []);
  await page.screenshot({ path: resolve(profile, 'drag-selection.png') });
  console.log('PASS: cell/row/column drag, reverse selection, auto-scroll/release, context targeting, keyboard menus, bulk insert, cell shifts, Undo, Save and reopen');
} finally { await app?.close(); }
