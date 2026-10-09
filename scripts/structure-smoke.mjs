import assert from 'node:assert/strict';
import { mkdir, mkdtemp, readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import ExcelJS from 'exceljs';
import { _electron as electron } from 'playwright';

const root = resolve('.');
const packaged = process.argv.includes('--packaged');
await mkdir(resolve('output/playwright'), { recursive: true });
const profile = await mkdtemp(resolve('output/playwright/structure-'));
const file = resolve(profile, 'structure.xlsx');
const fixture = new ExcelJS.Workbook();
const data = fixture.addWorksheet('Data', { views: [{ state: 'frozen', ySplit: 2, xSplit: 1 }] });
data.getCell('A2').value = 7;
data.getCell('A3').value = 11;
data.getCell('C1').value = { formula: '$A$2+A3', result: 18 };
data.mergeCells('D2:F3');
const summary = fixture.addWorksheet('Summary');
summary.getCell('A1').value = { formula: 'SUM(Data!A2:A3)', result: 18 };
summary.getCell('B1').value = { formula: 'Data!A2', result: 7 };
await fixture.xlsx.writeFile(file);
let application;
const errors = [];

async function launch() {
  application = await electron.launch({ ...(packaged ? { executablePath: resolve(root, 'release/win-unpacked/TXT Sheets.exe') } : {}), args: [...(packaged ? [] : [root]), `--user-data-dir=${profile}`, file], cwd: root });
  const page = await application.firstWindow();
  page.on('pageerror', (error) => errors.push(error.message));
  await page.getByRole('grid', { name: 'Data spreadsheet grid' }).waitFor();
  await page.getByRole('gridcell', { name: 'C1', exact: true }).waitFor();
  return page;
}

async function stop() {
  if (!application) return;
  const child = application.process();
  const exited = child.exitCode === null ? new Promise((done) => child.once('exit', done)) : Promise.resolve();
  await application.evaluate(({ app }) => { setImmediate(() => app.exit(0)); }).catch(() => {});
  await exited;
}

async function action(page, address, label) {
  await page.getByRole('gridcell', { name: address, exact: true }).click();
  await page.getByRole('button', { name: 'Rows and columns', exact: true }).click();
  await page.getByRole('button', { name: label, exact: true }).click();
}

async function formula(page, address, expected) {
  await page.getByRole('gridcell', { name: address, exact: true }).dblclick();
  assert.equal(await page.getByRole('textbox', { name: `Edit ${address}`, exact: true }).inputValue(), expected);
  await page.keyboard.press('Escape');
}

try {
  let page = await launch();
  await page.getByRole('gridcell', { name: 'C1', exact: true }).dblclick();
  await page.getByRole('textbox', { name: 'Edit C1', exact: true }).fill('=999');
  await page.keyboard.press('Escape');
  assert.equal(await page.getByRole('gridcell', { name: 'C1', exact: true }).innerText(), '18');
  assert.equal(await page.getByRole('button', { name: 'Undo', exact: true }).isDisabled(), true, 'Escape must not commit the cancelled edit');
  await action(page, 'A2', 'Insert row above');
  await formula(page, 'C1', '=$A$3+A4');
  assert.equal(await page.getByRole('gridcell', { name: 'C1', exact: true }).innerText(), '18');
  await page.getByRole('gridcell', { name: 'C1', exact: true }).dblclick();
  await page.keyboard.press('Enter');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await formula(page, 'C1', '=$A$2+A3');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await formula(page, 'C1', '=$A$3+A4');

  await action(page, 'A3', 'Delete row');
  await formula(page, 'C1', '=#REF!+A3');
  await page.getByRole('tab', { name: 'Summary', exact: true }).click();
  assert.equal(await page.getByRole('gridcell', { name: 'A1', exact: true }).innerText(), '11');
  assert.equal(await page.getByRole('gridcell', { name: 'B1', exact: true }).innerText(), '#REF!');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await page.getByRole('tab', { name: 'Data', exact: true }).click();

  await action(page, 'A1', 'Delete column');
  await formula(page, 'B1', '=#REF!+#REF!');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await action(page, 'A1', 'Insert column left');
  await formula(page, 'D1', '=$B$3+B4');
  await page.getByRole('tab', { name: 'Summary', exact: true }).click();
  await formula(page, 'A1', '=SUM(Data!B3:B4)');
  assert.equal(await page.getByRole('gridcell', { name: 'A1', exact: true }).innerText(), '18');
  await page.getByRole('tab', { name: 'Data', exact: true }).click();
  await page.getByRole('button', { name: 'Save', exact: true }).click();
  await page.getByText('Saved structure.xlsx', { exact: true }).waitFor();
  const saved = new ExcelJS.Workbook();
  await saved.xlsx.load(await readFile(file));
  assert.equal(saved.getWorksheet('Data').getCell('D1').formula, '$B$3+B4');
  assert.equal(saved.getWorksheet('Summary').getCell('A1').formula, 'SUM(Data!B3:B4)');
  assert.deepEqual(saved.getWorksheet('Data').model.merges, ['E3:G4']);
  assert.equal(saved.getWorksheet('Data').views[0].ySplit, 3);
  assert.equal(saved.getWorksheet('Data').views[0].xSplit, 2);
  await stop();
  page = await launch();
  assert.equal(await page.getByRole('gridcell', { name: 'D1', exact: true }).innerText(), '18');
  await formula(page, 'D1', '=$B$3+B4');
  await page.getByRole('tab', { name: 'Summary', exact: true }).click();
  assert.equal(await page.getByRole('gridcell', { name: 'A1', exact: true }).innerText(), '18');
  assert.deepEqual(errors, []);
  console.log('PASS: row/column insert/delete, same-sheet and cross-sheet references, #REF!, Undo/Redo, merged ranges, frozen panes, Save, and reopen.');
} finally { await stop(); }
