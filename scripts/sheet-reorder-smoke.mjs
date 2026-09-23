import assert from 'node:assert/strict';
import { mkdir, mkdtemp } from 'node:fs/promises';
import { resolve } from 'node:path';
import { _electron as electron } from 'playwright';

await mkdir(resolve('output/playwright'), { recursive: true });
const userData = await mkdtemp(resolve('output/playwright/sheet-reorder-'));
const packaged = process.argv.includes('--packaged');
const application = await electron.launch({
  ...(packaged ? { executablePath: resolve('release/win-unpacked/TXT Sheets.exe') } : {}),
  args: [...(packaged ? [] : [resolve('.')]), `--user-data-dir=${userData}`, resolve('tests/fixtures/fidelity-fixture.xlsx')],
  cwd: resolve('.'),
  env: { ...process.env, ELECTRON_DISABLE_SECURITY_WARNINGS: 'true' },
});

try {
  const page = await application.firstWindow();
  const tabs = page.getByRole('tablist', { name: 'Worksheets' }).getByRole('tab');
  await page.getByRole('grid', { name: /Overview spreadsheet grid/u }).waitFor();
  assert.deepEqual(await tabs.allTextContents(), ['Overview', 'Inputs']);

  const source = await tabs.nth(0).boundingBox();
  const target = await tabs.nth(1).boundingBox();
  assert.ok(source && target);
  await page.mouse.move(source.x + source.width / 2, source.y + source.height / 2);
  await page.mouse.down();
  await page.mouse.move(target.x + target.width - 5, target.y + target.height / 2, { steps: 12 });
  await page.locator('.sheet-tabs .drop-after').waitFor();
  await page.mouse.up();
  assert.deepEqual(await tabs.allTextContents(), ['Inputs', 'Overview']);
  assert.equal(await page.getByRole('tab', { name: 'Overview' }).getAttribute('aria-selected'), 'true');

  await page.getByRole('button', { name: 'Undo' }).click();
  assert.deepEqual(await tabs.allTextContents(), ['Overview', 'Inputs']);
  await page.getByRole('button', { name: 'Redo' }).click();
  assert.deepEqual(await tabs.allTextContents(), ['Inputs', 'Overview']);

  await page.getByRole('tab', { name: 'Overview' }).focus();
  await page.keyboard.press('Shift+F10');
  await page.getByRole('menuitem', { name: 'Move left' }).click();
  assert.deepEqual(await tabs.allTextContents(), ['Overview', 'Inputs']);
  console.log('PASS: drag reorder, active sheet, undo/redo, and keyboard menu move.');
} finally {
  const child = application.process();
  const exited = child.exitCode === null ? new Promise((resolveExit) => child.once('exit', resolveExit)) : Promise.resolve();
  await application.evaluate(({ app }) => { setImmediate(() => app.exit(0)); });
  await exited;
}
