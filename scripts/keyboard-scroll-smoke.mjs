import assert from 'node:assert/strict';
import { mkdir, mkdtemp } from 'node:fs/promises';
import { resolve } from 'node:path';
import { _electron as electron } from 'playwright';

const output = resolve('output/playwright');
await mkdir(output, { recursive: true });
const userData = await mkdtemp(resolve(output, 'keyboard-scroll-'));
const application = await electron.launch({
  ...(process.argv.includes('--packaged') ? { executablePath: resolve('release/win-unpacked/TXT Sheets.exe') } : {}),
  args: [...(process.argv.includes('--packaged') ? [] : [resolve('.')]), `--user-data-dir=${userData}`],
  cwd: resolve('.'),
  env: { ...process.env, ELECTRON_DISABLE_SECURITY_WARNINGS: 'true' },
});

try {
  const page = await application.firstWindow();
  await page.getByRole('grid').waitFor();
  await application.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].setSize(900, 650));
  const grid = page.getByRole('grid');
  await grid.focus();

  async function visibleActiveCell(address) {
    await page.waitForFunction((expected) => {
      const grid = document.querySelector('[role="grid"]');
      const cell = grid?.querySelector('.grid-cell.is-active');
      if (cell?.getAttribute('aria-label') !== expected) return false;
      const viewport = grid.getBoundingClientRect();
      const rect = cell.getBoundingClientRect();
      return rect.left >= viewport.left + 46 - 1 && rect.top >= viewport.top + 26 - 1
        && rect.right <= viewport.left + grid.clientWidth + 1
        && rect.bottom <= viewport.top + grid.clientHeight + 1;
    }, address);
  }

  async function press(key, count) {
    for (let i = 0; i < count; i += 1) await page.keyboard.press(key);
  }

  await press('ArrowRight', 20);
  await press('ArrowDown', 40);
  await visibleActiveCell('U41');
  await page.screenshot({ path: resolve(output, 'keyboard-scroll.png') });
  await press('Shift+ArrowRight', 4);
  await press('Shift+ArrowDown', 8);
  await visibleActiveCell('Y49');
  await page.keyboard.press('Tab');
  await visibleActiveCell('Z49');
  await page.keyboard.press('Enter');
  await visibleActiveCell('Z50');
  await press('ArrowLeft', 25);
  await press('ArrowUp', 49);
  await visibleActiveCell('A1');

  // Clicking a whole-column header must not jump to the final row.
  await page.getByRole('columnheader', { name: 'Column A', exact: true }).getByRole('button').click();
  assert.equal(await grid.evaluate((element) => element.scrollTop), 0);
  await page.getByRole('gridcell', { name: 'A1', exact: true }).click();
  await grid.focus();
  await press('ArrowLeft', 2);
  await press('ArrowUp', 2);
  await visibleActiveCell('A1');

  // Resized rows and columns use the same measured coordinates.
  await page.getByRole('separator', { name: 'Resize column A', exact: true }).press('Shift+ArrowRight');
  await page.getByRole('separator', { name: 'Resize row 1', exact: true }).press('Shift+ArrowDown');
  await grid.focus();
  await press('ArrowRight', 20);
  await press('ArrowDown', 40);
  await visibleActiveCell('U41');
  console.log('PASS: four directions, Shift selection, Tab, Enter, first-cell boundary, header selection, and resized cells.');
} finally {
  const child = application.process();
  const exited = child.exitCode === null ? new Promise((resolveExit) => child.once('exit', resolveExit)) : Promise.resolve();
  await application.evaluate(({ app }) => { setImmediate(() => app.exit(0)); });
  await exited;
}
