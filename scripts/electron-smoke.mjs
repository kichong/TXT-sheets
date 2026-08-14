import fs from 'node:fs/promises';
import { resolve } from 'node:path';
import { _electron as electron } from 'playwright';

const root = resolve('.');
const packaged = process.argv.includes('--packaged');
const fixture = resolve('tests/fixtures/fidelity-fixture.xlsx');
const output = resolve(`output/playwright/txt-sheets-${packaged ? 'packaged' : 'electron'}-fixture.png`);
const userData = resolve(`output/playwright/${packaged ? 'packaged' : 'electron'}-smoke-user-data`);
await fs.mkdir(resolve('output/playwright'), { recursive: true });

async function launch(extraArgs = []) {
  return electron.launch({
    ...(packaged ? { executablePath: resolve('release/win-unpacked/TXT Sheets.exe') } : {}),
    args: packaged ? [`--user-data-dir=${userData}`, ...extraArgs] : [root, `--user-data-dir=${userData}`, ...extraArgs],
    cwd: root,
    env: { ...process.env, ELECTRON_DISABLE_SECURITY_WARNINGS: 'true', TXT_SHEETS_SMOKE_LOG: '1' },
  });
}

async function stop(application) {
  const child = application.process();
  const exited = child.exitCode === null ? new Promise((resolveExit) => child.once('exit', resolveExit)) : Promise.resolve();
  await application.evaluate(({ app }) => { setImmediate(() => app.exit(0)); });
  await exited;
}

const application = await launch([fixture]);
let applicationStopped = false;

application.process().stderr?.on('data', (data) => console.error(`[main] ${data}`));

try {
  const argv = await application.evaluate(() => process.argv);
  console.log(JSON.stringify({ packaged, argv }));
  const window = await application.firstWindow();
  window.on('console', (message) => console.log(JSON.stringify({ browserConsole: message.type(), text: message.text() })));
  await window.waitForLoadState('domcontentloaded');
  if (packaged) {
    await window.waitForTimeout(3_000);
    console.log(JSON.stringify({ grids: await window.locator('[role="grid"]').getAttribute('aria-label'), body: (await window.locator('body').innerText()).slice(0, 600) }));
  }
  await window.getByRole('grid', { name: /Overview spreadsheet grid/u }).waitFor({ timeout: 15_000 });
  const total = await window.getByRole('gridcell', { name: 'D9', exact: true }).textContent();
  if (total?.trim() !== '$208.00') throw new Error(`Expected $208.00 in D9, received ${JSON.stringify(total)}`);
  const tabs = await window.getByRole('tab').allTextContents();
  if (tabs.join('|') !== 'Overview|Inputs') throw new Error(`Unexpected sheet tabs: ${tabs.join('|')}`);

  const columnHandle = window.getByRole('separator', { name: 'Resize column A', exact: true });
  const columnBefore = Number(await columnHandle.getAttribute('aria-valuenow'));
  const columnBox = await columnHandle.boundingBox();
  if (!columnBox) throw new Error('Column resize handle is not visible.');
  await window.mouse.move(columnBox.x + columnBox.width / 2, columnBox.y + columnBox.height / 2);
  await window.mouse.down();
  await window.mouse.move(columnBox.x + columnBox.width / 2 + 48, columnBox.y + columnBox.height / 2, { steps: 4 });
  await window.mouse.up();
  await window.waitForTimeout(250);

  const rowHandle = window.getByRole('separator', { name: 'Resize row 1', exact: true });
  const rowBefore = Number(await rowHandle.getAttribute('aria-valuenow'));
  const rowBox = await rowHandle.boundingBox();
  if (!rowBox) throw new Error('Row resize handle is not visible.');
  await window.mouse.move(rowBox.x + rowBox.width / 2, rowBox.y + rowBox.height / 2);
  await window.mouse.down();
  await window.mouse.move(rowBox.x + rowBox.width / 2, rowBox.y + rowBox.height / 2 + 24, { steps: 4 });
  await window.mouse.up();
  await window.waitForTimeout(250);

  const columnAfter = Number(await columnHandle.getAttribute('aria-valuenow'));
  const rowAfter = Number(await rowHandle.getAttribute('aria-valuenow'));
  if (columnAfter <= columnBefore) throw new Error(`Column did not resize: ${columnBefore} -> ${columnAfter}`);
  if (rowAfter <= rowBefore) throw new Error(`Row did not resize: ${rowBefore} -> ${rowAfter}`);
  await window.waitForTimeout(1_100);
  await window.screenshot({ path: output });
  const titleFont = await window.getByRole('gridcell', { name: 'A1', exact: true }).evaluate((element) => getComputedStyle(element).fontFamily);
  console.log(JSON.stringify({ title: await window.title(), total: total.trim(), tabs, titleFont, columnBefore, columnAfter, rowBefore, rowAfter, screenshot: output }));
  await stop(application);
  applicationStopped = true;

  let recoveryDialogs = 0;
  const recoveryApplication = await launch();
  recoveryApplication.on('window', (page) => page.on('dialog', async (dialog) => {
    recoveryDialogs += 1;
    await dialog.dismiss();
  }));
  try {
    const recoveryWindow = await recoveryApplication.firstWindow();
    await recoveryWindow.getByRole('grid', { name: /Overview spreadsheet grid/u }).waitFor({ timeout: 15_000 });
    await recoveryWindow.getByText('Recovered unsaved work', { exact: true }).waitFor({ timeout: 5_000 });
    const restoredColumn = Number(await recoveryWindow.getByRole('separator', { name: 'Resize column A', exact: true }).getAttribute('aria-valuenow'));
    const restoredRow = Number(await recoveryWindow.getByRole('separator', { name: 'Resize row 1', exact: true }).getAttribute('aria-valuenow'));
    if (recoveryDialogs !== 0) throw new Error(`Recovery opened ${recoveryDialogs} confirmation dialog(s).`);
    if (restoredColumn !== columnAfter || restoredRow !== rowAfter) throw new Error('Recovered dimensions did not match the resized sheet.');
    console.log(JSON.stringify({ recoveryDialogs, restoredColumn, restoredRow, recovery: 'silent' }));
  } finally {
    await stop(recoveryApplication);
  }
} finally {
  if (!applicationStopped) {
    await stop(application).catch(() => application.process().kill());
  }
}
