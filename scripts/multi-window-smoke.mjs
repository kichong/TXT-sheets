import assert from 'node:assert/strict';
import { mkdtemp, mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import electronBinary from 'electron';
import { spawn } from 'node:child_process';
import { _electron as electron } from 'playwright';

const root = resolve('.');
const docs = JSON.parse(await readFile(resolve('package.json'), 'utf8')).name === 'txt-docs';
const packaged = process.argv.includes('--packaged');
const executablePath = packaged ? resolve('release/win-unpacked', docs ? 'TXT Docs.exe' : 'TXT Sheets.exe') : electronBinary;
const baseArgs = packaged ? [] : [root];
const output = resolve('output/multi-window');
await mkdir(output, { recursive: true });
const profile = await mkdtemp(resolve(output, 'profile-'));
const firstPath = resolve(profile, docs ? 'first.txt' : 'first.csv');
const openedPath = resolve(profile, docs ? 'opened.txt' : 'opened.csv');
const savedPath = resolve(profile, docs ? 'saved.txt' : 'saved.csv');
await writeFile(firstPath, 'original');
await writeFile(openedPath, 'opened');
let application = await electron.launch({ executablePath, args: [...baseArgs, `--user-data-dir=${profile}`, firstPath], cwd: root });
const api = docs ? 'documentsApi' : 'spreadsheet';
const selector = docs ? '.document-editor' : '[role="grid"]';
const errors = [];
application.on('window', (page) => page.on('pageerror', (error) => errors.push(error.message)));

async function waitWindows(count) {
  for (let attempt = 0; attempt < 100; attempt++) {
    const pages = application.windows();
    if (pages.length === count) {
      await Promise.all(pages.map((page) => page.locator(selector).waitFor()));
      return pages;
    }
    await new Promise((done) => setTimeout(done, 100));
  }
  throw new Error(`Expected ${count} windows; got ${application.windows().length}`);
}

async function edit(page, text) {
  if (docs) {
    await page.locator(selector).click();
    await page.keyboard.press('Control+End');
    await page.keyboard.type(text);
  } else {
    await page.getByRole('gridcell', { name: 'A1', exact: true }).dblclick();
    await page.getByRole('textbox', { name: 'Edit A1', exact: true }).fill(text);
    await page.keyboard.press('Enter');
  }
}

async function content(page) {
  return docs ? page.locator(selector).innerText() : page.getByRole('gridcell', { name: 'A1', exact: true }).innerText();
}

try {
  const first = await application.firstWindow();
  await first.locator(selector).waitFor();
  const firstId = await (await application.browserWindow(first)).evaluate((window) => window.id);
  await first.waitForFunction((docs) => (docs ? document.querySelector('.document-editor') : document.querySelector('[role="gridcell"][aria-label="A1"]'))?.textContent?.includes('original'), docs);
  await edit(first, ' first unsaved');
  await first.evaluate((api) => window[api].newWindow(), api);
  const [, second] = await waitWindows(2);
  const secondId = await (await application.browserWindow(second)).evaluate((window) => window.id);
  assert.ok(!(await content(second)).includes('original'));
  await edit(second, 'second unsaved');
  await second.waitForTimeout(1600);
  const firstRecovery = await first.evaluate(({ api, docs }) => docs ? window[api].readRecovery() : window[api].getRecovery(), { api, docs });
  const secondRecovery = await second.evaluate(({ api, docs }) => docs ? window[api].readRecovery() : window[api].getRecovery(), { api, docs });
  assert.ok(JSON.stringify(firstRecovery).includes('first unsaved'));
  assert.ok(JSON.stringify(secondRecovery).includes('second unsaved'));
  await application.evaluate(({ dialog }, path) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath: path });
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [globalThis.openPath] });
    globalThis.closeChoice = 2;
    globalThis.promptCount = 0;
    dialog.showMessageBoxSync = () => { globalThis.promptCount++; return globalThis.closeChoice; };
    dialog.showMessageBox = async () => { globalThis.promptCount++; return { response: globalThis.closeChoice }; };
  }, savedPath);
  await second.bringToFront();
  if (docs) await second.getByRole('button', { name: 'Save', exact: true }).click();
  else await second.keyboard.press('Control+s');
  await second.waitForTimeout(600);
  assert.ok((await readFile(savedPath, 'utf8')).includes('second unsaved'));
  assert.ok(await first.evaluate(({ api, docs }) => docs ? window[api].readRecovery() : window[api].getRecovery(), { api, docs }));
  assert.equal(await second.evaluate(({ api, docs }) => docs ? window[api].readRecovery() : window[api].getRecovery(), { api, docs }), null);
  const blocked = await second.evaluate(async (api) => { try { await window[api].installUpdate(null); return false; } catch (error) { return error.message.includes('Save'); } }, api);
  assert.equal(blocked, true, 'A dirty sibling must block restart-to-update');

  await application.evaluate(({ BrowserWindow }, id) => BrowserWindow.fromId(id).close(), firstId);
  await first.waitForTimeout(200);
  assert.equal(application.windows().length, 2, 'Cancel keeps the dirty file open');
  assert.equal(await application.evaluate(() => globalThis.promptCount), 1);
  await application.evaluate(({ BrowserWindow }, id) => BrowserWindow.fromId(id).close(), secondId);
  await waitWindows(1);
  assert.equal(await application.evaluate(() => globalThis.promptCount), 1, 'Saved sibling closes without prompting');

  await application.evaluate((_electron, path) => { globalThis.openPath = path; }, openedPath);
  await first.evaluate(({ api, docs }) => docs ? window[api].openDocument() : window[api].open(), { api, docs });
  const [, opened] = await waitWindows(2);
  await opened.waitForTimeout(300);
  assert.ok((await content(opened)).includes('opened'));
  assert.ok((await content(first)).includes('first unsaved'));
  const recent = await first.evaluate((api) => window[api].getRecentFiles(), api);
  await first.evaluate(({ api, id }) => window[api].openRecent(id), { api, id: recent.find((file) => file.displayName === (docs ? 'opened.txt' : 'opened.csv')).id });
  await waitWindows(3);

  await first.bringToFront();
  if (docs) await application.evaluate(({ Menu, BrowserWindow }, id) => {
    BrowserWindow.fromId(id).focus();
    Menu.getApplicationMenu().items[0].submenu.items[0].click();
  }, firstId);
  else await first.keyboard.press('Control+n');
  await waitWindows(4);
  const environment = { ...process.env };
  delete environment.ELECTRON_RUN_AS_NODE;
  const relaunch = spawn(executablePath, [...baseArgs, `--user-data-dir=${profile}`], { cwd: root, env: environment, windowsHide: true, stdio: 'ignore' });
  await waitWindows(5);
  await new Promise((done, reject) => { if (relaunch.exitCode !== null) return done(); relaunch.once('exit', done); relaunch.once('error', reject); });
  const external = spawn(executablePath, [...baseArgs, `--user-data-dir=${profile}`, openedPath], { cwd: root, env: environment, windowsHide: true, stdio: 'ignore' });
  const pages = await waitWindows(6);
  await pages.at(-1).waitForTimeout(300);
  assert.ok((await content(pages.at(-1))).includes('opened'));
  assert.ok((await content(first)).includes('first unsaved'));
  await new Promise((done, reject) => { if (external.exitCode !== null) return done(); external.once('exit', done); external.once('error', reject); });
  await application.evaluate(({ app }) => { setImmediate(() => app.exit(0)); });
  await new Promise((done) => application.process().exitCode !== null ? done() : application.process().once('exit', done));
  application = await electron.launch({ executablePath, args: [...baseArgs, `--user-data-dir=${profile}`], cwd: root });
  const recoveredPages = await waitWindows(2);
  const drafts = await Promise.all(recoveredPages.map((page) => page.evaluate(({ api, docs }) => docs ? window[api].readRecovery() : window[api].getRecovery(), { api, docs })));
  const recoveryIndex = drafts.findIndex((draft) => draft !== null);
  assert.notEqual(recoveryIndex, -1, 'The unsaved recovery draft must remain available after restart');
  const recovered = recoveredPages[recoveryIndex];
  const fresh = recoveredPages[1 - recoveryIndex];
  if (docs) await recovered.getByRole('button', { name: 'Restore draft', exact: true }).click();
  await recovered.waitForTimeout(300);
  assert.ok((await content(recovered)).includes('first unsaved'));
  assert.ok(!(await content(fresh)).includes('first unsaved'), 'Relaunch still provides a fresh file beside recovery');
  assert.deepEqual(errors, []);
  console.log(`${docs ? 'TXT Docs' : 'TXT Sheets'}: multiple windows, independent recovery/save/close, Open/recent, New menu/shortcut, real relaunch, Open with, and restart recovery passed.`);
} finally {
  await application.evaluate(({ app }) => app.exit(0)).catch(() => {});
}
