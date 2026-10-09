import assert from 'node:assert/strict';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { _electron as electron } from 'playwright';

const docs = process.argv.includes('--docs');
const packaged = process.argv.includes('--packaged');
const root = docs ? resolve('../TXT-docs') : resolve('.');
const output = resolve(root, 'output/playwright');
await mkdir(output, { recursive: true });
const directory = await mkdtemp(resolve(output, 'save-close-'));
const processes = new WeakMap();

async function launch(file) {
  const userData = await mkdtemp(resolve(directory, 'profile-'));
  const application = await electron.launch({
    ...(packaged ? { executablePath: resolve(root, 'release/win-unpacked', docs ? 'TXT Docs.exe' : 'TXT Sheets.exe') } : {}),
    args: [...(packaged ? [] : [root]), `--user-data-dir=${userData}`, ...(file ? [file] : [])],
    cwd: root,
  });
  processes.set(application, application.process());
  const page = await application.firstWindow();
  await page.locator(docs ? '.tiptap' : '[role="grid"]').waitFor();
  await application.evaluate(({ dialog }) => {
    globalThis.closeChoice = 2;
    globalThis.closeDialogs = [];
    const answer = (options) => {
      globalThis.closeDialogs.push(options.buttons);
      return globalThis.closeChoice;
    };
    dialog.showMessageBoxSync = (_window, options) => answer(options);
    dialog.showMessageBox = async (_window, options) => ({ response: answer(options), checkboxChecked: false });
    dialog.showSaveDialog = async () => globalThis.saveAnswer ?? { canceled: true };
  });
  return { application, page };
}

async function edit(page, text) {
  if (docs) {
    await page.locator('.tiptap').click();
    await page.keyboard.press('Control+End');
    await page.keyboard.type(text);
  } else {
    await page.getByRole('gridcell', { name: 'A1', exact: true }).dblclick();
    await page.getByRole('textbox', { name: 'Edit A1', exact: true }).fill(text);
  }
  await page.waitForTimeout(150);
}

async function close(application, choice, saveAnswer) {
  await application.evaluate(({ BrowserWindow }, { choice, saveAnswer }) => {
    globalThis.closeChoice = choice;
    globalThis.saveAnswer = saveAnswer;
    BrowserWindow.getAllWindows()[0].close();
  }, { choice, saveAnswer });
}

async function stopped(application) {
  if (processes.get(application).exitCode !== null) return;
  await Promise.race([
    new Promise((resolveExit) => processes.get(application).once('exit', resolveExit)),
    new Promise((_, reject) => setTimeout(() => reject(new Error('Save did not close the app')), 15000).unref()),
  ]);
}

async function stop(application) {
  if (processes.get(application).exitCode === null) {
    await application.evaluate(({ app }) => { setImmediate(() => app.exit(0)); }).catch(() => {});
    await stopped(application);
  }
}

const existing = resolve(directory, docs ? 'existing.txt' : 'existing.csv');
await writeFile(existing, 'original');
let run = await launch(existing);
try {
  // Wait for the launch file, then edit without committing the spreadsheet cell.
  await run.page.waitForFunction(() => document.body.textContent.includes('original'));
  await edit(run.page, 'saved edit');
  await close(run.application, 2);
  assert.deepEqual(await run.application.evaluate(() => globalThis.closeDialogs.at(-1)), ['Save', 'Close without saving', 'Cancel']);
  assert.equal(run.page.isClosed(), false);
  await close(run.application, 0);
  await stopped(run.application);
  assert.match(await readFile(existing, 'utf8'), /saved edit/);
} finally { await stop(run.application); }

run = await launch();
try {
  await edit(run.page, 'new content');
  await close(run.application, 0, { canceled: true });
  await run.page.waitForTimeout(300);
  assert.equal(run.page.isClosed(), false);
  const failedPath = resolve(directory, docs ? 'fail.txt' : 'fail.csv');
  await mkdir(failedPath);
  await close(run.application, 0, { canceled: false, filePath: failedPath });
  await run.page.waitForTimeout(700);
  assert.equal(run.page.isClosed(), false);
  const saved = resolve(directory, docs ? 'new.txt' : 'new.csv');
  await close(run.application, 0, { canceled: false, filePath: saved });
  await stopped(run.application);
  assert.match(await readFile(saved, 'utf8'), /new content/);
} finally { await stop(run.application); }

run = await launch(existing);
try {
  await run.page.waitForFunction(() => document.body.textContent.includes('original') || document.body.textContent.includes('saved edit'));
  await edit(run.page, 'before delayed save');
  if (!docs) await run.page.keyboard.press('Enter');
  await run.application.evaluate(({ ipcMain }, channel) => {
    const original = ipcMain._invokeHandlers.get(channel);
    globalThis.saveCalls = 0;
    ipcMain.removeHandler(channel);
    ipcMain.handle(channel, async (...args) => {
      globalThis.saveCalls++;
      await new Promise((resolve) => { globalThis.releaseSave = resolve; });
      return original(...args);
    });
  }, docs ? 'documents:save' : 'workbooks:save');
  await run.application.evaluate(({ BrowserWindow }) => {
    BrowserWindow.getAllWindows()[0].webContents.send('app:command', 'save-and-close');
  });
  for (let attempt = 0; attempt < 100; attempt++) {
    if (await run.application.evaluate(() => typeof globalThis.releaseSave === 'function')) break;
    await run.page.waitForTimeout(50);
  }
  assert.equal(await run.application.evaluate(() => typeof globalThis.releaseSave), 'function');
  await run.application.evaluate(({ BrowserWindow }) => {
    BrowserWindow.getAllWindows()[0].webContents.send('app:command', 'save');
  });
  await edit(run.page, 'latest edit during save');
  await run.application.evaluate(() => globalThis.releaseSave());
  await run.page.getByText('Saved the earlier changes. Your latest edits still need saving.', { exact: true }).waitFor();
  assert.equal(run.page.isClosed(), false, 'Newer edits must keep Save and Close open');
  assert.equal(await run.application.evaluate(() => globalThis.saveCalls), 1, 'Repeated saves must not overlap');
  assert.match(await readFile(existing, 'utf8'), /before delayed save/);
  assert.doesNotMatch(await readFile(existing, 'utf8'), /latest edit during save/);
  if (docs) assert.match(await run.page.locator('.tiptap').innerText(), /latest edit during save/);
  else {
    const input = run.page.getByRole('textbox', { name: 'Edit A1', exact: true });
    // Restoring grid focus can commit the unfinished edit; either state must keep its value.
    const value = await input.count() ? await input.inputValue() : await run.page.getByRole('gridcell', { name: 'A1', exact: true }).innerText();
    assert.match(value, /latest edit during save/);
  }
  const recovery = await run.page.evaluate((docs) => docs ? window.documentsApi.readRecovery() : window.spreadsheet.getRecovery(), docs);
  assert.match(JSON.stringify(recovery), /latest edit during save/);
  await run.application.evaluate(({ ipcMain }, channel) => {
    // The next save can run without a delay.
    const delayed = ipcMain._invokeHandlers.get(channel);
    ipcMain.removeHandler(channel);
    ipcMain.handle(channel, (...args) => {
      const operation = delayed(...args);
      globalThis.releaseSave();
      return operation;
    });
  }, docs ? 'documents:save' : 'workbooks:save');
  await close(run.application, 0);
  await stopped(run.application);
  assert.match(await readFile(existing, 'utf8'), /latest edit during save/);
} finally { await stop(run.application); }

run = await launch(existing);
try {
  await edit(run.page, 'discarded edit');
  await close(run.application, 1);
  await stopped(run.application);
  assert.doesNotMatch(await readFile(existing, 'utf8'), /discarded edit/);
} finally { await stop(run.application); }
if (docs) {
  run = await launch();
  try {
    await run.page.locator('.tiptap').click();
    await run.page.keyboard.type('Al');
    await run.page.keyboard.press('Control+b');
    await run.page.keyboard.type('pha');
    await run.page.keyboard.press('Control+b');
    assert.equal(await run.page.locator('.tiptap strong').innerText(), 'pha');
    await run.page.keyboard.press('Control+f');
    await run.page.getByRole('textbox', { name: 'Find in document', exact: true }).fill('alpha');
    await run.page.getByText('1 of 1', { exact: true }).waitFor();
    assert.equal((await run.page.locator('.document-search-match.is-current').allTextContents()).join(''), 'Alpha');
    await run.page.keyboard.press('Escape');
  } finally { await stop(run.application); }
} else {
  run = await launch(existing);
  try {
    await run.page.waitForFunction(() => document.body.textContent.includes('latest edit during save'));
    const original = await readFile(existing, 'utf8');
    await edit(run.page, 'saved copy edit');
    await run.page.keyboard.press('Enter');
    const copy = resolve(directory, 'undo-copy.csv');
    await run.application.evaluate(({ dialog, BrowserWindow }, path) => {
      dialog.showSaveDialog = async () => ({ canceled: false, filePath: path });
      BrowserWindow.getAllWindows()[0].webContents.send('app:command', 'save-as');
    }, copy);
    await run.page.getByText('Saved as undo-copy.csv', { exact: true }).waitFor();
    await run.page.getByRole('button', { name: 'Undo', exact: true }).click();
    await run.page.getByRole('button', { name: 'Save', exact: true }).click();
    await run.page.getByText('Saved undo-copy.csv', { exact: true }).waitFor();
    assert.equal(await readFile(existing, 'utf8'), original, 'Undo after Save As must not switch back to the original file');
    assert.match(await readFile(copy, 'utf8'), /latest edit during save/);
    assert.doesNotMatch(await readFile(copy, 'utf8'), /saved copy edit/);
  } finally { await stop(run.application); }
}
console.log(`PASS ${docs ? 'Docs' : 'Sheets'}: save/close, cancelled/failed saves, newer edits during save, overlapping-save guard, recovery, and discard.`);
