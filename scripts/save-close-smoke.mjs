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
  await edit(run.page, 'discarded edit');
  await close(run.application, 1);
  await stopped(run.application);
  assert.doesNotMatch(await readFile(existing, 'utf8'), /discarded edit/);
} finally { await stop(run.application); }
console.log(`PASS ${docs ? 'Docs' : 'Sheets'}: Save, Cancel, cancelled Save As, failed save, successful Save As, and close without saving.`);
