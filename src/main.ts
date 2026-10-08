import { app, BrowserWindow, dialog, ipcMain, Menu, shell, type MenuItemConstructorOptions } from 'electron';
import { randomUUID } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import { basename, extname, join } from 'node:path';
import { AppStorage } from './main/storage';
import { findLaunchWorkbookPath } from './main/launch-files';
import { AppUpdateManager } from './main/updater';
import { exportWorkbook, formatForPath, importWorkbook } from './main/workbook-io';
import { buildCompatibilityReportUrl } from './shared/compatibility-report';
import type { AppCommand, CompatibilityReportRequest, OpenResult, SaveResult, WorkbookDocument, WorkbookFormat } from './shared/types';

interface FileWindow {
  window: BrowserWindow;
  dirty: boolean;
  forceClose: boolean;
  closePromptOpen: boolean;
  recoveryKey: string;
  externalPath?: string;
  initialResult?: OpenResult;
}
const windows = new Map<number, FileWindow>();
let storage: AppStorage;
let updateManager: AppUpdateManager;

function focusedWindow(): BrowserWindow | null {
  return BrowserWindow.getFocusedWindow() ?? [...windows.values()].at(-1)?.window ?? null;
}

function fileWindow(event: Electron.IpcMainEvent | Electron.IpcMainInvokeEvent): FileWindow {
  const state = windows.get(event.sender.id);
  if (!state) throw new Error('The file window is no longer available.');
  return state;
}

function anyDirty(): boolean {
  return [...windows.values()].some((state) => state.dirty);
}

const pendingExternalPaths: string[] = [];

function queueExternalWorkbook(args: string[]): void {
  const path = findLaunchWorkbookPath(args);
  if (path && !pendingExternalPaths.includes(path)) pendingExternalPaths.push(path);
}

queueExternalWorkbook(process.argv.slice(1));

function sendCommand(command: AppCommand): void {
  focusedWindow()?.webContents.send('app:command', command);
}

function createApplicationMenu(): void {
  const template: MenuItemConstructorOptions[] = [
    {
      label: 'File',
      submenu: [
        { label: 'New', accelerator: 'Ctrl+N', click: () => sendCommand('new') },
        { label: 'Open…', accelerator: 'Ctrl+O', click: () => sendCommand('open') },
        { type: 'separator' },
        { label: 'Save', accelerator: 'Ctrl+S', click: () => sendCommand('save') },
        { label: 'Save As…', accelerator: 'Ctrl+Shift+S', click: () => sendCommand('save-as') },
        { type: 'separator' },
        { role: 'quit' },
      ],
    },
    {
      label: 'Edit',
      submenu: [
        { label: 'Undo', accelerator: 'Ctrl+Z', click: () => sendCommand('undo') },
        { label: 'Redo', accelerator: 'Ctrl+Y', click: () => sendCommand('redo') },
        { type: 'separator' },
        { role: 'cut' }, { role: 'copy' }, { role: 'paste' }, { role: 'selectAll' },
        { type: 'separator' },
        { label: 'Find…', accelerator: 'Ctrl+F', click: () => sendCommand('find') },
      ],
    },
    {
      label: 'View',
      submenu: [{ role: 'resetZoom' }, { role: 'zoomIn' }, { role: 'zoomOut' }, { type: 'separator' }, { role: 'togglefullscreen' }],
    },
    {
      label: 'Help',
      submenu: [
        { label: 'Check for Updates…', click: () => void updateManager.checkForUpdates() },
        { type: 'separator' },
        {
          label: `About TXT Sheets v${app.getVersion()}`,
          click: () => void dialog.showMessageBox(focusedWindow()!, {
            type: 'info', title: 'About TXT Sheets', message: `TXT Sheets v${app.getVersion()}`,
            detail: 'A free, open-source spreadsheet application.\n\nLicensed under Apache License 2.0.', buttons: ['OK'],
          }),
        },
      ],
    },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

async function openPath(path: string): Promise<OpenResult> {
  const format = formatForPath(path);
  const source = await storage.sourceFor(path, format);
  const workbook = await importWorkbook(new Uint8Array(await readFile(path)), path, source);
  const recentFiles = await storage.remember(path, format);
  return { workbook, recentFiles };
}

async function chooseOpenPath(owner: BrowserWindow): Promise<string | null> {
  const result = await dialog.showOpenDialog(owner, {
    title: 'Open a spreadsheet',
    properties: ['openFile'],
    filters: [
      { name: 'Supported spreadsheets', extensions: ['xlsx', 'csv', 'tsv'] },
      { name: 'Excel workbooks', extensions: ['xlsx'] },
      { name: 'Delimited text', extensions: ['csv', 'tsv'] },
    ],
  });
  return result.canceled ? null : result.filePaths[0] ?? null;
}

function safeTitle(title: string): string {
  return title.replace(/[<>:"/\\|?*\u0000-\u001F]/gu, '').trim() || 'Untitled workbook';
}

async function chooseSavePath(workbook: WorkbookDocument, owner: BrowserWindow): Promise<string | null> {
  const preferred = workbook.source?.format ?? 'xlsx';
  const result = await dialog.showSaveDialog(owner, {
    title: 'Save spreadsheet as',
    defaultPath: `${safeTitle(workbook.title)}.${preferred}`,
    filters: [
      { name: 'Excel workbook', extensions: ['xlsx'] },
      { name: 'Comma-separated values', extensions: ['csv'] },
      { name: 'Tab-separated values', extensions: ['tsv'] },
    ],
  });
  return result.canceled ? null : result.filePath ?? null;
}

async function saveToPath(workbook: WorkbookDocument, path: string, state: FileWindow): Promise<SaveResult> {
  const format = formatForPath(path);
  await storage.atomicWrite(path, await exportWorkbook(workbook, path));
  const source = await storage.sourceFor(path, format);
  const recentFiles = await storage.remember(path, format);
  await storage.clearRecovery(state.recoveryKey);
  state.dirty = false;
  state.window.setTitle(`${basename(path)} — TXT Sheets`);
  return { status: 'saved', source, recentFiles };
}

function validateWorkbook(value: unknown): asserts value is WorkbookDocument {
  if (!value || typeof value !== 'object' || (value as WorkbookDocument).schemaVersion !== 1 || !Array.isArray((value as WorkbookDocument).sheets)) {
    throw new Error('The workbook data was invalid.');
  }
}

function validateCompatibilityReport(value: unknown): asserts value is CompatibilityReportRequest {
  const report = value as Partial<CompatibilityReportRequest> | null;
  if (!report || !['xlsx', 'csv', 'tsv', 'unsaved'].includes(String(report.sourceFormat)) || !Array.isArray(report.issues) || !report.issues.length || report.issues.length > 100) {
    throw new Error('The compatibility report was invalid.');
  }
  if (report.issues.some((issue) => !issue || typeof issue.feature !== 'string' || typeof issue.detail !== 'string' || issue.feature.length > 300 || issue.detail.length > 2000)) {
    throw new Error('The compatibility report was invalid.');
  }
}

function installIpcHandlers(): void {
  ipcMain.handle('workbooks:new', () => { createWindow(); });
  ipcMain.handle('workbooks:open', async (event) => {
    const path = await chooseOpenPath(fileWindow(event).window);
    if (path) createWindow(undefined, await openPath(path));
    return null;
  });
  ipcMain.handle('workbooks:open-external', async (event) => {
    const state = fileWindow(event);
    const result = state.initialResult;
    state.initialResult = undefined;
    if (result) return result;
    const path = state.externalPath;
    state.externalPath = undefined;
    return path ? openPath(path) : null;
  });
  ipcMain.on('workbooks:cancel-external', (event) => {
    fileWindow(event).externalPath = undefined;
  });
  ipcMain.handle('workbooks:open-recent', async (_event, id: unknown) => {
    if (typeof id !== 'string') throw new Error('Invalid recent file identifier.');
    const path = storage.getPath(id);
    if (!path) throw new Error('This recent file is no longer available.');
    createWindow(undefined, await openPath(path));
    return null;
  });
  ipcMain.handle('workbooks:save', async (event, value: unknown) => {
    validateWorkbook(value);
    const path = value.source ? storage.getPath(value.source.id) : null;
    return path ? saveToPath(value, path, fileWindow(event)) : saveToPathOrCancel(value, fileWindow(event));
  });
  ipcMain.handle('workbooks:save-as', async (event, value: unknown) => {
    validateWorkbook(value);
    return saveToPathOrCancel(value, fileWindow(event));
  });
  ipcMain.handle('workbooks:report-compatibility', async (_event, value: unknown) => {
    validateCompatibilityReport(value);
    await shell.openExternal(buildCompatibilityReportUrl({
      ...value,
      appName: 'TXT Sheets',
      appVersion: app.getVersion(),
      operatingSystem: process.platform,
      repositoryUrl: 'https://github.com/kichong/TXT-sheets',
    }));
  });
  ipcMain.handle('workbooks:recent', () => storage.getRecentFiles());
  ipcMain.handle('workbooks:recovery', (event) => storage.getRecovery(fileWindow(event).recoveryKey));
  ipcMain.handle('workbooks:write-recovery', async (event, value: unknown) => {
    validateWorkbook(value);
    await storage.writeRecovery(value, fileWindow(event).recoveryKey);
  });
  ipcMain.handle('workbooks:clear-recovery', (event) => storage.clearRecovery(fileWindow(event).recoveryKey));
  ipcMain.handle('updates:get-state', () => updateManager.getState());
  ipcMain.handle('updates:check', () => updateManager.checkForUpdates());
  ipcMain.handle('updates:download', () => updateManager.downloadUpdate());
  ipcMain.handle('updates:install', () => {
    if (anyDirty()) throw new Error('Save all open workbooks before restarting to install the update.');
    for (const state of windows.values()) state.forceClose = true;
    try { updateManager.installUpdate(); }
    catch (error) {
      for (const state of windows.values()) state.forceClose = false;
      throw error;
    }
  });
  ipcMain.on('workbooks:dirty', (event, value: unknown) => {
    fileWindow(event).dirty = value === true;
  });
  ipcMain.on('workbooks:close-after-save', (event) => {
    const state = fileWindow(event);
    if (state.dirty) return;
    state.forceClose = true;
    state.window.close();
  });
}

async function saveToPathOrCancel(workbook: WorkbookDocument, state: FileWindow): Promise<SaveResult> {
  const path = await chooseSavePath(workbook, state.window);
  return path ? saveToPath(workbook, path, state) : { status: 'canceled', recentFiles: storage.getRecentFiles() };
}

function createWindow(externalPath?: string, initialResult?: OpenResult, recoveryKey: string = randomUUID()): void {
  const window = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 760,
    minHeight: 520,
    backgroundColor: '#f5f6f8',
    title: 'TXT Sheets',
    icon: app.isPackaged ? join(process.resourcesPath, 'icon.ico') : join(process.cwd(), 'build', 'icon.ico'),
    webPreferences: {
      preload: join(__dirname, '../preload/preload.cjs'),
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: true,
    },
  });
  const state: FileWindow = { window, dirty: false, forceClose: false, closePromptOpen: false, recoveryKey, externalPath, initialResult };
  const windowId = window.webContents.id;
  windows.set(windowId, state);

  const developmentUrl = process.env.VITE_DEV_SERVER_URL;
  if (developmentUrl) void window.loadURL(developmentUrl);
  else void window.loadFile(join(__dirname, '../renderer/index.html'));
  window.on('close', (event) => {
    if (!state.dirty || state.forceClose) return;
    event.preventDefault();
    if (state.closePromptOpen) return;
    state.closePromptOpen = true;
    void dialog.showMessageBox(window, {
      type: 'warning', title: 'Unsaved changes', message: 'Save changes before closing?',
      buttons: ['Save', 'Close without saving', 'Cancel'], defaultId: 0, cancelId: 2, noLink: true,
    }).then(async ({ response }) => {
      if (response === 0) window.webContents.send('app:command', 'save-and-close');
      else if (response === 1) {
        await storage.clearRecovery(state.recoveryKey);
        state.forceClose = true;
        window?.close();
      }
    }).finally(() => { state.closePromptOpen = false; });
  });
  window.on('closed', () => { windows.delete(windowId); });
}

const hasLock = app.requestSingleInstanceLock();
if (!hasLock) app.quit();
else {
  app.on('second-instance', (_event, commandLine) => {
    const path = findLaunchWorkbookPath(commandLine) ?? undefined;
    if (windows.size) createWindow(path);
    else {
      // A launch can arrive while storage is still initializing.
      pendingExternalPaths.push(path ?? '');
    }
  });
  void app.whenReady().then(async () => {
    app.setAppUserModelId('com.txtsheets.app');
    storage = new AppStorage(app.getPath('userData'));
    await storage.initialize();
    updateManager = new AppUpdateManager(app.getVersion(), app.isPackaged, (state) => {
      for (const file of windows.values()) file.window.webContents.send('app:update-state', state);
    });
    updateManager.initialize();
    installIpcHandlers();
    createApplicationMenu();
    const recoveryKeys = await storage.getRecoveryKeys();
    for (const key of recoveryKeys) createWindow(undefined, undefined, key);
    const launchPaths = pendingExternalPaths.splice(0);
    if (launchPaths.length) {
      for (const path of launchPaths) createWindow(path || undefined);
    } else {
      createWindow();
    }
    if (app.isPackaged) {
      const firstCheck = setTimeout(() => void updateManager.checkForUpdates(), 4_000);
      firstCheck.unref();
      const recurringCheck = setInterval(() => void updateManager.checkForUpdates(), 6 * 60 * 60 * 1_000);
      recurringCheck.unref();
    }
  });
}

app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });
