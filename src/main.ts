import { app, BrowserWindow, dialog, ipcMain, Menu, type MenuItemConstructorOptions } from 'electron';
import { readFile } from 'node:fs/promises';
import { basename, extname, join } from 'node:path';
import { AppStorage } from './main/storage';
import { findLaunchWorkbookPath } from './main/launch-files';
import { AppUpdateManager } from './main/updater';
import { exportWorkbook, formatForPath, importWorkbook } from './main/workbook-io';
import type { AppCommand, OpenResult, SaveResult, WorkbookDocument, WorkbookFormat } from './shared/types';

let mainWindow: BrowserWindow | null = null;
let storage: AppStorage;
let updateManager: AppUpdateManager;
let dirty = false;
let forceClose = false;
const pendingExternalPaths: string[] = [];

function queueExternalWorkbook(args: string[]): void {
  const path = findLaunchWorkbookPath(args);
  if (path && !pendingExternalPaths.includes(path)) pendingExternalPaths.push(path);
}

queueExternalWorkbook(process.argv.slice(1));

function sendCommand(command: AppCommand): void {
  mainWindow?.webContents.send('app:command', command);
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
          click: () => void dialog.showMessageBox(mainWindow!, {
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
  dirty = false;
  return { workbook, recentFiles };
}

async function chooseOpenPath(): Promise<string | null> {
  const result = await dialog.showOpenDialog(mainWindow!, {
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

async function chooseSavePath(workbook: WorkbookDocument): Promise<string | null> {
  const preferred = workbook.source?.format ?? 'xlsx';
  const result = await dialog.showSaveDialog(mainWindow!, {
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

async function saveToPath(workbook: WorkbookDocument, path: string): Promise<SaveResult> {
  const format = formatForPath(path);
  await storage.atomicWrite(path, await exportWorkbook(workbook, path));
  const source = await storage.sourceFor(path, format);
  const recentFiles = await storage.remember(path, format);
  await storage.clearRecovery();
  dirty = false;
  mainWindow?.setTitle(`${basename(path)} — TXT Sheets`);
  return { status: 'saved', source, recentFiles };
}

function validateWorkbook(value: unknown): asserts value is WorkbookDocument {
  if (!value || typeof value !== 'object' || (value as WorkbookDocument).schemaVersion !== 1 || !Array.isArray((value as WorkbookDocument).sheets)) {
    throw new Error('The workbook data was invalid.');
  }
}

function installIpcHandlers(): void {
  ipcMain.handle('workbooks:open', async () => {
    const path = await chooseOpenPath();
    return path ? openPath(path) : null;
  });
  ipcMain.handle('workbooks:open-external', async () => {
    const path = pendingExternalPaths.shift();
    if (!path) return null;
    try { return await openPath(path); }
    catch (error) {
      if (process.env.TXT_SHEETS_SMOKE_LOG) console.error(error);
      throw error;
    }
  });
  ipcMain.on('workbooks:cancel-external', () => { pendingExternalPaths.shift(); });
  ipcMain.handle('workbooks:open-recent', async (_event, id: unknown) => {
    if (typeof id !== 'string') throw new Error('Invalid recent file.');
    const path = storage.getPath(id);
    if (!path) throw new Error('This recent file is no longer available.');
    return openPath(path);
  });
  ipcMain.handle('workbooks:save', async (_event, value: unknown) => {
    validateWorkbook(value);
    const path = value.source ? storage.getPath(value.source.id) : null;
    return path ? saveToPath(value, path) : saveToPathOrCancel(value);
  });
  ipcMain.handle('workbooks:save-as', async (_event, value: unknown) => {
    validateWorkbook(value);
    return saveToPathOrCancel(value);
  });
  ipcMain.handle('workbooks:recent', () => storage.getRecentFiles());
  ipcMain.handle('workbooks:recovery', () => storage.getRecovery());
  ipcMain.handle('workbooks:write-recovery', async (_event, value: unknown) => {
    validateWorkbook(value);
    await storage.writeRecovery(value);
  });
  ipcMain.handle('workbooks:clear-recovery', () => storage.clearRecovery());
  ipcMain.handle('updates:get-state', () => updateManager.getState());
  ipcMain.handle('updates:check', () => updateManager.checkForUpdates());
  ipcMain.handle('updates:download', () => updateManager.downloadUpdate());
  ipcMain.handle('updates:install', () => {
    if (dirty) throw new Error('Save your workbook before restarting to install the update.');
    forceClose = true;
    updateManager.installUpdate();
  });
  ipcMain.on('workbooks:dirty', (_event, value: unknown) => { dirty = value === true; });
}

async function saveToPathOrCancel(workbook: WorkbookDocument): Promise<SaveResult> {
  const path = await chooseSavePath(workbook);
  return path ? saveToPath(workbook, path) : { status: 'canceled', recentFiles: storage.getRecentFiles() };
}

function createWindow(): void {
  mainWindow = new BrowserWindow({
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
  const developmentUrl = process.env.VITE_DEV_SERVER_URL;
  if (developmentUrl) void mainWindow.loadURL(developmentUrl);
  else void mainWindow.loadFile(join(__dirname, '../renderer/index.html'));
  mainWindow.on('close', (event) => {
    if (!dirty || forceClose) return;
    event.preventDefault();
    void dialog.showMessageBox(mainWindow!, {
      type: 'warning', title: 'Unsaved changes', message: 'Close without saving?',
      detail: 'Your latest work is kept in recovery until you open TXT Sheets again.',
      buttons: ['Keep editing', 'Close'], defaultId: 0, cancelId: 0,
    }).then(({ response }) => {
      if (response === 1) { forceClose = true; mainWindow?.close(); }
    });
  });
  mainWindow.on('closed', () => { mainWindow = null; });
}

const hasLock = app.requestSingleInstanceLock();
if (!hasLock) app.quit();
else {
  app.on('second-instance', (_event, commandLine) => {
    queueExternalWorkbook(commandLine);
    if (mainWindow) {
      if (mainWindow.isMinimized()) mainWindow.restore();
      mainWindow.focus();
      mainWindow.webContents.send('workbooks:external-ready');
    }
  });
  void app.whenReady().then(async () => {
    app.setAppUserModelId('com.txtsheets.app');
    storage = new AppStorage(app.getPath('userData'));
    await storage.initialize();
    updateManager = new AppUpdateManager(app.getVersion(), app.isPackaged, (state) => {
      mainWindow?.webContents.send('app:update-state', state);
    });
    updateManager.initialize();
    installIpcHandlers();
    createApplicationMenu();
    createWindow();
    mainWindow?.webContents.once('did-finish-load', () => {
      if (pendingExternalPaths.length) mainWindow?.webContents.send('workbooks:external-ready');
    });
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
