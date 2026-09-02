import { contextBridge, ipcRenderer } from 'electron';
import type { AppCommand, AppUpdateState, CompatibilityReportRequest, SpreadsheetApi, UpdatePhase, WorkbookDocument } from './shared/types';

const UPDATE_PHASES = new Set<UpdatePhase>(['idle', 'checking', 'available', 'downloading', 'downloaded', 'up-to-date', 'error', 'unavailable']);

function parseUpdateState(value: unknown): AppUpdateState {
  if (!value || typeof value !== 'object') throw new Error('Invalid update state.');
  const state = value as Partial<AppUpdateState>;
  if (typeof state.currentVersion !== 'string' || typeof state.phase !== 'string' || !UPDATE_PHASES.has(state.phase as UpdatePhase) || typeof state.canCheck !== 'boolean') {
    throw new Error('Invalid update state.');
  }
  return state as AppUpdateState;
}

const api: SpreadsheetApi = {
  open: () => ipcRenderer.invoke('workbooks:open'),
  openExternal: () => ipcRenderer.invoke('workbooks:open-external'),
  cancelExternal: () => ipcRenderer.send('workbooks:cancel-external'),
  openRecent: (id) => ipcRenderer.invoke('workbooks:open-recent', id),
  save: (workbook) => ipcRenderer.invoke('workbooks:save', workbook),
  saveAs: (workbook) => ipcRenderer.invoke('workbooks:save-as', workbook),
  reportCompatibility: (request: CompatibilityReportRequest) => ipcRenderer.invoke('workbooks:report-compatibility', request),
  getRecentFiles: () => ipcRenderer.invoke('workbooks:recent'),
  getRecovery: () => ipcRenderer.invoke('workbooks:recovery'),
  writeRecovery: (workbook: WorkbookDocument) => ipcRenderer.invoke('workbooks:write-recovery', workbook),
  clearRecovery: () => ipcRenderer.invoke('workbooks:clear-recovery'),
  getUpdateState: async () => parseUpdateState(await ipcRenderer.invoke('updates:get-state')),
  checkForUpdates: async () => parseUpdateState(await ipcRenderer.invoke('updates:check')),
  downloadUpdate: async () => parseUpdateState(await ipcRenderer.invoke('updates:download')),
  installUpdate: () => ipcRenderer.invoke('updates:install'),
  setDirty: (dirty) => ipcRenderer.send('workbooks:dirty', dirty),
  onCommand: (callback) => {
    const listener = (_event: Electron.IpcRendererEvent, command: AppCommand) => callback(command);
    ipcRenderer.on('app:command', listener);
    return () => ipcRenderer.removeListener('app:command', listener);
  },
  onExternalFile: (callback) => {
    const listener = () => callback();
    ipcRenderer.on('workbooks:external-ready', listener);
    return () => ipcRenderer.removeListener('workbooks:external-ready', listener);
  },
  onUpdateState: (callback) => {
    const listener = (_event: Electron.IpcRendererEvent, state: unknown) => callback(parseUpdateState(state));
    ipcRenderer.on('app:update-state', listener);
    return () => ipcRenderer.removeListener('app:update-state', listener);
  },
};

contextBridge.exposeInMainWorld('spreadsheet', api);
