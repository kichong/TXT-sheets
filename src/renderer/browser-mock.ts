import type { SpreadsheetApi } from '../shared/types';

// Browser-only development adapter. Electron always provides the real preload
// API; this keeps the renderer independently previewable for visual QA.
if (!window.spreadsheet) {
  const mock: SpreadsheetApi = {
    open: async () => null,
    openExternal: async () => null,
    cancelExternal: () => undefined,
    openRecent: async () => { throw new Error('Recent files are available in the Windows app.'); },
    save: async () => ({ status: 'saved', source: { id: 'browser-preview', displayName: 'Untitled workbook.xlsx', format: 'xlsx' }, recentFiles: [] }),
    saveAs: async () => ({ status: 'saved', source: { id: 'browser-preview', displayName: 'Untitled workbook.xlsx', format: 'xlsx' }, recentFiles: [] }),
    reportCompatibility: async () => undefined,
    getRecentFiles: async () => [],
    getRecovery: async () => null,
    writeRecovery: async () => undefined,
    clearRecovery: async () => undefined,
    getUpdateState: async () => ({ currentVersion: '0.2.0', phase: 'unavailable', canCheck: false, message: 'Update checks are available in installed builds.' }),
    checkForUpdates: async () => ({ currentVersion: '0.2.0', phase: 'unavailable', canCheck: false }),
    downloadUpdate: async () => ({ currentVersion: '0.2.0', phase: 'unavailable', canCheck: false }),
    installUpdate: async () => undefined,
    setDirty: () => undefined,
    onCommand: () => () => undefined,
    onExternalFile: () => () => undefined,
    onUpdateState: () => () => undefined,
  };
  window.spreadsheet = mock;
}
