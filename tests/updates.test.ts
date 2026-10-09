import { describe, expect, it, vi } from 'vitest';
import { autoUpdater } from 'electron-updater';
import { AppUpdateManager } from '../src/main/updater';
import { presentUpdate } from '../src/shared/updates';

vi.mock('electron-updater', () => ({ autoUpdater: {
  on: vi.fn(), quitAndInstall: vi.fn(),
  autoInstallOnAppQuit: true, autoDownload: true,
} }));

it('requires explicit installation and keeps normal quit from installing a downloaded update', () => {
  const manager = new AppUpdateManager('0.2.7', true, () => undefined);
  manager.initialize();
  expect(autoUpdater.autoInstallOnAppQuit).toBe(false);
  expect(autoUpdater.autoDownload).toBe(false);
  expect(() => manager.installUpdate()).toThrow('No downloaded update');
  const downloaded = vi.mocked(autoUpdater.on).mock.calls.find(([event]) => event === 'update-downloaded')![1];
  (downloaded as (info: { version: string }) => void)({ version: '0.2.8' });
  expect(autoUpdater.quitAndInstall).not.toHaveBeenCalled();
  manager.installUpdate();
  expect(autoUpdater.quitAndInstall).toHaveBeenCalledWith(false, true);
});

describe('update presentation', () => {
  it('offers a user-confirmed download for available releases', () => {
    const presentation = presentUpdate({ currentVersion: '0.2.0', phase: 'available', canCheck: true, availableVersion: '0.2.1' });
    expect(presentation.action).toBe('download');
    expect(presentation.actionLabel).toBe('Download v0.2.1');
  });

  it('offers restart only after the update is downloaded', () => {
    const presentation = presentUpdate({ currentVersion: '0.2.0', phase: 'downloaded', canCheck: true, availableVersion: '0.2.1' });
    expect(presentation.action).toBe('install');
    expect(presentation.actionLabel).toBe('Restart to update');
  });
});
