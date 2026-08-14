import * as electronUpdater from 'electron-updater';
import type { AppUpdater, ProgressInfo, UpdateInfo } from 'electron-updater';
import type { AppUpdateState } from '../shared/types';

type StateListener = (state: AppUpdateState) => void;

function updaterInstance(): AppUpdater {
  return electronUpdater.autoUpdater;
}

function checkedAt(): string {
  return new Date().toISOString();
}

export class AppUpdateManager {
  private readonly updater = updaterInstance();
  private state: AppUpdateState;
  private initialized = false;

  constructor(currentVersion: string, private readonly canCheck: boolean, private readonly onState: StateListener) {
    this.state = {
      currentVersion,
      phase: canCheck ? 'idle' : 'unavailable',
      canCheck,
      message: canCheck ? undefined : 'Update checks are available in installed builds.',
    };
  }

  initialize(): void {
    if (this.initialized || !this.canCheck) return;
    this.initialized = true;
    this.updater.autoDownload = false;
    this.updater.autoInstallOnAppQuit = true;
    this.updater.allowPrerelease = false;
    this.updater.logger = console;
    this.updater.on('checking-for-update', () => this.setState({ phase: 'checking', message: undefined, downloadPercent: undefined }));
    this.updater.on('update-available', (info: UpdateInfo) => this.setState({ phase: 'available', availableVersion: info.version, lastCheckedAt: checkedAt(), message: undefined, downloadPercent: undefined }));
    this.updater.on('update-not-available', () => this.setState({ phase: 'up-to-date', availableVersion: undefined, lastCheckedAt: checkedAt(), message: undefined, downloadPercent: undefined }));
    this.updater.on('download-progress', (progress: ProgressInfo) => this.setState({ phase: 'downloading', downloadPercent: Math.max(0, Math.min(100, progress.percent)) }));
    this.updater.on('update-downloaded', (info: UpdateInfo) => this.setState({ phase: 'downloaded', availableVersion: info.version, downloadPercent: 100, message: undefined }));
    this.updater.on('update-cancelled', () => this.setState({ phase: 'available', downloadPercent: undefined, message: 'The update download was cancelled.' }));
    this.updater.on('error', (error: Error) => {
      console.error('TXT Sheets update error', error);
      this.setState({ phase: 'error', downloadPercent: undefined, lastCheckedAt: checkedAt(), message: 'Unable to update TXT Sheets. Check your connection and try again.' });
    });
  }

  getState(): AppUpdateState {
    return { ...this.state };
  }

  async checkForUpdates(): Promise<AppUpdateState> {
    if (!this.canCheck || this.state.phase === 'checking' || this.state.phase === 'downloading' || this.state.phase === 'downloaded') return this.getState();
    this.setState({ phase: 'checking', message: undefined, downloadPercent: undefined });
    try {
      const result = await this.updater.checkForUpdates();
      if (this.getState().phase === 'checking') {
        this.setState(result?.isUpdateAvailable
          ? { phase: 'available', availableVersion: result.updateInfo.version, lastCheckedAt: checkedAt() }
          : { phase: 'up-to-date', availableVersion: undefined, lastCheckedAt: checkedAt() });
      }
    } catch (error) {
      console.error('TXT Sheets update check failed', error);
      this.setState({ phase: 'error', downloadPercent: undefined, lastCheckedAt: checkedAt(), message: 'Unable to check for updates. Check your connection and try again.' });
    }
    return this.getState();
  }

  async downloadUpdate(): Promise<AppUpdateState> {
    if (!this.canCheck || this.state.phase !== 'available') return this.getState();
    this.setState({ phase: 'downloading', downloadPercent: 0, message: undefined });
    try {
      await this.updater.downloadUpdate();
    } catch (error) {
      console.error('TXT Sheets update download failed', error);
      this.setState({ phase: 'error', downloadPercent: undefined, message: 'The update could not be downloaded. Check your connection and try again.' });
    }
    return this.getState();
  }

  installUpdate(): void {
    if (this.state.phase !== 'downloaded') throw new Error('No downloaded update is ready to install.');
    this.updater.quitAndInstall(false, true);
  }

  private setState(patch: Partial<AppUpdateState>): void {
    this.state = { ...this.state, ...patch };
    this.onState(this.getState());
  }
}
