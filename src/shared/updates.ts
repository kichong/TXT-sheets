import type { AppUpdateState } from './types';

export type UpdateAction = 'check' | 'download' | 'install' | null;

export interface UpdatePresentation {
  action: UpdateAction;
  actionLabel: string;
  detail: string;
  busy: boolean;
}

export function presentUpdate(state: AppUpdateState): UpdatePresentation {
  const availableVersion = state.availableVersion ? `v${state.availableVersion}` : 'the new version';
  switch (state.phase) {
    case 'checking':
      return { action: null, actionLabel: 'Checking…', detail: 'Checking GitHub for updates…', busy: true };
    case 'available':
      return { action: 'download', actionLabel: `Download ${availableVersion}`, detail: `TXT Sheets ${availableVersion} is available.`, busy: false };
    case 'downloading': {
      const percent = Math.round(state.downloadPercent ?? 0);
      return { action: null, actionLabel: `Downloading ${percent}%`, detail: `Downloading ${availableVersion}…`, busy: true };
    }
    case 'downloaded':
      return { action: 'install', actionLabel: 'Restart to update', detail: `TXT Sheets ${availableVersion} is ready to install.`, busy: false };
    case 'up-to-date':
      return { action: 'check', actionLabel: 'Check again', detail: `TXT Sheets v${state.currentVersion} is up to date.`, busy: false };
    case 'error':
      return { action: state.canCheck ? 'check' : null, actionLabel: state.canCheck ? 'Try again' : 'Updates unavailable', detail: state.message ?? 'Unable to check for updates.', busy: false };
    case 'unavailable':
      return { action: null, actionLabel: 'Updates unavailable', detail: state.message ?? 'Update checks are available in installed builds.', busy: false };
    default:
      return { action: state.canCheck ? 'check' : null, actionLabel: state.canCheck ? 'Check for updates' : 'Updates unavailable', detail: `TXT Sheets v${state.currentVersion}`, busy: false };
  }
}
