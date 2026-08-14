import { describe, expect, it } from 'vitest';
import { presentUpdate } from '../src/shared/updates';

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
