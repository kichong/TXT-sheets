import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { expect, it } from 'vitest';
import { AppStorage } from '../src/main/storage';
import { createBlankWorkbook } from '../src/shared/types';

it('keeps each window recovery independent and discovers drafts after restart, including legacy recovery', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'txt-window-recovery-'));
  try {
    const storage = new AppStorage(directory);
    await storage.initialize();
    const first = { ...createBlankWorkbook(), title: 'First' };
    const second = { ...createBlankWorkbook(), title: 'Second' };
    await storage.writeRecovery(first, 'first');
    await storage.writeRecovery(second, 'second');
    await storage.writeRecovery(createBlankWorkbook());
    const restarted = new AppStorage(directory);
    await restarted.initialize();
    expect((await restarted.getRecoveryKeys()).sort()).toEqual(['first', 'legacy', 'second']);
    expect(await restarted.getRecovery('first')).toEqual(first);
    await restarted.clearRecovery('first');
    expect(await restarted.getRecovery('first')).toBeNull();
    expect(await restarted.getRecovery('second')).toEqual(second);
    expect(await restarted.getRecovery('legacy')).not.toBeNull();
    expect(await restarted.getRecovery('fresh-window')).toBeNull();
  } finally { await rm(directory, { recursive: true, force: true }); }
});

it('keeps both recent files when different windows remember files simultaneously', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'txt-window-recent-'));
  try {
    const storage = new AppStorage(directory);
    await storage.initialize();
    await Promise.all([storage.remember(join(directory, 'first.csv'), 'csv'), storage.remember(join(directory, 'second.csv'), 'csv')]);
    const restarted = new AppStorage(directory);
    await restarted.initialize();
    expect((await restarted.getRecentFiles()).map((file) => file.displayName).sort()).toEqual(['first.csv', 'second.csv']);
  } finally { await rm(directory, { recursive: true, force: true }); }
});
