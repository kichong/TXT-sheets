import { mkdtemp, rm, writeFile, mkdir, readdir } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { expect, it, vi } from 'vitest';
import { AppStorage } from '../src/main/storage';
import { createBlankWorkbook } from '../src/shared/types';

it('clears recovery after a pending write and preserves the order of later writes', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'txt-recovery-order-'));
  try {
    const storage = new AppStorage(directory);
    await storage.initialize();
    const write = storage.atomicWrite.bind(storage);
    let release!: () => void;
    const blocked = new Promise<void>((resolve) => { release = resolve; });
    vi.spyOn(storage, 'atomicWrite').mockImplementationOnce(async (path, bytes) => {
      await blocked;
      await write(path, bytes);
    });
    const pending = storage.writeRecovery({ ...createBlankWorkbook(), title: 'Old' }, 'first');
    const cleared = storage.clearRecovery('first');
    release();
    await Promise.all([pending, cleared]);
    expect(await storage.getRecovery('first')).toBeNull();
    await Promise.all([
      storage.writeRecovery({ ...createBlankWorkbook(), title: 'Earlier' }, 'first'),
      storage.writeRecovery({ ...createBlankWorkbook(), title: 'Latest' }, 'first'),
    ]);
    expect((await storage.getRecovery('first'))?.title).toBe('Latest');
    vi.spyOn(storage, 'atomicWrite').mockRejectedValueOnce(new Error('Write failed'));
    await expect(storage.writeRecovery(createBlankWorkbook(), 'first')).rejects.toThrow('Write failed');
    await storage.clearRecovery('first');
    expect(await storage.getRecovery('first')).toBeNull();
  } finally { await rm(directory, { recursive: true, force: true }); }
});

it('retains an open workbook save path when it leaves the recent list', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'txt-source-path-'));
  try {
    const storage = new AppStorage(directory);
    await storage.initialize();
    const path = join(directory, 'first.csv');
    const source = await storage.sourceFor(path, 'csv');
    await storage.remember(path, 'csv');
    for (let index = 0; index < 11; index++) await storage.remember(join(directory, `${index}.csv`), 'csv');
    expect(storage.getRecentFiles().some((entry) => entry.id === source.id)).toBe(false);
    expect(storage.getPath(source.id)).toBe(path);
  } finally { await rm(directory, { recursive: true, force: true }); }
});

it('ignores malformed recent entries and removes temporary files after a failed replacement', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'txt-storage-failure-'));
  try {
    await writeFile(join(directory, 'state.json'), JSON.stringify({ recent: [null, {}, { path: 123 }] }));
    const storage = new AppStorage(directory);
    await storage.initialize();
    expect(storage.getRecentFiles()).toEqual([]);
    await storage.remember(join(directory, 'valid.csv'), 'csv');
    const target = join(directory, 'folder.csv');
    await mkdir(target);
    await expect(storage.atomicWrite(target, Buffer.from('data'))).rejects.toThrow();
    expect((await readdir(directory)).some((name) => name.endsWith('.tmp'))).toBe(false);
  } finally { await rm(directory, { recursive: true, force: true }); }
});

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
