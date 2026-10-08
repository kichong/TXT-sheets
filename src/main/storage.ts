import { createHash, randomUUID } from 'node:crypto';
import { mkdir, readdir, readFile, rename, unlink, writeFile } from 'node:fs/promises';
import { basename, dirname, join } from 'node:path';
import type { RecentFile, WorkbookDocument, WorkbookFormat, WorkbookSource } from '../shared/types';

interface StoredRecent extends RecentFile { path: string; }
interface StorageState { recent: StoredRecent[]; }

export class AppStorage {
  private recentWrites: Promise<unknown> = Promise.resolve();
  private readonly statePath: string;
  private readonly recoveryPath: string;
  private state: StorageState = { recent: [] };

  constructor(userDataPath: string) {
    this.statePath = join(userDataPath, 'state.json');
    this.recoveryPath = join(userDataPath, 'recovery.json');
  }

  private recoveryFile(key?: string): string {
    if (key && !/^[a-zA-Z0-9-]+$/u.test(key)) throw new Error('Invalid recovery identifier.');
    return key && key !== 'legacy' ? join(dirname(this.recoveryPath), `recovery-${key}.json`) : this.recoveryPath;
  }

  async getRecoveryKeys(): Promise<string[]> {
    const files = await readdir(dirname(this.recoveryPath));
    return files.flatMap((file) => {
      if (file === 'recovery.json') return ['legacy'];
      const match = /^recovery-([a-zA-Z0-9-]+)\.json$/u.exec(file);
      return match ? [match[1]] : [];
    });
  }

  async initialize(): Promise<void> {
    await mkdir(dirname(this.statePath), { recursive: true });
    try {
      this.state = JSON.parse(await readFile(this.statePath, 'utf8')) as StorageState;
    } catch { this.state = { recent: [] }; }
  }

  private async persist(): Promise<void> {
    await this.atomicWrite(this.statePath, Buffer.from(JSON.stringify(this.state), 'utf8'));
  }

  async sourceFor(path: string, format: WorkbookFormat): Promise<WorkbookSource> {
    const existing = this.state.recent.find((entry) => entry.path.toLocaleLowerCase() === path.toLocaleLowerCase());
    const id = existing?.id ?? createHash('sha256').update(path.toLocaleLowerCase()).digest('hex').slice(0, 24);
    return { id, displayName: basename(path), format };
  }

  remember(path: string, format: WorkbookFormat): Promise<RecentFile[]> {
    const operation = this.recentWrites.then(() => this.rememberNow(path, format));
    this.recentWrites = operation.catch(() => undefined);
    return operation;
  }

  private async rememberNow(path: string, format: WorkbookFormat): Promise<RecentFile[]> {
    const source = await this.sourceFor(path, format);
    const entry: StoredRecent = { ...source, path, lastOpenedAt: new Date().toISOString() };
    this.state.recent = [entry, ...this.state.recent.filter((item) => item.id !== entry.id)].slice(0, 10);
    await this.persist();
    return this.getRecentFiles();
  }

  getRecentFiles(): RecentFile[] {
    return this.state.recent.map(({ path: _path, ...entry }) => entry);
  }

  getPath(id: string): string | null {
    return this.state.recent.find((entry) => entry.id === id)?.path ?? null;
  }

  async writeRecovery(workbook: WorkbookDocument, key?: string): Promise<void> {
    await this.atomicWrite(this.recoveryFile(key), Buffer.from(JSON.stringify(workbook), 'utf8'));
  }

  async getRecovery(key?: string): Promise<WorkbookDocument | null> {
    try { return JSON.parse(await readFile(this.recoveryFile(key), 'utf8')) as WorkbookDocument; }
    catch { return null; }
  }

  async clearRecovery(key?: string): Promise<void> {
    try { await unlink(this.recoveryFile(key)); } catch { /* no recovery is a valid state */ }
  }

  async atomicWrite(path: string, bytes: Uint8Array): Promise<void> {
    await mkdir(dirname(path), { recursive: true });
    const temporary = `${path}.${randomUUID()}.tmp`;
    await writeFile(temporary, bytes);
    await rename(temporary, path);
  }
}
