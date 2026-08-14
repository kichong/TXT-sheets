import { extname, resolve } from 'node:path';

const SUPPORTED = new Set(['.xlsx', '.csv', '.tsv']);

export function findLaunchWorkbookPath(args: string[]): string | null {
  for (const argument of args) {
    if (!argument || argument.startsWith('--')) continue;
    const candidate = argument.replace(/^"|"$/gu, '');
    if (SUPPORTED.has(extname(candidate).toLowerCase())) return resolve(candidate);
  }
  return null;
}
