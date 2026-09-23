import type { WorkbookDocument } from '../shared/types';

/** Move a worksheet to the boundary before insertionIndex (or after the last sheet). */
export function reorderWorksheet(workbook: WorkbookDocument, id: string, insertionIndex: number): boolean {
  const from = workbook.sheets.findIndex((sheet) => sheet.id === id);
  if (from < 0 || !Number.isInteger(insertionIndex) || insertionIndex < 0 || insertionIndex > workbook.sheets.length) return false;
  const to = insertionIndex > from ? insertionIndex - 1 : insertionIndex;
  if (to === from) return false;
  const [sheet] = workbook.sheets.splice(from, 1);
  workbook.sheets.splice(to, 0, sheet);
  return true;
}

export function renameWorksheet(workbook: WorkbookDocument, id: string, name: string): string | null {
  const sheet = workbook.sheets.find((item) => item.id === id);
  if (!sheet) return 'This sheet is no longer available.';
  if (!name || name.length > 31 || /[\\/:?*\[\]\u0000-\u001f]/u.test(name) || name.startsWith("'") || name.endsWith("'")) {
    return 'Use 1–31 characters, without \\ / : ? * [ ] or an apostrophe at either end.';
  }
  if (workbook.sheets.some((item) => item.id !== id && item.name.toLocaleLowerCase() === name.toLocaleLowerCase())) return 'A sheet with this name already exists.';
  const previous = sheet.name.toLocaleLowerCase();
  const replacement = `'${name.replace(/'/gu, "''")}'!`;
  for (const item of workbook.sheets) {
    for (const cell of Object.values(item.cells)) {
      if (!cell.formula) continue;
      // Consume string literals intact; rename only sheet qualifiers, never cell text.
      cell.formula = cell.formula.replace(/"(?:[^"]|"")*"|'((?:[^']|'')+)'!|(?<![\p{L}\p{N}_.\]])([\p{L}_][\p{L}\p{N}_.]*)!/gu, (match, quoted: string | undefined, bare: string | undefined) => {
        const qualifier = quoted?.replace(/''/gu, "'") ?? bare;
        return qualifier?.toLocaleLowerCase() === previous ? replacement : match;
      });
    }
  }
  sheet.name = name;
  return null;
}
