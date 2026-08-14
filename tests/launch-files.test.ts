import { describe, expect, it } from 'vitest';
import { findLaunchWorkbookPath } from '../src/main/launch-files';

describe('findLaunchWorkbookPath', () => {
  it('finds supported workbook files in command lines', () => {
    expect(findLaunchWorkbookPath(['--flag', 'C:\\Files\\Budget.xlsx'])).toMatch(/Budget\.xlsx$/u);
    expect(findLaunchWorkbookPath(['C:\\Files\\notes.txt'])).toBeNull();
  });
});
