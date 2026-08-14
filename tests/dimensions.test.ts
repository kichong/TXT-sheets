import { describe, expect, it } from 'vitest';
import { createBlankWorkbook } from '../src/shared/types';
import {
  MAX_COLUMN_WIDTH, MAX_ROW_HEIGHT, MIN_COLUMN_WIDTH, MIN_ROW_HEIGHT, resizeColumn, resizeRow,
} from '../src/renderer/workbook-model';

describe('sheet dimensions', () => {
  it('stores resized columns and rows', () => {
    const sheet = createBlankWorkbook().sheets[0];
    resizeColumn(sheet, 2, 144);
    resizeRow(sheet, 4, 52);
    expect(sheet.columnWidths['2']).toBe(144);
    expect(sheet.rowHeights['4']).toBe(52);
  });

  it('keeps dimensions within usable limits', () => {
    const sheet = createBlankWorkbook().sheets[0];
    resizeColumn(sheet, 0, 1);
    resizeColumn(sheet, 1, 10_000);
    resizeRow(sheet, 0, 1);
    resizeRow(sheet, 1, 10_000);
    expect(sheet.columnWidths['0']).toBe(MIN_COLUMN_WIDTH);
    expect(sheet.columnWidths['1']).toBe(MAX_COLUMN_WIDTH);
    expect(sheet.rowHeights['0']).toBe(MIN_ROW_HEIGHT);
    expect(sheet.rowHeights['1']).toBe(MAX_ROW_HEIGHT);
  });
});
