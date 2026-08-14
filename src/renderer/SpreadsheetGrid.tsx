import { useVirtualizer } from '@tanstack/react-virtual';
import { forwardRef, useEffect, useMemo, useRef, useState } from 'react';
import type { CSSProperties, KeyboardEvent, PointerEvent } from 'react';
import type { CellData, SheetDocument } from '../shared/types';
import type { FormulaEvaluator } from './formulas';
import { addressForCell, cellKey, columnName, editableCellText, parseCellAddress } from './formulas';
import type { Selection } from './workbook-model';
import {
  MAX_COLUMN_WIDTH, MAX_ROW_HEIGHT, MIN_COLUMN_WIDTH, MIN_ROW_HEIGHT, selectionBounds,
} from './workbook-model';

const ROW_HEADER_WIDTH = 46;
const COLUMN_HEADER_HEIGHT = 26;
const DEFAULT_COLUMN_WIDTH = 96;
const DEFAULT_ROW_HEIGHT = 25;

interface MergeInfo { masterRow: number; masterColumn: number; endRow: number; endColumn: number; }

interface SpreadsheetGridProps {
  sheet: SheetDocument;
  evaluator: FormulaEvaluator;
  selection: Selection;
  editing: boolean;
  editValue: string;
  onEditValueChange(value: string): void;
  onCommitEdit(move?: 'down' | 'right'): void;
  onCancelEdit(): void;
  onStartEdit(initial?: string): void;
  onSelectionChange(selection: Selection): void;
  onColumnResize(column: number, width: number): void;
  onRowResize(row: number, height: number): void;
  onKeyDown(event: KeyboardEvent<HTMLDivElement>): void;
}

interface ResizeSession {
  axis: 'column' | 'row';
  index: number;
  pointerId: number;
  startPosition: number;
  startSize: number;
  currentSize: number;
}

function displayNumber(value: number, format?: string): string {
  if (!format || format === 'General') return Number.isInteger(value) ? String(value) : String(Number(value.toPrecision(12)));
  const decimals = Math.min(6, (format.split('.')[1]?.match(/[0#]/gu) ?? []).length);
  if (format.includes('%')) return new Intl.NumberFormat(undefined, { style: 'percent', minimumFractionDigits: decimals, maximumFractionDigits: decimals }).format(value);
  if (format.includes('$')) return new Intl.NumberFormat(undefined, { style: 'currency', currency: 'USD', minimumFractionDigits: decimals, maximumFractionDigits: decimals }).format(value);
  return new Intl.NumberFormat(undefined, { minimumFractionDigits: decimals, maximumFractionDigits: decimals }).format(value);
}

function displayValue(cell: CellData | undefined, evaluated: ReturnType<FormulaEvaluator['evaluateCell']>): string {
  const value = cell?.formula ? evaluated : cell?.value;
  if (value === null || value === undefined) return '';
  if (cell?.valueType === 'date' && typeof value === 'string') {
    const date = new Date(value);
    return Number.isNaN(date.valueOf()) ? value : date.toLocaleDateString();
  }
  return typeof value === 'number' ? displayNumber(value, cell?.style?.numberFormat) : String(value);
}

function styleForCell(cell: CellData | undefined): CSSProperties {
  const style = cell?.style;
  const border = (side: 'top' | 'right' | 'bottom' | 'left') => {
    const value = style?.border?.[side];
    return value?.style ? `${value.style === 'thick' ? 3 : value.style === 'medium' ? 2 : 1}px ${value.style === 'dashed' || value.style === 'dotted' || value.style === 'double' ? value.style : 'solid'} ${value.color ?? 'var(--gridline-strong)'}` : undefined;
  };
  return {
    fontFamily: style?.fontFamily,
    fontSize: style?.fontSize ? `${style.fontSize}pt` : undefined,
    fontWeight: style?.bold ? 700 : undefined,
    fontStyle: style?.italic ? 'italic' : undefined,
    textDecoration: style?.underline ? 'underline' : undefined,
    color: style?.textColor,
    backgroundColor: style?.fillColor,
    textAlign: style?.horizontal,
    alignItems: style?.vertical === 'top' ? 'flex-start' : style?.vertical === 'bottom' ? 'flex-end' : 'center',
    whiteSpace: style?.wrapText ? 'normal' : 'nowrap',
    borderTop: border('top'), borderRight: border('right'), borderBottom: border('bottom'), borderLeft: border('left'),
  };
}

export const SpreadsheetGrid = forwardRef<HTMLDivElement, SpreadsheetGridProps>(function SpreadsheetGrid(props, forwardedRef) {
  const {
    sheet, evaluator, selection, editing, editValue, onEditValueChange, onCommitEdit, onCancelEdit,
    onStartEdit, onSelectionChange, onColumnResize, onRowResize, onKeyDown,
  } = props;
  const scrollRef = useRef<HTMLDivElement | null>(null);
  const [scroll, setScroll] = useState({ left: 0, top: 0 });
  const dragging = useRef(false);
  const resizing = useRef<ResizeSession | null>(null);
  const bounds = selectionBounds(selection);

  const rowVirtualizer = useVirtualizer({
    count: sheet.rowCount,
    getScrollElement: () => scrollRef.current,
    estimateSize: (index) => sheet.rowHeights[String(index)] ?? DEFAULT_ROW_HEIGHT,
    overscan: 8,
  });
  const columnVirtualizer = useVirtualizer({
    horizontal: true,
    count: sheet.columnCount,
    getScrollElement: () => scrollRef.current,
    estimateSize: (index) => sheet.columnWidths[String(index)] ?? DEFAULT_COLUMN_WIDTH,
    overscan: 5,
  });

  useEffect(() => {
    rowVirtualizer.measure();
    columnVirtualizer.measure();
  }, [columnVirtualizer, rowVirtualizer, sheet.columnWidths, sheet.id, sheet.rowHeights]);

  const mergeMap = useMemo(() => {
    const map = new Map<string, MergeInfo>();
    sheet.merges.forEach((range) => {
      const [startText, endText] = range.split(':');
      const start = parseCellAddress(startText);
      const end = parseCellAddress(endText);
      if (!start || !end) return;
      const info = { masterRow: start.row, masterColumn: start.column, endRow: end.row, endColumn: end.column };
      for (let row = start.row; row <= end.row; row += 1) {
        for (let column = start.column; column <= end.column; column += 1) map.set(cellKey(row, column), info);
      }
    });
    return map;
  }, [sheet.merges]);

  const setScrollElement = (node: HTMLDivElement | null) => {
    scrollRef.current = node;
    if (typeof forwardedRef === 'function') forwardedRef(node);
    else if (forwardedRef) forwardedRef.current = node;
  };

  const pointerDown = (event: PointerEvent, row: number, column: number) => {
    if (event.button !== 0) return;
    dragging.current = true;
    const point = { row, column };
    onSelectionChange({ anchor: event.shiftKey ? selection.anchor : point, focus: point });
    (event.currentTarget as HTMLElement).setPointerCapture(event.pointerId);
  };

  const clampSize = (axis: ResizeSession['axis'], size: number) => Math.round(Math.min(
    axis === 'column' ? MAX_COLUMN_WIDTH : MAX_ROW_HEIGHT,
    Math.max(axis === 'column' ? MIN_COLUMN_WIDTH : MIN_ROW_HEIGHT, size),
  ));

  const beginResize = (event: PointerEvent<HTMLSpanElement>, axis: ResizeSession['axis'], index: number, size: number) => {
    if (event.button !== 0) return;
    event.preventDefault();
    event.stopPropagation();
    resizing.current = {
      axis,
      index,
      pointerId: event.pointerId,
      startPosition: axis === 'column' ? event.clientX : event.clientY,
      startSize: size,
      currentSize: size,
    };
    event.currentTarget.setPointerCapture(event.pointerId);
  };

  const moveResize = (event: PointerEvent<HTMLSpanElement>) => {
    const session = resizing.current;
    if (!session || session.pointerId !== event.pointerId) return;
    event.preventDefault();
    const pointerPosition = session.axis === 'column' ? event.clientX : event.clientY;
    const nextSize = clampSize(session.axis, session.startSize + pointerPosition - session.startPosition);
    if (nextSize === session.currentSize) return;
    session.currentSize = nextSize;
    if (session.axis === 'column') columnVirtualizer.resizeItem(session.index, nextSize);
    else rowVirtualizer.resizeItem(session.index, nextSize);
  };

  const finishResize = (event: PointerEvent<HTMLSpanElement>) => {
    const session = resizing.current;
    if (!session || session.pointerId !== event.pointerId) return;
    event.preventDefault();
    event.stopPropagation();
    resizing.current = null;
    if (event.currentTarget.hasPointerCapture(event.pointerId)) event.currentTarget.releasePointerCapture(event.pointerId);
    if (session.currentSize === session.startSize) return;
    if (session.axis === 'column') onColumnResize(session.index, session.currentSize);
    else onRowResize(session.index, session.currentSize);
  };

  const resizeWithKeyboard = (
    event: KeyboardEvent<HTMLSpanElement>, axis: ResizeSession['axis'], index: number, currentSize: number,
  ) => {
    const decrement = axis === 'column' ? event.key === 'ArrowLeft' : event.key === 'ArrowUp';
    const increment = axis === 'column' ? event.key === 'ArrowRight' : event.key === 'ArrowDown';
    if (!decrement && !increment) return;
    event.preventDefault();
    event.stopPropagation();
    const nextSize = clampSize(axis, currentSize + (increment ? 1 : -1) * (event.shiftKey ? 24 : 8));
    if (axis === 'column') onColumnResize(index, nextSize);
    else onRowResize(index, nextSize);
  };

  const rows = rowVirtualizer.getVirtualItems();
  const columns = columnVirtualizer.getVirtualItems();
  const editorRow = rows.find((item) => item.index === selection.focus.row);
  const editorColumn = columns.find((item) => item.index === selection.focus.column);

  return (
    <div
      ref={setScrollElement}
      className="spreadsheet-scroll"
      tabIndex={0}
      role="grid"
      aria-label={`${sheet.name} spreadsheet grid`}
      onKeyDown={onKeyDown}
      onPointerUp={() => { dragging.current = false; }}
      onPointerCancel={() => { dragging.current = false; }}
      onScroll={(event) => setScroll({ left: event.currentTarget.scrollLeft, top: event.currentTarget.scrollTop })}
    >
      <div className="grid-canvas" style={{ width: ROW_HEADER_WIDTH + columnVirtualizer.getTotalSize(), height: COLUMN_HEADER_HEIGHT + rowVirtualizer.getTotalSize() }}>
        {columns.map((column) => (
          <div
            role="columnheader"
            aria-label={`Column ${columnName(column.index)}`}
            className={`column-header ${column.index >= bounds.left && column.index <= bounds.right ? 'is-selected' : ''}`}
            key={`column-${column.key}`}
            style={{ width: column.size, transform: `translate(${ROW_HEADER_WIDTH + column.start}px, ${scroll.top}px)` }}
          >
            <button
              type="button"
              className="header-select-button"
              onClick={(event) => onSelectionChange({ anchor: { row: event.shiftKey ? selection.anchor.row : 0, column: event.shiftKey ? selection.anchor.column : column.index }, focus: { row: sheet.rowCount - 1, column: column.index } })}
            >{columnName(column.index)}</button>
            <span
              className="column-resize-handle"
              role="separator"
              tabIndex={0}
              aria-label={`Resize column ${columnName(column.index)}`}
              aria-orientation="vertical"
              aria-valuemin={MIN_COLUMN_WIDTH}
              aria-valuemax={MAX_COLUMN_WIDTH}
              aria-valuenow={Math.round(column.size)}
              onClick={(event) => event.stopPropagation()}
              onPointerDown={(event) => beginResize(event, 'column', column.index, column.size)}
              onPointerMove={moveResize}
              onPointerUp={finishResize}
              onPointerCancel={finishResize}
              onKeyDown={(event) => resizeWithKeyboard(event, 'column', column.index, column.size)}
            />
          </div>
        ))}
        {rows.map((row) => (
          <div
            role="rowheader"
            aria-label={`Row ${row.index + 1}`}
            className={`row-header ${row.index >= bounds.top && row.index <= bounds.bottom ? 'is-selected' : ''}`}
            key={`row-${row.key}`}
            style={{ height: row.size, transform: `translate(${scroll.left}px, ${COLUMN_HEADER_HEIGHT + row.start}px)` }}
          >
            <button
              type="button"
              className="header-select-button"
              onClick={(event) => onSelectionChange({ anchor: { row: event.shiftKey ? selection.anchor.row : row.index, column: event.shiftKey ? selection.anchor.column : 0 }, focus: { row: row.index, column: sheet.columnCount - 1 } })}
            >{row.index + 1}</button>
            <span
              className="row-resize-handle"
              role="separator"
              tabIndex={0}
              aria-label={`Resize row ${row.index + 1}`}
              aria-orientation="horizontal"
              aria-valuemin={MIN_ROW_HEIGHT}
              aria-valuemax={MAX_ROW_HEIGHT}
              aria-valuenow={Math.round(row.size)}
              onClick={(event) => event.stopPropagation()}
              onPointerDown={(event) => beginResize(event, 'row', row.index, row.size)}
              onPointerMove={moveResize}
              onPointerUp={finishResize}
              onPointerCancel={finishResize}
              onKeyDown={(event) => resizeWithKeyboard(event, 'row', row.index, row.size)}
            />
          </div>
        ))}
        {rows.flatMap((row) => columns.map((column) => {
          const key = cellKey(row.index, column.index);
          const merge = mergeMap.get(key);
          if (merge && (merge.masterRow !== row.index || merge.masterColumn !== column.index)) return null;
          let width = column.size;
          let height = row.size;
          if (merge) {
            width = 0;
            for (let index = merge.masterColumn; index <= merge.endColumn; index += 1) width += sheet.columnWidths[String(index)] ?? DEFAULT_COLUMN_WIDTH;
            height = 0;
            for (let index = merge.masterRow; index <= merge.endRow; index += 1) height += sheet.rowHeights[String(index)] ?? DEFAULT_ROW_HEIGHT;
          }
          const cell = sheet.cells[key];
          const selected = row.index >= bounds.top && row.index <= bounds.bottom && column.index >= bounds.left && column.index <= bounds.right;
          const active = selection.focus.row === row.index && selection.focus.column === column.index;
          return (
            <div
              role="gridcell"
              aria-label={addressForCell(row.index, column.index)}
              aria-selected={selected}
              key={key}
              className={`grid-cell ${selected ? 'is-selected' : ''} ${active ? 'is-active' : ''} ${cell?.formula ? 'has-formula' : ''}`}
              style={{
                width, height,
                transform: `translate(${ROW_HEADER_WIDTH + column.start}px, ${COLUMN_HEADER_HEIGHT + row.start}px)`,
                ...styleForCell(cell),
              }}
              onPointerDown={(event) => pointerDown(event, row.index, column.index)}
              onPointerEnter={() => {
                if (dragging.current) onSelectionChange({ anchor: selection.anchor, focus: { row: row.index, column: column.index } });
              }}
              onDoubleClick={() => onStartEdit(editableCellText(cell))}
            >
              <span>{displayValue(cell, evaluator.evaluateCell(sheet.id, row.index, column.index))}</span>
            </div>
          );
        }))}
        {editing && editorRow && editorColumn ? (
          <input
            autoFocus
            className="cell-editor"
            style={{
              width: editorColumn.size,
              height: editorRow.size,
              transform: `translate(${ROW_HEADER_WIDTH + editorColumn.start}px, ${COLUMN_HEADER_HEIGHT + editorRow.start}px)`,
            }}
            value={editValue}
            aria-label={`Edit ${addressForCell(selection.focus.row, selection.focus.column)}`}
            onChange={(event) => onEditValueChange(event.target.value)}
            onBlur={() => onCommitEdit()}
            onKeyDown={(event) => {
              if (event.key === 'Enter') { event.preventDefault(); onCommitEdit('down'); }
              if (event.key === 'Tab') { event.preventDefault(); onCommitEdit('right'); }
              if (event.key === 'Escape') { event.preventDefault(); onCancelEdit(); }
            }}
          />
        ) : null}
        <div className="grid-corner" style={{ transform: `translate(${scroll.left}px, ${scroll.top}px)` }} aria-hidden="true" />
      </div>
    </div>
  );
});
