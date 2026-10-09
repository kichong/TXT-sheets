import {
  AlignCenter, AlignLeft, AlignRight, Bold, Bug, ChevronDown, ChevronLeft, ChevronRight,
  CircleAlert, Columns3, Download, FilePlus2, FolderOpen, FunctionSquare, Italic, Moon, Plus,
  Redo2, RefreshCw, Rows3, Save, Search, Settings2, Sigma, Sun, Trash2, Underline, Undo2, X,
} from 'lucide-react';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { KeyboardEvent } from 'react';
import { createBlankWorkbook } from '../shared/types';
import type { AppCommand, AppUpdateState, CellStyle, RecentFile, WorkbookDocument } from '../shared/types';
import { presentUpdate, type UpdateAction } from '../shared/updates';
import { GridContextMenu, type GridMenuTarget, type GridMenuItem } from './GridContextMenu';
import { SheetTabMenu, type SheetMenuTarget } from './SheetTabMenu';
import { renameWorksheet, reorderWorksheet } from './sheet-actions';
import { SpreadsheetGrid } from './SpreadsheetGrid';
import {
  addressForCell, cellKey, createFormulaEvaluator, editableCellText, isDateNumberFormat, normalizeCellInput,
} from './formulas';
import {
  applyStyle, cloneWorkbook, deleteColumn, deleteRow, fillSelection, insertColumn, insertRow, setCellInput, selectedCells,
  resizeColumn, resizeRow, selectionBounds, selectionLabel, uniqueSheetName, cellShiftProblem, shiftSelectedCells,
} from './workbook-model';
import type { Selection } from './workbook-model';

interface HistoryState {
  past: WorkbookDocument[];
  present: WorkbookDocument;
  future: WorkbookDocument[];
}

const INITIAL_SELECTION: Selection = { anchor: { row: 0, column: 0 }, focus: { row: 0, column: 0 } };

function numberFormatChoice(format: string | undefined): string {
  if (!format || format === 'General') return 'General';
  if (isDateNumberFormat(format)) return 'm/d/yy';
  if (format.includes('$')) return '$#,##0.00;($#,##0.00)';
  if (format.includes('%')) return '0.0%';
  return '#,##0.00';
}

function workbookWithFormulaResults(workbook: WorkbookDocument): WorkbookDocument {
  const result = cloneWorkbook(workbook);
  const evaluator = createFormulaEvaluator(result);
  result.sheets.forEach((sheet) => Object.entries(sheet.cells).forEach(([key, cell]) => {
    if (!cell.formula) return;
    const [row, column] = key.split(':').map(Number);
    cell.cachedValue = evaluator.evaluateCell(sheet.id, row, column);
  }));
  return result;
}

function IconButton(props: React.ButtonHTMLAttributes<HTMLButtonElement> & { label: string; active?: boolean }) {
  const { label, active, className = '', ...buttonProps } = props;
  return <button type="button" className={`icon-button ${active ? 'is-active' : ''} ${className}`} title={label} aria-label={label} {...buttonProps} />;
}

export function App() {
  const [history, setHistory] = useState<HistoryState>(() => ({ past: [], present: createBlankWorkbook(), future: [] }));
  const [selection, setSelection] = useState<Selection>(INITIAL_SELECTION);
  const [dirty, setDirty] = useState(false);
  const [editing, setEditing] = useState(false);
  const [editValue, setEditValue] = useState('');
  const [theme, setTheme] = useState<'light' | 'dark'>(() => localStorage.getItem('txt-sheets-theme') === 'dark' ? 'dark' : 'light');
  const [recentFiles, setRecentFiles] = useState<RecentFile[]>([]);
  const [fileMenuOpen, setFileMenuOpen] = useState(false);
  const [gridMenu, setGridMenu] = useState<GridMenuTarget | null>(null);
  const [sheetMenu, setSheetMenu] = useState<SheetMenuTarget | null>(null);
  const [sheetDrop, setSheetDrop] = useState<{ id: string; side: 'before' | 'after' } | null>(null);
  const draggedSheetId = useRef<string | null>(null);
  const [structureMenuOpen, setStructureMenuOpen] = useState(false);
  const [compatibilityOpen, setCompatibilityOpen] = useState(false);
  const [findOpen, setFindOpen] = useState(false);
  const [findQuery, setFindQuery] = useState('');
  const [findIndex, setFindIndex] = useState(0);
  const [message, setMessage] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);
  const [updateState, setUpdateState] = useState<AppUpdateState>({ currentVersion: '…', phase: 'unavailable', canCheck: false });
  const gridRef = useRef<HTMLDivElement>(null);
  const formulaRef = useRef<HTMLInputElement>(null);
  const findRef = useRef<HTMLInputElement>(null);
  const workbook = history.present;
  const savingRef = useRef(false);
  const latestEdit = useRef({ workbook, editing, editValue, selection });
  latestEdit.current = { workbook, editing, editValue, selection };
  const activeSheet = workbook.sheets.find((sheet) => sheet.id === workbook.activeSheetId) ?? workbook.sheets[0];
  const evaluator = useMemo(() => createFormulaEvaluator(workbook), [workbook]);
  const activeCell = activeSheet.cells[cellKey(selection.focus.row, selection.focus.column)];

  useEffect(() => {
    const clamp = (point: Selection['focus']) => ({
      row: Math.max(0, Math.min(point.row, activeSheet.rowCount - 1)),
      column: Math.max(0, Math.min(point.column, activeSheet.columnCount - 1)),
    });
    setSelection((current) => {
      const anchor = clamp(current.anchor);
      const focus = clamp(current.focus);
      return anchor.row === current.anchor.row && anchor.column === current.anchor.column &&
        focus.row === current.focus.row && focus.column === current.focus.column ? current : { ...current, anchor, focus };
    });
  }, [activeSheet.rowCount, activeSheet.columnCount]);

  const commit = useCallback((mutator: (draft: WorkbookDocument) => void) => {
    setHistory((current) => {
      const next = cloneWorkbook(current.present);
      mutator(next);
      return { past: [...current.past.slice(-49), current.present], present: next, future: [] };
    });
    setDirty(true);
  }, []);

  const replaceWorkbook = useCallback((next: WorkbookDocument) => {
    setHistory({ past: [], present: next, future: [] });
    setSelection(INITIAL_SELECTION);
    setEditing(false);
    setDirty(false);
    setCompatibilityOpen(false);
  }, []);

  const undo = useCallback(() => {
    setHistory((current) => {
      const previous = current.past.at(-1);
      if (!previous) return current;
      setDirty(true);
      return { past: current.past.slice(0, -1), present: previous, future: [current.present, ...current.future] };
    });
  }, []);

  const redo = useCallback(() => {
    setHistory((current) => {
      const next = current.future[0];
      if (!next) return current;
      setDirty(true);
      return { past: [...current.past, current.present], present: next, future: current.future.slice(1) };
    });
  }, []);

  const handleError = useCallback((error: unknown) => {
    setMessage(error instanceof Error ? error.message : 'Something went wrong.');
  }, []);

  const openResult = useCallback((result: Awaited<ReturnType<typeof window.spreadsheet.open>>) => {
    if (!result) return;
    replaceWorkbook(result.workbook);
    setRecentFiles(result.recentFiles);
  }, [replaceWorkbook]);

  const openWorkbook = useCallback(async () => {
    try { openResult(await window.spreadsheet.open()); }
    catch (error) { handleError(error); }
  }, [handleError, openResult]);

  const saveWorkbook = useCallback(async (saveAs = false, closeWhenDone = false) => {
    if (savingRef.current) return;
    savingRef.current = true;
    const originalEdit = latestEdit.current;
    const focusTarget = document.activeElement instanceof HTMLInputElement ? document.activeElement : gridRef.current;
    setSaving(true);
    try {
      const current = cloneWorkbook(workbook);
      if (editing) {
        const sheet = current.sheets.find((item) => item.id === current.activeSheetId)!;
        const { row, column } = selection.focus;
        setCellInput(sheet, row, column, editValue);
      }
      const prepared = workbookWithFormulaResults(current);
      const result = saveAs ? await window.spreadsheet.saveAs(prepared) : await window.spreadsheet.save(prepared);
      if (result.status === 'saved' && result.source) {
        const next = { ...prepared, source: result.source, title: result.source.displayName.replace(/\.(xlsx|csv|tsv)$/iu, '') };
        const latest = latestEdit.current;
        const changedDuringSave = latest.workbook !== originalEdit.workbook ||
          latest.editing !== originalEdit.editing || latest.editValue !== originalEdit.editValue ||
          (latest.editing && (latest.selection.focus.row !== originalEdit.selection.focus.row || latest.selection.focus.column !== originalEdit.selection.focus.column));
        const withSource = (document: WorkbookDocument) => ({ ...document, source: result.source, title: next.title });
        setHistory((current) => ({
          past: current.past.map(withSource),
          present: changedDuringSave ? withSource(current.present) : next,
          future: current.future.map(withSource),
        }));
        setRecentFiles(result.recentFiles);
        setDirty(changedDuringSave);
        if (!changedDuringSave) setEditing(false);
        window.spreadsheet.setDirty(changedDuringSave);
        if (changedDuringSave) {
          const recovery = cloneWorkbook(withSource(latest.workbook));
          if (latest.editing) {
            const sheet = recovery.sheets.find((item) => item.id === recovery.activeSheetId)!;
            setCellInput(sheet, latest.selection.focus.row, latest.selection.focus.column, latest.editValue);
          }
          await window.spreadsheet.writeRecovery(workbookWithFormulaResults(recovery));
          setMessage('Saved the earlier changes. Your latest edits still need saving.');
          return;
        }
        setMessage(saveAs ? `Saved as ${result.source.displayName}` : `Saved ${result.source.displayName}`);
        if (closeWhenDone) window.spreadsheet.requestCloseAfterSave();
      }
    } catch (error) { handleError(error); }
    finally {
      savingRef.current = false;
      setSaving(false);
      requestAnimationFrame(() => {
        if (focusTarget?.isConnected) focusTarget.focus();
        else gridRef.current?.focus();
      });
    }
  }, [editValue, editing, handleError, saving, selection.focus, workbook]);

  const reportCompatibility = useCallback(async () => {
    try {
      await window.spreadsheet.reportCompatibility({
        sourceFormat: workbook.source?.format ?? 'unsaved',
        issues: workbook.compatibilityIssues,
      });
      setMessage('Compatibility report opened in GitHub. Review it, then submit.');
    } catch (error) { handleError(error); }
  }, [handleError, workbook.compatibilityIssues, workbook.source?.format]);

  const newWorkbook = useCallback(() => {
    void window.spreadsheet.newWindow().catch(handleError);
  }, [handleError]);

  const runCommand = useCallback((command: AppCommand) => {
    if (command === 'new') newWorkbook();
    if (command === 'open') void openWorkbook();
    if (command === 'save') void saveWorkbook(false);
    if (command === 'save-as') void saveWorkbook(true);
    if (command === 'save-and-close') void saveWorkbook(false, true);
    if (command === 'undo') undo();
    if (command === 'redo') redo();
    if (command === 'find') { setFindOpen(true); requestAnimationFrame(() => findRef.current?.focus()); }
  }, [newWorkbook, openWorkbook, redo, saveWorkbook, undo]);

  useEffect(() => {
    void window.spreadsheet.getRecentFiles().then(setRecentFiles).catch(handleError);
    return window.spreadsheet.onCommand(runCommand);
  }, [handleError, runCommand]);

  useEffect(() => {
    let disposed = false;
    let externalOpened = false;
    let openQueue = Promise.resolve();
    const openQueuedFile = () => {
      openQueue = openQueue.then(async () => {
        const result = await window.spreadsheet.openExternal();
        if (disposed || !result) return;
        externalOpened = true;
        openResult(result);
      }).catch(handleError);
      return openQueue;
    };
    const unsubscribe = window.spreadsheet.onExternalFile(() => { void openQueuedFile(); });
    void (async () => {
      await openQueuedFile();
      const recovery = await window.spreadsheet.getRecovery();
      if (!disposed && !externalOpened && recovery) {
        replaceWorkbook(recovery);
        setDirty(true);
        setMessage('Recovered unsaved work');
      }
    })().catch(handleError);
    return () => { disposed = true; unsubscribe(); };
  }, [handleError, openResult, replaceWorkbook]);

  useEffect(() => {
    window.spreadsheet.setDirty(dirty || editing);
  }, [dirty, editing]);

  useEffect(() => {
    if ((!dirty && !editing) || saving) return;
    const timer = window.setTimeout(() => {
      const recovery = cloneWorkbook(workbook);
      if (editing) {
        const sheet = recovery.sheets.find((item) => item.id === recovery.activeSheetId)!;
        setCellInput(sheet, selection.focus.row, selection.focus.column, editValue);
      }
      void window.spreadsheet.writeRecovery(workbookWithFormulaResults(recovery)).catch(handleError);
    }, 800);
    return () => window.clearTimeout(timer);
  }, [dirty, editing, editValue, saving, selection.focus.row, selection.focus.column, handleError, workbook]);

  useEffect(() => {
    document.documentElement.dataset.theme = theme;
    localStorage.setItem('txt-sheets-theme', theme);
  }, [theme]);

  useEffect(() => {
    const unsubscribe = window.spreadsheet.onUpdateState(setUpdateState);
    void window.spreadsheet.getUpdateState().then(setUpdateState).catch(handleError);
    return unsubscribe;
  }, [handleError]);

  useEffect(() => {
    if (!message) return;
    const timer = window.setTimeout(() => setMessage(null), 2600);
    return () => window.clearTimeout(timer);
  }, [message]);

  const commitEdit = useCallback((move?: 'down' | 'right') => {
    if (!editing || !latestEdit.current.editing) return;
    latestEdit.current = { ...latestEdit.current, editing: false };
    const point = selection.focus;
    if (editValue !== editableCellText(activeCell)) {
      commit((draft) => {
        const sheet = draft.sheets.find((item) => item.id === draft.activeSheetId)!;
        setCellInput(sheet, point.row, point.column, editValue);
      });
    }
    setEditing(false);
    if (move) {
      const focus = {
        row: Math.min(activeSheet.rowCount - 1, point.row + (move === 'down' ? 1 : 0)),
        column: Math.min(activeSheet.columnCount - 1, point.column + (move === 'right' ? 1 : 0)),
      };
      setSelection({ anchor: focus, focus });
      requestAnimationFrame(() => gridRef.current?.focus());
    }
  }, [activeCell, activeSheet.columnCount, activeSheet.rowCount, commit, editValue, editing, selection.focus]);

  const cancelEdit = useCallback(() => {
    latestEdit.current = { ...latestEdit.current, editing: false };
    setEditing(false);
    gridRef.current?.focus();
  }, []);

  const startEdit = useCallback((initial?: string) => {
    setEditValue(initial ?? editableCellText(activeCell));
    setEditing(true);
  }, [activeCell]);

  const clearSelection = useCallback(() => {
    commit((draft) => {
      const sheet = draft.sheets.find((item) => item.id === draft.activeSheetId)!;
      const bounds = selectionBounds(selection);
      Object.entries(sheet.cells).forEach(([key, cell]) => {
        const [row, column] = key.split(':').map(Number);
        if (row < bounds.top || row > bounds.bottom || column < bounds.left || column > bounds.right) return;
        if (cell?.style) sheet.cells[key] = { value: null, valueType: 'blank', style: structuredClone(cell.style) };
        else delete sheet.cells[key];
      });
    });
  }, [commit, selection]);

  const copySelection = useCallback(async (cut = false) => {
    const bounds = selectionBounds(selection);
    const rows: string[] = [];
    for (let row = bounds.top; row <= bounds.bottom; row += 1) {
      const values = [];
      for (let column = bounds.left; column <= bounds.right; column += 1) {
        values.push(editableCellText(activeSheet.cells[cellKey(row, column)]));
      }
      rows.push(values.join('\t'));
    }
    try {
      await navigator.clipboard.writeText(rows.join('\r\n'));
      if (cut) clearSelection();
    } catch (error) { handleError(error); }
  }, [activeSheet.cells, clearSelection, handleError, selection]);

  const pasteSelection = useCallback(async () => {
    try {
      const text = await navigator.clipboard.readText();
      const rows = text.replace(/\r/gu, '').split('\n').map((row) => row.split('\t'));
      const bounds = selectionBounds(selection);
      const origin = { row: bounds.top, column: bounds.left };
      commit((draft) => {
        const sheet = draft.sheets.find((item) => item.id === draft.activeSheetId)!;
        rows.forEach((values, rowOffset) => values.forEach((value, columnOffset) => {
          const row = origin.row + rowOffset;
          const column = origin.column + columnOffset;
          if (row >= sheet.rowCount || column >= sheet.columnCount) return;
          setCellInput(sheet, row, column, value);
        }));
      });
      const focus = { row: Math.min(activeSheet.rowCount - 1, origin.row + rows.length - 1), column: Math.min(activeSheet.columnCount - 1, origin.column + rows.reduce((maximum, row) => Math.max(maximum, row.length - 1), 0)) };
      setSelection({ anchor: origin, focus });
    } catch (error) { handleError(error); }
  }, [activeSheet.columnCount, activeSheet.rowCount, commit, handleError, selection]);

  const closeGridMenu = () => { setGridMenu(null); gridRef.current?.focus({ preventScroll: true }); };
  const changeSelectedStructure = (axis: 'row' | 'column', delta: 1 | -1, after = false) => {
    const bounds = selectionBounds(selection);
    const first = axis === 'row' ? bounds.top : bounds.left;
    const last = axis === 'row' ? bounds.bottom : bounds.right;
    const size = axis === 'row' ? activeSheet.rowCount : activeSheet.columnCount;
    const count = delta === -1 ? Math.min(last - first + 1, size - 1) : last - first + 1;
    if (count === 0) return;
    if (delta === 1 && size + count > (axis === 'row' ? 1_048_576 : 16_384)) { setMessage('There is no room to insert this selection at the worksheet limit.'); return; }
    commit((draft) => {
      const operation = axis === 'row' ? (delta === 1 ? insertRow : deleteRow) : (delta === 1 ? insertColumn : deleteColumn);
      for (let n = 0; n < count; n++) operation(draft, draft.activeSheetId, after ? last + 1 : first);
    });
    const focus = { row: bounds.top, column: bounds.left };
    setSelection({ anchor: focus, focus });
  };
  const shiftCells = (axis: 'row' | 'column', delta: 1 | -1) => {
    const problem = cellShiftProblem(workbook, activeSheet.id, selection, axis, delta);
    if (problem) { setMessage(problem); return; }
    commit((draft) => shiftSelectedCells(draft, draft.activeSheetId, selection, axis, delta));
  };
  const gridMenuItems: Array<GridMenuItem | null> = gridMenu ? [
    { label: 'Cut', shortcut: 'Ctrl+X', action: () => { void copySelection(true); } },
    { label: 'Copy', shortcut: 'Ctrl+C', action: () => { void copySelection(); } },
    { label: 'Paste', shortcut: 'Ctrl+V', action: () => { void pasteSelection(); } },
    null,
    ...(selection.kind !== 'columns' ? [
      { label: 'Insert rows above', action: () => changeSelectedStructure('row', 1) },
      { label: 'Insert rows below', action: () => changeSelectedStructure('row', 1, true) },
      { label: 'Delete rows', disabled: activeSheet.rowCount <= 1, action: () => changeSelectedStructure('row', -1) },
    ] : []),
    ...(selection.kind !== 'rows' ? [
      { label: 'Insert columns left', action: () => changeSelectedStructure('column', 1) },
      { label: 'Insert columns right', action: () => changeSelectedStructure('column', 1, true) },
      { label: 'Delete columns', disabled: activeSheet.columnCount <= 1, action: () => changeSelectedStructure('column', -1) },
    ] : []),
    ...(selection.kind !== 'rows' && selection.kind !== 'columns' ? [
      null,
      { label: 'Insert cells, shift down', action: () => shiftCells('row', 1) },
      { label: 'Insert cells, shift right', action: () => shiftCells('column', 1) },
      { label: 'Delete cells, shift up', action: () => shiftCells('row', -1) },
      { label: 'Delete cells, shift left', action: () => shiftCells('column', -1) },
    ] : []),
    null,
    { label: 'Clear contents', shortcut: 'Delete', action: clearSelection },
    { label: 'Undo', shortcut: 'Ctrl+Z', disabled: history.past.length === 0, action: undo },
    { label: 'Redo', shortcut: 'Ctrl+Y', disabled: history.future.length === 0, action: redo },
  ] : [];

  const gridKeyDown = useCallback((event: KeyboardEvent<HTMLDivElement>) => {
    if (editing) return;
    const shortcut = event.ctrlKey || event.metaKey;
    const key = event.key.toLowerCase();
    if (shortcut && key === 'c') { event.preventDefault(); void copySelection(); return; }
    if (shortcut && key === 'x') { event.preventDefault(); void copySelection(true); return; }
    if (shortcut && key === 'v') { event.preventDefault(); void pasteSelection(); return; }
    if (shortcut && key === 'z') { event.preventDefault(); event.shiftKey ? redo() : undo(); return; }
    if (shortcut && key === 'y') { event.preventDefault(); redo(); return; }
    if (shortcut && key === 's') { event.preventDefault(); void saveWorkbook(event.shiftKey); return; }
    if (shortcut && key === 'o') { event.preventDefault(); void openWorkbook(); return; }
    if (shortcut && key === 'n') { event.preventDefault(); newWorkbook(); return; }
    if (shortcut && key === 'f') { event.preventDefault(); setFindOpen(true); requestAnimationFrame(() => findRef.current?.focus()); return; }
    if (shortcut && key === 'a') { event.preventDefault(); setSelection({ anchor: { row: 0, column: 0 }, focus: { row: activeSheet.rowCount - 1, column: activeSheet.columnCount - 1 } }); return; }
    if (event.key === 'Delete' || event.key === 'Backspace') { event.preventDefault(); clearSelection(); return; }
    if (event.key === 'F2') { event.preventDefault(); startEdit(); return; }
    if (event.key === 'Enter' || event.key === 'Tab') {
      event.preventDefault();
      const focus = {
        row: Math.max(0, Math.min(activeSheet.rowCount - 1, selection.focus.row + (event.key === 'Enter' ? (event.shiftKey ? -1 : 1) : 0))),
        column: Math.max(0, Math.min(activeSheet.columnCount - 1, selection.focus.column + (event.key === 'Tab' ? (event.shiftKey ? -1 : 1) : 0))),
      };
      setSelection({ anchor: focus, focus });
      return;
    }
    const movement: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (movement[event.key]) {
      event.preventDefault();
      const [rowDelta, columnDelta] = movement[event.key];
      const focus = {
        row: Math.max(0, Math.min(activeSheet.rowCount - 1, selection.focus.row + rowDelta)),
        column: Math.max(0, Math.min(activeSheet.columnCount - 1, selection.focus.column + columnDelta)),
      };
      setSelection({ anchor: event.shiftKey ? selection.anchor : focus, focus });
      return;
    }
    if (!shortcut && !event.altKey && event.key.length === 1) { event.preventDefault(); startEdit(event.key); }
  }, [activeSheet.columnCount, activeSheet.rowCount, clearSelection, copySelection, editing, newWorkbook, openWorkbook, pasteSelection, redo, saveWorkbook, selection.anchor, selection.focus, startEdit, undo]);

  const formatSelection = useCallback((style: Partial<CellStyle>) => {
    commit((draft) => applyStyle(draft.sheets.find((sheet) => sheet.id === draft.activeSheetId)!, selection, style));
  }, [commit, selection]);

  const applyNumberFormat = useCallback((numberFormat: string) => {
    commit((draft) => {
      const sheet = draft.sheets.find((item) => item.id === draft.activeSheetId)!;
      selectedCells(sheet, selection).forEach(({ row, column, cell }) => {
        const key = cellKey(row, column);
        const style = { ...cell?.style, numberFormat };
        if (!cell || cell.formula || cell.value === null) {
          sheet.cells[key] = { value: cell?.value ?? null, valueType: cell?.valueType ?? 'blank', ...cell, style };
          return;
        }
        const normalized = normalizeCellInput(String(cell.value), { ...cell, style });
        sheet.cells[key] = { ...(normalized ?? cell), style };
      });
    });
  }, [commit, selection]);

  const fillSelectedCells = useCallback((source: Selection, target: Selection) => {
    commit((draft) => fillSelection(draft.sheets.find((sheet) => sheet.id === draft.activeSheetId)!, source, target));
  }, [commit]);

  const pickFormulaReference = useCallback((row: number, column: number) => {
    setEditValue((current) => `${current}${addressForCell(row, column)}`);
  }, []);

  const selectionStats = useMemo(() => {
    const bounds = selectionBounds(selection);
    const values = Object.keys(activeSheet.cells).flatMap((key) => {
      const [row, column] = key.split(':').map(Number);
      if (row < bounds.top || row > bounds.bottom || column < bounds.left || column > bounds.right) return [];
      const value = evaluator.evaluateCell(activeSheet.id, row, column);
      return typeof value === 'number' ? [value] : [];
    });
    const sum = values.reduce((total, value) => total + value, 0);
    return { count: values.length, sum, average: values.length ? sum / values.length : 0 };
  }, [activeSheet, evaluator, selection]);

  const findMatches = useMemo(() => {
    if (!findQuery.trim()) return [];
    const query = findQuery.toLocaleLowerCase();
    return Object.entries(activeSheet.cells).flatMap(([key, cell]) => {
      const [row, column] = key.split(':').map(Number);
      const value = cell.formula ? evaluator.evaluateCell(activeSheet.id, row, column) : cell.value;
      return String(value ?? '').toLocaleLowerCase().includes(query) ? [{ row, column }] : [];
    });
  }, [activeSheet, evaluator, findQuery]);

  const selectFindResult = useCallback((index: number) => {
    if (!findMatches.length) return;
    const normalized = (index + findMatches.length) % findMatches.length;
    setFindIndex(normalized);
    const point = findMatches[normalized];
    setSelection({ anchor: point, focus: point });
  }, [findMatches]);

  const addSum = useCallback(() => {
    const bounds = selectionBounds(selection);
    const target = bounds.top === bounds.bottom && bounds.left === bounds.right
      ? selection.focus
      : { row: Math.min(activeSheet.rowCount - 1, bounds.bottom + 1), column: bounds.left };
    const formula = bounds.top === bounds.bottom && bounds.left === bounds.right
      ? '=SUM('
      : `=SUM(${addressForCell(bounds.top, bounds.left)}:${addressForCell(bounds.bottom, bounds.right)})`;
    setSelection({ anchor: target, focus: target });
    if (formula.endsWith('(')) startEdit(formula);
    else {
      commit((draft) => { draft.sheets.find((sheet) => sheet.id === draft.activeSheetId)!.cells[cellKey(target.row, target.column)] = { value: null, formula, valueType: 'number' }; });
    }
  }, [activeSheet.rowCount, commit, selection, startEdit]);

  const addSheet = useCallback(() => {
    commit((draft) => {
      const id = crypto.randomUUID();
      draft.sheets.push({ id, name: uniqueSheetName(draft), rowCount: 200, columnCount: 26, cells: {}, merges: [], columnWidths: {}, rowHeights: {} });
      draft.activeSheetId = id;
    });
    setSelection(INITIAL_SELECTION);
  }, [commit]);

  const deleteSheet = useCallback((sheetId: string) => {
    if (workbook.sheets.length <= 1) { setMessage('A workbook needs at least one sheet.'); return; }
    if (!window.confirm('Delete this sheet? This can be undone.')) return;
    commit((draft) => {
      const index = draft.sheets.findIndex((sheet) => sheet.id === sheetId);
      draft.sheets.splice(index, 1);
      if (draft.activeSheetId === sheetId) draft.activeSheetId = draft.sheets[Math.max(0, index - 1)].id;
    });
    setSelection(INITIAL_SELECTION);
  }, [commit, workbook.sheets.length]);

  const dropSheet = (targetId: string, side: 'before' | 'after') => {
    const sourceId = draggedSheetId.current;
    const targetIndex = workbook.sheets.findIndex((sheet) => sheet.id === targetId);
    if (sourceId && targetIndex >= 0) {
      const insertionIndex = targetIndex + (side === 'after' ? 1 : 0);
      const sourceIndex = workbook.sheets.findIndex((sheet) => sheet.id === sourceId);
      const destinationIndex = insertionIndex > sourceIndex ? insertionIndex - 1 : insertionIndex;
      if (sourceIndex >= 0 && sourceIndex !== destinationIndex) {
        commit((draft) => { reorderWorksheet(draft, sourceId, insertionIndex); });
      }
    }
    draggedSheetId.current = null;
    setSheetDrop(null);
  };

  const selectedStyle = activeCell?.style ?? {};
  const numberFormat = numberFormatChoice(selectedStyle.numberFormat);
  const updatePresentation = presentUpdate(updateState);

  const runUpdateAction = async (action: UpdateAction) => {
    if (!action) return;
    try {
      if (action === 'check') setUpdateState(await window.spreadsheet.checkForUpdates());
      if (action === 'download') setUpdateState(await window.spreadsheet.downloadUpdate());
      if (action === 'install') {
        if (dirty) { setMessage('Save your workbook before restarting to install the update.'); return; }
        await window.spreadsheet.installUpdate();
      }
    } catch (error) { handleError(error); }
  };

  return (
    <main className="app-shell">
      <header className="app-header">
        <div className="titlebar">
          <img className="brand-mark" src="./txt-sheets-logo.svg" alt="" aria-hidden="true" />
          <div className="workbook-identity">
            <input
              value={workbook.title}
              aria-label="Workbook title"
              onChange={(event) => commit((draft) => { draft.title = event.target.value; })}
            />
            <span>{dirty ? 'Unsaved changes' : workbook.source ? 'Saved locally' : 'New workbook'}</span>
          </div>
          <div className="title-actions">
            {workbook.compatibilityIssues.length ? (
              <button className="compatibility-button" onClick={() => setCompatibilityOpen((value) => !value)}>
                <CircleAlert size={15} /> Compatibility <span>{workbook.compatibilityIssues.length}</span>
              </button>
            ) : null}
            <IconButton label={theme === 'light' ? 'Use dark theme' : 'Use light theme'} onClick={() => setTheme((value) => value === 'light' ? 'dark' : 'light')}>
              {theme === 'light' ? <Moon size={16} /> : <Sun size={16} />}
            </IconButton>
          </div>
        </div>

        <div className="toolbar" role="toolbar" aria-label="Spreadsheet tools">
          <div className="file-actions tool-group">
            <div className="menu-anchor">
              <button className="file-menu-button" onClick={() => setFileMenuOpen((value) => !value)}><FilePlus2 size={16} /> File <ChevronDown size={13} /></button>
              {fileMenuOpen ? (
                <div className="popover file-menu">
                  <button onClick={() => { newWorkbook(); setFileMenuOpen(false); }}><FilePlus2 size={15} /> New workbook <kbd>Ctrl N</kbd></button>
                  <button onClick={() => { void openWorkbook(); setFileMenuOpen(false); }}><FolderOpen size={15} /> Open… <kbd>Ctrl O</kbd></button>
                  <button onClick={() => { void saveWorkbook(false); setFileMenuOpen(false); }}><Save size={15} /> Save <kbd>Ctrl S</kbd></button>
                  <button onClick={() => { void saveWorkbook(true); setFileMenuOpen(false); }}><Save size={15} /> Save as…</button>
                  {recentFiles.length ? <div className="menu-label">Recent</div> : null}
                  {recentFiles.slice(0, 5).map((file) => (
                    <button key={file.id} onClick={() => { void window.spreadsheet.openRecent(file.id).then(openResult).catch(handleError); setFileMenuOpen(false); }}>
                      <span className="file-type">{file.format}</span><span className="recent-name">{file.displayName}</span>
                    </button>
                  ))}
                </div>
              ) : null}
            </div>
            <IconButton label="Save" disabled={saving} onClick={() => void saveWorkbook(false)}><Save size={16} /></IconButton>
          </div>
          <div className="tool-group">
            <IconButton label="Undo" disabled={!history.past.length} onClick={undo}><Undo2 size={16} /></IconButton>
            <IconButton label="Redo" disabled={!history.future.length} onClick={redo}><Redo2 size={16} /></IconButton>
          </div>
          <div className="tool-group format-group">
            <IconButton label="Bold" active={selectedStyle.bold} onClick={() => formatSelection({ bold: !selectedStyle.bold })}><Bold size={16} /></IconButton>
            <IconButton label="Italic" active={selectedStyle.italic} onClick={() => formatSelection({ italic: !selectedStyle.italic })}><Italic size={16} /></IconButton>
            <IconButton label="Underline" active={selectedStyle.underline} onClick={() => formatSelection({ underline: !selectedStyle.underline })}><Underline size={16} /></IconButton>
            <label className="color-control" title="Text color"><span>A</span><input aria-label="Text color" type="color" value={selectedStyle.textColor ?? '#202124'} onChange={(event) => formatSelection({ textColor: event.target.value })} /></label>
            <label className="color-control fill-control" title="Fill color"><span /><input aria-label="Fill color" type="color" value={selectedStyle.fillColor ?? '#fff4be'} onChange={(event) => formatSelection({ fillColor: event.target.value })} /></label>
          </div>
          <div className="tool-group">
            <select aria-label="Number format" value={numberFormat} onChange={(event) => applyNumberFormat(event.target.value)}>
              <option value="General">General</option><option value="#,##0.00">Number</option><option value="$#,##0.00;($#,##0.00)">Currency</option><option value="0.0%">Percent</option><option value="m/d/yy">Date</option>
            </select>
            <IconButton label="Align left" active={selectedStyle.horizontal === 'left'} onClick={() => formatSelection({ horizontal: 'left' })}><AlignLeft size={16} /></IconButton>
            <IconButton label="Align center" active={selectedStyle.horizontal === 'center'} onClick={() => formatSelection({ horizontal: 'center' })}><AlignCenter size={16} /></IconButton>
            <IconButton label="Align right" active={selectedStyle.horizontal === 'right'} onClick={() => formatSelection({ horizontal: 'right' })}><AlignRight size={16} /></IconButton>
          </div>
          <div className="tool-group">
            <IconButton label="AutoSum" onClick={addSum}><Sigma size={17} /></IconButton>
            <div className="menu-anchor">
              <IconButton label="Rows and columns" onClick={() => setStructureMenuOpen((value) => !value)}><Settings2 size={16} /></IconButton>
              {structureMenuOpen ? (
                <div className="popover structure-menu">
                  <button onClick={() => { commit((draft) => insertRow(draft, draft.activeSheetId, selection.focus.row)); setStructureMenuOpen(false); }}><Rows3 size={15} /> Insert row above</button>
                  <button onClick={() => { commit((draft) => deleteRow(draft, draft.activeSheetId, selection.focus.row)); setStructureMenuOpen(false); }}><Trash2 size={15} /> Delete row</button>
                  <button onClick={() => { commit((draft) => insertColumn(draft, draft.activeSheetId, selection.focus.column)); setStructureMenuOpen(false); }}><Columns3 size={15} /> Insert column left</button>
                  <button onClick={() => { commit((draft) => deleteColumn(draft, draft.activeSheetId, selection.focus.column)); setStructureMenuOpen(false); }}><Trash2 size={15} /> Delete column</button>
                </div>
              ) : null}
            </div>
          </div>
          <div className="toolbar-spacer" />
          <IconButton label="Find" onClick={() => { setFindOpen(true); requestAnimationFrame(() => findRef.current?.focus()); }}><Search size={16} /></IconButton>
        </div>

        <div className="formula-row">
          <div className="name-box">{selectionLabel(selection)}</div>
          <FunctionSquare size={15} aria-hidden="true" />
          <input
            ref={formulaRef}
            value={editing ? editValue : editableCellText(activeCell)}
            aria-label="Formula bar"
            placeholder="Enter a value or formula"
            onFocus={() => { if (!editing) startEdit(); }}
            onChange={(event) => { if (!editing) startEdit(event.target.value); else setEditValue(event.target.value); }}
            onKeyDown={(event) => { if (event.key === 'Enter') { event.preventDefault(); commitEdit('down'); } if (event.key === 'Escape') { event.preventDefault(); cancelEdit(); } }}
          />
        </div>
      </header>

      {findOpen ? (
        <div className="find-bar">
          <Search size={15} />
          <input ref={findRef} value={findQuery} placeholder="Find in this sheet" onChange={(event) => { setFindQuery(event.target.value); setFindIndex(0); }} onKeyDown={(event) => { if (event.key === 'Enter') selectFindResult(findIndex + (event.shiftKey ? -1 : 1)); if (event.key === 'Escape') setFindOpen(false); }} />
          <span>{findQuery ? `${findMatches.length ? findIndex + 1 : 0} of ${findMatches.length}` : ''}</span>
          <IconButton label="Previous match" onClick={() => selectFindResult(findIndex - 1)}><ChevronLeft size={15} /></IconButton>
          <IconButton label="Next match" onClick={() => selectFindResult(findIndex + 1)}><ChevronRight size={15} /></IconButton>
          <IconButton label="Close find" onClick={() => setFindOpen(false)}><X size={15} /></IconButton>
        </div>
      ) : null}

      {gridMenu && <GridContextMenu target={gridMenu} items={gridMenuItems} onClose={closeGridMenu} />}
      <section className="workspace">
        <SpreadsheetGrid
          ref={gridRef}
          sheet={activeSheet}
          evaluator={evaluator}
          selection={selection}
          editing={editing}
          editValue={editValue}
          onEditValueChange={setEditValue}
          onCommitEdit={commitEdit}
          onCancelEdit={cancelEdit}
          onStartEdit={startEdit}
          onSelectionChange={(next) => { commitEdit(); setSelection(next); gridRef.current?.focus({ preventScroll: true }); }}
          onContextMenu={(target) => { commitEdit(); setSelection(target.selection); setGridMenu(target); }}
          onFillSelection={fillSelectedCells}
          onPickFormulaReference={pickFormulaReference}
          referencePicking={editing && editValue.trimStart().startsWith('=')}
          onColumnResize={(column, width) => commit((draft) => resizeColumn(draft.sheets.find((sheet) => sheet.id === draft.activeSheetId)!, column, width))}
          onRowResize={(row, height) => commit((draft) => resizeRow(draft.sheets.find((sheet) => sheet.id === draft.activeSheetId)!, row, height))}
          onKeyDown={gridKeyDown}
        />
        {compatibilityOpen && workbook.compatibilityIssues.length ? (
          <aside className="compatibility-panel">
            <div><div><strong>Compatibility</strong><span>Features detected when this file was opened</span></div><IconButton label="Close" onClick={() => setCompatibilityOpen(false)}><X size={15} /></IconButton></div>
            <p>Saving may change the features below. Use Save As to keep the original file unchanged.</p>
            <ul>{workbook.compatibilityIssues.map((issue, index) => <li key={`${issue.feature}-${index}`}><strong>{issue.feature}</strong><span>{issue.detail}</span></li>)}</ul>
            <div className="compatibility-report">
              <button type="button" onClick={() => void reportCompatibility()}><Bug size={14} /> Review report on GitHub</button>
              <small>Opens a draft GitHub issue containing compatibility categories only. Review it before submitting. Workbook names, sheet names, cells, formulas, comments, authors, and paths stay private.</small>
            </div>
          </aside>
        ) : null}
      </section>

      <footer className="bottom-bar">
        <div className="sheet-tabs" role="tablist" aria-label="Worksheets">
          {workbook.sheets.map((sheet) => (
            <button
              role="tab" aria-selected={sheet.id === workbook.activeSheetId}
              className={[sheet.id === workbook.activeSheetId ? 'is-active' : '', sheetDrop?.id === sheet.id ? `drop-${sheetDrop.side}` : ''].filter(Boolean).join(' ')}
              key={sheet.id} draggable={workbook.sheets.length > 1}
              title="Drag to reorder sheets; use the context menu to move with a keyboard"
              onDragStart={(event) => {
                draggedSheetId.current = sheet.id;
                event.dataTransfer.effectAllowed = 'move';
                event.dataTransfer.setData('text/plain', sheet.id);
              }}
              onDragOver={(event) => {
                if (!draggedSheetId.current) return;
                event.preventDefault();
                event.dataTransfer.dropEffect = 'move';
                const rect = event.currentTarget.getBoundingClientRect();
                const side = event.clientX < rect.left + rect.width / 2 ? 'before' : 'after';
                if (sheetDrop?.id !== sheet.id || sheetDrop.side !== side) setSheetDrop({ id: sheet.id, side });
              }}
              onDrop={(event) => {
                if (!draggedSheetId.current) return;
                event.preventDefault();
                const rect = event.currentTarget.getBoundingClientRect();
                dropSheet(sheet.id, event.clientX < rect.left + rect.width / 2 ? 'before' : 'after');
              }}
              onDragEnd={() => { draggedSheetId.current = null; setSheetDrop(null); }}
              onClick={() => {
                setHistory((current) => current.present.activeSheetId === sheet.id ? current : {
                  ...current, present: { ...current.present, activeSheetId: sheet.id },
                });
                setSelection(INITIAL_SELECTION);
                cancelEdit();
              }}
              onDoubleClick={(event) => setSheetMenu({ id: sheet.id, x: 0, y: 0, trigger: event.currentTarget, rename: true })}
              onContextMenu={(event) => {
                event.preventDefault();
                const rect = event.currentTarget.getBoundingClientRect();
                setSheetMenu({ id: sheet.id, x: event.clientX || rect.left, y: event.clientY || rect.top, trigger: event.currentTarget });
              }}
              onKeyDown={(event) => {
                if (event.key === 'ContextMenu' || (event.shiftKey && event.key === 'F10')) {
                  event.preventDefault();
                  const rect = event.currentTarget.getBoundingClientRect();
                  setSheetMenu({ id: sheet.id, x: rect.left, y: rect.top, trigger: event.currentTarget });
                }
              }}
            >{sheet.name}</button>
          ))}
          <IconButton label="Add sheet" onClick={addSheet}><Plus size={15} /></IconButton>
        </div>
        <div className="status-summary">
          <div className={`update-control is-${updateState.phase}`} aria-live="polite">
            <span>TXT Sheets v{updateState.currentVersion}</span>
            {updatePresentation.action ? (
              <button type="button" title={updatePresentation.detail} onClick={() => void runUpdateAction(updatePresentation.action)}>
                {updateState.phase === 'available' ? <Download size={11} aria-hidden="true" /> : null}
                {updateState.phase === 'downloaded' ? <RefreshCw size={11} aria-hidden="true" /> : null}
                {updatePresentation.actionLabel}
              </button>
            ) : updatePresentation.busy ? <span className="update-busy"><RefreshCw size={10} aria-hidden="true" />{updatePresentation.actionLabel}</span> : null}
          </div>
          <span className="selection-stat">{selectionStats.count ? `Count ${selectionStats.count}` : 'Ready'}</span>
          {selectionStats.count > 1 ? <><span className="selection-stat">Average {Number(selectionStats.average.toPrecision(8))}</span><span className="selection-stat">Sum {Number(selectionStats.sum.toPrecision(10))}</span></> : null}
          <span className="sheet-count">{workbook.sheets.length} {workbook.sheets.length === 1 ? 'sheet' : 'sheets'}</span>
        </div>
      </footer>

      {sheetMenu && workbook.sheets.some((sheet) => sheet.id === sheetMenu.id) && <SheetTabMenu
        key={`${sheetMenu.id}-${sheetMenu.x}-${sheetMenu.y}-${sheetMenu.rename}`}
        target={sheetMenu}
        name={workbook.sheets.find((sheet) => sheet.id === sheetMenu.id)!.name}
        canMoveLeft={workbook.sheets.findIndex((sheet) => sheet.id === sheetMenu.id) > 0}
        canMoveRight={workbook.sheets.findIndex((sheet) => sheet.id === sheetMenu.id) < workbook.sheets.length - 1}
        canDelete={workbook.sheets.length > 1}
        onClose={() => setSheetMenu(null)}
        onRename={(name) => {
          const problem = renameWorksheet(cloneWorkbook(workbook), sheetMenu.id, name);
          if (problem) return problem;
          commit((draft) => { renameWorksheet(draft, sheetMenu.id, name); });
          return null;
        }}
        onMove={(direction) => commit((draft) => {
          const index = draft.sheets.findIndex((sheet) => sheet.id === sheetMenu.id);
          const next = index + direction;
          if (index < 0 || next < 0 || next >= draft.sheets.length) return;
          reorderWorksheet(draft, sheetMenu.id, next + (direction > 0 ? 1 : 0));
        })}
        onAdd={addSheet}
        onDelete={() => deleteSheet(sheetMenu.id)}
      />}
      {message ? <div className="toast" role="status">{message}</div> : null}
    </main>
  );
}
