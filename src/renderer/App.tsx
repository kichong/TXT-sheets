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
import { SpreadsheetGrid } from './SpreadsheetGrid';
import {
  addressForCell, cellKey, createFormulaEvaluator, editableCellText, isDateNumberFormat, normalizeCellInput,
} from './formulas';
import {
  applyStyle, cloneWorkbook, deleteColumn, deleteRow, fillSelection, insertColumn, insertRow, nearestCellTemplate, selectedCells,
  resizeColumn, resizeRow, selectionBounds, selectionLabel, uniqueSheetName,
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
  const activeSheet = workbook.sheets.find((sheet) => sheet.id === workbook.activeSheetId) ?? workbook.sheets[0];
  const evaluator = useMemo(() => createFormulaEvaluator(workbook), [workbook]);
  const activeCell = activeSheet.cells[cellKey(selection.focus.row, selection.focus.column)];

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

  const saveWorkbook = useCallback(async (saveAs = false) => {
    const focusTarget = document.activeElement instanceof HTMLInputElement ? document.activeElement : gridRef.current;
    setSaving(true);
    try {
      const prepared = workbookWithFormulaResults(workbook);
      const result = saveAs ? await window.spreadsheet.saveAs(prepared) : await window.spreadsheet.save(prepared);
      if (result.status === 'saved' && result.source) {
        const next = { ...prepared, source: result.source, title: result.source.displayName.replace(/\.(xlsx|csv|tsv)$/iu, '') };
        setHistory((current) => ({ ...current, present: next }));
        setRecentFiles(result.recentFiles);
        setDirty(false);
        setMessage(saveAs ? `Saved as ${result.source.displayName}` : `Saved ${result.source.displayName}`);
      }
    } catch (error) { handleError(error); }
    finally {
      setSaving(false);
      requestAnimationFrame(() => {
        if (focusTarget?.isConnected) focusTarget.focus();
        else gridRef.current?.focus();
      });
    }
  }, [handleError, workbook]);

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
    replaceWorkbook(createBlankWorkbook());
    void window.spreadsheet.clearRecovery();
  }, [replaceWorkbook]);

  const runCommand = useCallback((command: AppCommand) => {
    if (command === 'new') newWorkbook();
    if (command === 'open') void openWorkbook();
    if (command === 'save') void saveWorkbook(false);
    if (command === 'save-as') void saveWorkbook(true);
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
    window.spreadsheet.setDirty(dirty);
    if (!dirty) return;
    const timer = window.setTimeout(() => { void window.spreadsheet.writeRecovery(workbookWithFormulaResults(workbook)).catch(handleError); }, 800);
    return () => window.clearTimeout(timer);
  }, [dirty, handleError, workbook]);

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
    if (!editing) return;
    const point = selection.focus;
    commit((draft) => {
      const sheet = draft.sheets.find((item) => item.id === draft.activeSheetId)!;
      const key = cellKey(point.row, point.column);
      const existing = sheet.cells[key];
      const template = existing?.style ? existing : nearestCellTemplate(sheet, point.row, point.column) ?? existing;
      const normalized = normalizeCellInput(editValue, template);
      const inheritedStyle = existing?.style ?? template?.style;
      if (normalized) sheet.cells[key] = { ...normalized, style: inheritedStyle ? structuredClone(inheritedStyle) : undefined };
      else if (inheritedStyle) sheet.cells[key] = { value: null, valueType: 'blank', style: structuredClone(inheritedStyle) };
      else delete sheet.cells[key];
    });
    setEditing(false);
    if (move) {
      const focus = {
        row: Math.min(activeSheet.rowCount - 1, point.row + (move === 'down' ? 1 : 0)),
        column: Math.min(activeSheet.columnCount - 1, point.column + (move === 'right' ? 1 : 0)),
      };
      setSelection({ anchor: focus, focus });
      requestAnimationFrame(() => gridRef.current?.focus());
    }
  }, [activeSheet.columnCount, activeSheet.rowCount, commit, editValue, editing, selection.focus]);

  const startEdit = useCallback((initial?: string) => {
    setEditValue(initial ?? editableCellText(activeCell));
    setEditing(true);
  }, [activeCell]);

  const clearSelection = useCallback(() => {
    commit((draft) => {
      const sheet = draft.sheets.find((item) => item.id === draft.activeSheetId)!;
      selectedCells(sheet, selection).forEach(({ row, column, cell }) => {
        const key = cellKey(row, column);
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
    await navigator.clipboard.writeText(rows.join('\r\n'));
    if (cut) clearSelection();
  }, [activeSheet.cells, clearSelection, selection]);

  const pasteSelection = useCallback(async () => {
    try {
      const text = await navigator.clipboard.readText();
      const rows = text.replace(/\r/gu, '').split('\n').map((row) => row.split('\t'));
      const origin = selection.focus;
      commit((draft) => {
        const sheet = draft.sheets.find((item) => item.id === draft.activeSheetId)!;
        rows.forEach((values, rowOffset) => values.forEach((value, columnOffset) => {
          const row = origin.row + rowOffset;
          const column = origin.column + columnOffset;
          if (row >= sheet.rowCount || column >= sheet.columnCount) return;
          const key = cellKey(row, column);
          const existing = sheet.cells[key];
          const template = existing?.style ? existing : nearestCellTemplate(sheet, row, column) ?? existing;
          const normalized = normalizeCellInput(value, template);
          const inheritedStyle = existing?.style ?? template?.style;
          if (normalized) sheet.cells[key] = { ...normalized, style: inheritedStyle ? structuredClone(inheritedStyle) : undefined };
          else if (inheritedStyle) sheet.cells[key] = { value: null, valueType: 'blank', style: structuredClone(inheritedStyle) };
          else delete sheet.cells[key];
        }));
      });
      const focus = { row: Math.min(activeSheet.rowCount - 1, origin.row + rows.length - 1), column: Math.min(activeSheet.columnCount - 1, origin.column + Math.max(0, ...rows.map((row) => row.length - 1))) };
      setSelection({ anchor: origin, focus });
    } catch (error) { handleError(error); }
  }, [activeSheet.columnCount, activeSheet.rowCount, commit, handleError, selection.focus]);

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
    const values = selectedCells(activeSheet, selection)
      .map(({ row, column }) => evaluator.evaluateCell(activeSheet.id, row, column))
      .filter((value): value is number => typeof value === 'number');
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

  const renameSheet = useCallback((sheetId: string) => {
    const sheet = workbook.sheets.find((item) => item.id === sheetId);
    if (!sheet) return;
    const name = window.prompt('Sheet name', sheet.name)?.trim();
    if (!name || workbook.sheets.some((item) => item.id !== sheetId && item.name.toLocaleLowerCase() === name.toLocaleLowerCase())) return;
    commit((draft) => { draft.sheets.find((item) => item.id === sheetId)!.name = name; });
  }, [commit, workbook.sheets]);

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
            <label className="color-control" title="Text color"><span>A</span><input type="color" value={selectedStyle.textColor ?? '#202124'} onChange={(event) => formatSelection({ textColor: event.target.value })} /></label>
            <label className="color-control fill-control" title="Fill color"><span /><input type="color" value={selectedStyle.fillColor ?? '#fff4be'} onChange={(event) => formatSelection({ fillColor: event.target.value })} /></label>
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
                  <button onClick={() => { commit((draft) => insertRow(draft.sheets.find((sheet) => sheet.id === draft.activeSheetId)!, selection.focus.row)); setStructureMenuOpen(false); }}><Rows3 size={15} /> Insert row above</button>
                  <button onClick={() => { commit((draft) => deleteRow(draft.sheets.find((sheet) => sheet.id === draft.activeSheetId)!, selection.focus.row)); setStructureMenuOpen(false); }}><Trash2 size={15} /> Delete row</button>
                  <button onClick={() => { commit((draft) => insertColumn(draft.sheets.find((sheet) => sheet.id === draft.activeSheetId)!, selection.focus.column)); setStructureMenuOpen(false); }}><Columns3 size={15} /> Insert column left</button>
                  <button onClick={() => { commit((draft) => deleteColumn(draft.sheets.find((sheet) => sheet.id === draft.activeSheetId)!, selection.focus.column)); setStructureMenuOpen(false); }}><Trash2 size={15} /> Delete column</button>
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
            onKeyDown={(event) => { if (event.key === 'Enter') { event.preventDefault(); commitEdit('down'); } if (event.key === 'Escape') { setEditing(false); gridRef.current?.focus(); } }}
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
          onCancelEdit={() => { setEditing(false); gridRef.current?.focus(); }}
          onStartEdit={startEdit}
          onSelectionChange={(next) => { setSelection(next); setEditing(false); gridRef.current?.focus(); }}
          onFillSelection={fillSelectedCells}
          onPickFormulaReference={pickFormulaReference}
          referencePicking={editing && editValue.trimStart().startsWith('=')}
          onColumnResize={(column, width) => commit((draft) => resizeColumn(draft.sheets.find((sheet) => sheet.id === draft.activeSheetId)!, column, width))}
          onRowResize={(row, height) => commit((draft) => resizeRow(draft.sheets.find((sheet) => sheet.id === draft.activeSheetId)!, row, height))}
          onKeyDown={gridKeyDown}
        />
        {compatibilityOpen && workbook.compatibilityIssues.length ? (
          <aside className="compatibility-panel">
            <div><div><strong>Compatibility notes</strong><span>Detected when this workbook was opened.</span></div><IconButton label="Close" onClick={() => setCompatibilityOpen(false)}><X size={15} /></IconButton></div>
            <p>Save updates the current file. Save As creates a separate file. Features below may change if TXT Sheets cannot preserve them yet.</p>
            <ul>{workbook.compatibilityIssues.map((issue, index) => <li key={`${issue.feature}-${index}`}><strong>{issue.feature}</strong><span>{issue.detail}</span></li>)}</ul>
            <div className="compatibility-report">
              <button type="button" onClick={() => void reportCompatibility()}><Bug size={14} /> Report these issues</button>
              <small>Opens a prefilled GitHub issue with app and system details only—never the file name or contents.</small>
            </div>
          </aside>
        ) : null}
      </section>

      <footer className="bottom-bar">
        <div className="sheet-tabs" role="tablist" aria-label="Worksheets">
          {workbook.sheets.map((sheet) => (
            <button
              role="tab" aria-selected={sheet.id === workbook.activeSheetId} className={sheet.id === workbook.activeSheetId ? 'is-active' : ''} key={sheet.id}
              onClick={() => { commit((draft) => { draft.activeSheetId = sheet.id; }); setSelection(INITIAL_SELECTION); setDirty(dirty); }}
              onDoubleClick={() => renameSheet(sheet.id)}
              onContextMenu={(event) => { event.preventDefault(); deleteSheet(sheet.id); }}
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

      {message ? <div className="toast" role="status">{message}</div> : null}
    </main>
  );
}
