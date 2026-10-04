import { useEffect, useMemo, useRef, useState } from 'react';
import type { ClipboardEvent, CSSProperties, KeyboardEvent } from 'react';
import { AlignCenter, AlignLeft, AlignRight, ArrowDown, ArrowDownAZ, ArrowLeft, ArrowRight, ArrowUpAZ, Bold, Check, ChevronDown, Copy, Download, FileSpreadsheet, History, Italic, Keyboard, PaintBucket, Plus, Redo2, Search, Sigma, Snowflake, Trash2, Underline, Undo2, X } from 'lucide-react';
import type { CellStyle, EditorProps, SheetContent, SheetTab } from '../types';
import { normalizeWorkbook, packWorkbook, renameSheetReferences } from '../lib/spreadsheetModel';
import { exportWorkbook } from '../lib/workbookIO';
import { calculateCells, shiftFormulaReferences } from './spreadsheetFormula';
import './spreadsheet-editor.css';

export { calculateCells } from './spreadsheetFormula';
type Point = { row: number; col: number };
type Workbook = { sheets: SheetTab[]; activeSheetId: string };
type SheetDialog = { mode: 'rename' | 'delete'; id: string; value: string };
const columnName = (n: number): string => { let s = ''; for (n++; n; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
const cellKey = ({ row, col }: Point) => `${columnName(col)}${row + 1}`;
const pointFor = (key: string): Point => { const match = key.replaceAll('$', '').match(/^([A-Z]+)(\d+)$/i); if (!match || Number(match[2]) < 1) throw Error('#REF!'); return { col: [...match[1].toUpperCase()].reduce((v, c) => v * 26 + c.charCodeAt(0) - 64, 0) - 1, row: Number(match[2]) - 1 }; };
const sheetId = () => crypto.randomUUID();
const escapeRegex = (value: string) => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const formatValue = (value: string | number | boolean, style: CellStyle = {}) => {
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (value === '' || String(value).startsWith('#')) return value;
  const decimals = style.decimals ?? (style.numberFormat === 'percentage' ? 0 : 2);
  if (style.numberFormat === 'date') {
    const date = typeof value === 'number' ? new Date(Date.UTC(1899, 11, value < 60 ? 31 : 30) + value * 86400000) : new Date(value);
    return Number.isNaN(date.getTime()) ? value : date.toLocaleDateString(undefined, { timeZone: 'UTC' });
  }
  if (typeof value !== 'number') return value;
  if (style.numberFormat === 'number') return value.toLocaleString(undefined, { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
  if (style.numberFormat === 'currency') return value.toLocaleString(undefined, { style: 'currency', currency: 'EUR', minimumFractionDigits: decimals, maximumFractionDigits: decimals });
  if (style.numberFormat === 'percentage') return value.toLocaleString(undefined, { style: 'percent', minimumFractionDigits: decimals, maximumFractionDigits: decimals });
  return value;
};

export default function SpreadsheetEditor({ file, onChange, onRename, onBack, onNotify, onOpenVersions, saveState = 'saved' }: EditorProps) {
  const [workbook, setWorkbook] = useState<Workbook>(() => normalizeWorkbook(file.content as SheetContent));
  const workbookRef = useRef(workbook);
  const [title, setTitle] = useState(file.name);
  const [active, setActive] = useState<Point>({ row: 0, col: 0 });
  const [anchor, setAnchor] = useState<Point | null>(null);
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState('');
  const editRef = useRef<{ sheetId: string; key: string; value: string } | null>(null);
  const [rowCount, setRowCount] = useState(40);
  const [colCount, setColCount] = useState(12);
  const [ribbon, setRibbon] = useState<'home' | 'data' | 'view'>('home');
  const [exportOpen, setExportOpen] = useState(false);
  const [helpOpen, setHelpOpen] = useState(false);
  const [findOpen, setFindOpen] = useState(false);
  const [search, setSearch] = useState('');
  const [replacement, setReplacement] = useState('');
  const [filter, setFilter] = useState<{ col: number; text: string } | null>(null);
  const [protectHeader, setProtectHeader] = useState(true);
  const [showFormulas, setShowFormulas] = useState(false);
  const [showGridlines, setShowGridlines] = useState(true);
  const [zoom, setZoom] = useState(100);
  const [sheetDialog, setSheetDialog] = useState<SheetDialog | null>(null);
  const [dialogError, setDialogError] = useState('');
  const [sortConfirmation, setSortConfirmation] = useState<1 | -1 | null>(null);
  const [resizing, setResizing] = useState<{ col: number; width: number } | null>(null);
  const [undoStack, setUndoStack] = useState<Workbook[]>([]);
  const [redoStack, setRedoStack] = useState<Workbook[]>([]);
  const undoRef = useRef<Workbook[]>([]);
  const redoRef = useRef<Workbook[]>([]);
  const resizeCleanup = useRef<(() => void) | null>(null);
  const inputs = useRef(new Map<string, HTMLInputElement>());
  const formulaRef = useRef<HTMLInputElement>(null);
  const findRef = useRef<HTMLInputElement>(null);
  const sheet = workbook.sheets.find(item => item.id === workbook.activeSheetId) || workbook.sheets[0];
  const evaluate = useMemo(() => calculateCells(sheet.cells, { sheets: workbook.sheets, currentSheetId: sheet.id }), [sheet.cells, sheet.id, workbook.sheets]);
  const activeKey = cellKey(active);
  const activeStyle = sheet.styles?.[activeKey] || {};
  const lastUsed = useMemo(() => [...Object.keys(sheet.cells), ...Object.keys(sheet.styles || {})].reduce((acc, key) => { try { const point = pointFor(key); return { row: Math.max(acc.row, point.row), col: Math.max(acc.col, point.col) }; } catch { return acc; } }, { row: 0, col: 0 }), [sheet.cells, sheet.styles]);
  const rows = Math.min(2000, Math.max(rowCount, lastUsed.row + 1));
  const cols = Math.min(100, Math.max(colCount, lastUsed.col + 1));
  const range = { firstRow: Math.min(active.row, anchor?.row ?? active.row), lastRow: Math.max(active.row, anchor?.row ?? active.row), firstCol: Math.min(active.col, anchor?.col ?? active.col), lastCol: Math.max(active.col, anchor?.col ?? active.col) };
  const rangeLabel = anchor ? `${cellKey({ row: range.firstRow, col: range.firstCol })}:${cellKey({ row: range.lastRow, col: range.lastCol })}` : activeKey;
  const selectedKeys: string[] = [];
  for (let row = range.firstRow; row <= range.lastRow; row++) for (let col = range.firstCol; col <= range.lastCol; col++) selectedKeys.push(cellKey({ row, col }));
  const numericValues = selectedKeys.map(evaluate).filter((v): v is number => typeof v === 'number');
  const matches = useMemo(() => search.trim() ? Object.keys(sheet.cells).filter(key => String(evaluate(key)).toLowerCase().includes(search.toLowerCase()) || sheet.cells[key].toLowerCase().includes(search.toLowerCase())) : [], [sheet.cells, evaluate, search]);
  const matchSet = useMemo(() => new Set(matches), [matches]);
  const visibleRows = useMemo(() => Array.from({ length: rows }, (_, row) => row).filter(row => !filter?.text || (protectHeader && row === 0) || String(evaluate(cellKey({ row, col: filter.col }))).toLowerCase().includes(filter.text.toLowerCase())), [rows, filter, protectHeader, evaluate]);

  useEffect(() => { const next = normalizeWorkbook(file.content as SheetContent); workbookRef.current = next; setWorkbook(next); setTitle(file.name); setActive({ row: 0, col: 0 }); setAnchor(null); undoRef.current = []; redoRef.current = []; setUndoStack([]); setRedoStack([]); editRef.current = null; setEditing(false); setFilter(null); setRowCount(40); setColCount(12); }, [file.id]);
  useEffect(() => () => resizeCleanup.current?.(), []);

  const updateWorkbook = (next: Workbook, record = true) => {
    if (record) { undoRef.current = [...undoRef.current.slice(-79), workbookRef.current]; redoRef.current = []; setUndoStack(undoRef.current); setRedoStack([]); }
    workbookRef.current = next; setWorkbook(next); onChange(packWorkbook(next.sheets, next.activeSheetId));
  };
  const currentSheet = () => workbookRef.current.sheets.find(item => item.id === workbookRef.current.activeSheetId) || workbookRef.current.sheets[0];
  const updateSheet = (next: SheetTab) => updateWorkbook({ ...workbookRef.current, sheets: workbookRef.current.sheets.map(item => item.id === next.id ? next : item) });
  const commit = () => {
    const pending = editRef.current;
    if (!pending) return;
    editRef.current = null; setEditing(false);
    const current = workbookRef.current.sheets.find(item => item.id === pending.sheetId);
    if (!current || (current.cells[pending.key] || '') === pending.value) return;
    const cells = { ...current.cells };
    if (pending.value) cells[pending.key] = pending.value; else delete cells[pending.key];
    updateSheet({ ...current, cells });
  };
  const beginEdit = (value = currentSheet().cells[activeKey] || '') => { editRef.current = { sheetId: currentSheet().id, key: activeKey, value }; setDraft(value); setEditing(true); };
  const changeDraft = (value: string) => { editRef.current = { sheetId: currentSheet().id, key: activeKey, value }; setDraft(value); setEditing(true); };
  const select = (point: Point, extend = false) => { commit(); if (extend) setAnchor(previous => previous || active); else setAnchor(null); setActive(point); };
  const focus = (point: Point, extend = false) => { const next = { row: Math.max(0, Math.min(rows - 1, point.row)), col: Math.max(0, Math.min(cols - 1, point.col)) }; select(next, extend); requestAnimationFrame(() => inputs.current.get(cellKey(next))?.focus()); };
  const clear = (formatting = false) => { commit(); const current = currentSheet(); if (formatting) { const styles = { ...current.styles }; selectedKeys.forEach(key => delete styles[key]); updateSheet({ ...current, styles }); } else { const cells = { ...current.cells }; selectedKeys.forEach(key => delete cells[key]); updateSheet({ ...current, cells }); } };
  const undo = () => { commit(); if (!undoRef.current.length) return; const previous = undoRef.current.at(-1)!; redoRef.current = [...redoRef.current, workbookRef.current]; undoRef.current = undoRef.current.slice(0, -1); setRedoStack(redoRef.current); setUndoStack(undoRef.current); setFilter(null); updateWorkbook(previous, false); };
  const redo = () => { if (!redoRef.current.length) return; editRef.current = null; setEditing(false); const next = redoRef.current.at(-1)!; undoRef.current = [...undoRef.current, workbookRef.current]; redoRef.current = redoRef.current.slice(0, -1); setUndoStack(undoRef.current); setRedoStack(redoRef.current); setFilter(null); updateWorkbook(next, false); };
  const applyStyle = (patch: Partial<CellStyle>) => { commit(); const current = currentSheet(); const styles = { ...current.styles }; selectedKeys.forEach(key => { styles[key] = { ...styles[key], ...patch }; }); updateSheet({ ...current, styles }); };
  const moveRow = (row: number, direction: number) => {
    if (!filter?.text) return row + direction;
    const index = visibleRows.indexOf(row);
    return visibleRows[Math.max(0, Math.min(visibleRows.length - 1, index + direction))] ?? row;
  };
  const keyboard = (event: KeyboardEvent<HTMLInputElement>, point: Point) => {
    const modifier = event.ctrlKey || event.metaKey;
    if (modifier && event.key.toLowerCase() === 'z') { event.preventDefault(); event.shiftKey ? redo() : undo(); return; }
    if (modifier && event.key.toLowerCase() === 'y') { event.preventDefault(); redo(); return; }
    if (modifier && event.key.toLowerCase() === 'f') { event.preventDefault(); setFindOpen(true); requestAnimationFrame(() => findRef.current?.focus()); return; }
    if (modifier && !editing && ['b', 'i', 'u'].includes(event.key.toLowerCase())) { event.preventDefault(); const prop = { b: 'bold', i: 'italic', u: 'underline' }[event.key.toLowerCase()] as 'bold' | 'italic' | 'underline'; applyStyle({ [prop]: !activeStyle[prop] }); return; }
    if (modifier && !editing && event.key.toLowerCase() === 'a') { event.preventDefault(); setAnchor({ row: 0, col: 0 }); setActive({ row: Math.min(1999, lastUsed.row), col: Math.min(99, lastUsed.col) }); return; }
    if (modifier && !editing && ['d', 'r'].includes(event.key.toLowerCase())) { event.preventDefault(); fill(event.key.toLowerCase() === 'd' ? 'down' : 'right'); return; }
    if (event.key === 'Escape') { editRef.current = null; setEditing(false); setAnchor(null); return; }
    if (event.key === 'Enter') { event.preventDefault(); focus({ ...point, row: moveRow(point.row, event.shiftKey ? -1 : 1) }); return; }
    if (event.key === 'Tab') { event.preventDefault(); const col = point.col + (event.shiftKey ? -1 : 1); focus({ row: point.row + (col >= cols ? 1 : col < 0 ? -1 : 0), col: (col + cols) % cols }); return; }
    if (editing) return;
    if (event.key.startsWith('Arrow')) { event.preventDefault(); const delta = { ArrowDown: [1, 0], ArrowUp: [-1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[event.key]!; focus({ row: delta[0] ? moveRow(point.row, delta[0]) : point.row, col: point.col + delta[1] }, event.shiftKey); }
    else if (event.key === 'Delete' || event.key === 'Backspace') { event.preventDefault(); clear(); }
    else if (event.key === 'F2') { event.preventDefault(); beginEdit(); }
    else if (event.key.length === 1 && !modifier && !event.altKey) { event.preventDefault(); beginEdit(event.key); }
  };
  const copy = (event: ClipboardEvent<HTMLDivElement>) => {
    if (editing) return;
    const lines: string[] = []; const values: string[][] = []; const styles: CellStyle[][] = [];
    for (let row = range.firstRow; row <= range.lastRow; row++) { const line: string[] = []; const raw: string[] = []; const rowStyles: CellStyle[] = []; for (let col = range.firstCol; col <= range.lastCol; col++) { const key = cellKey({ row, col }); line.push(String(evaluate(key))); raw.push(sheet.cells[key] || ''); rowStyles.push(sheet.styles?.[key] || {}); } lines.push(line.join('\t')); values.push(raw); styles.push(rowStyles); }
    event.clipboardData.setData('text/plain', lines.join('\n'));
    event.clipboardData.setData('application/x-folio-cells', JSON.stringify({ version: 1, row: range.firstRow, col: range.firstCol, values, styles })); event.preventDefault();
  };
  const paste = (event: ClipboardEvent<HTMLDivElement>) => {
    const text = event.clipboardData.getData('text/plain');
    if (editing && !text.includes('\t') && !text.includes('\n')) return;
    event.preventDefault(); commit();
    let data = text.replace(/\r/g, '').replace(/\n$/, '').split('\n').map(row => row.split('\t'));
    let copied: { row: number; col: number; values: string[][]; styles?: CellStyle[][] } | null = null;
    try { const payload = JSON.parse(event.clipboardData.getData('application/x-folio-cells')); if (payload?.version === 1 && Array.isArray(payload.values) && payload.values.every((row: unknown) => Array.isArray(row) && row.every(value => typeof value === 'string')) && Number.isInteger(payload.row) && Number.isInteger(payload.col)) { copied = payload; data = payload.values; } } catch { /* Standard clipboard text. */ }
    const current = currentSheet(); const cells = { ...current.cells }; const styles = { ...current.styles };
    const pasted = data.slice(0, 2000 - active.row);
    pasted.forEach((values, row) => values.slice(0, 100 - active.col).forEach((value, col) => { const key = cellKey({ row: active.row + row, col: active.col + col }); const shifted = copied && value.startsWith('=') ? shiftFormulaReferences(value, active.row - copied.row, active.col - copied.col) : value; if (shifted) cells[key] = shifted; else delete cells[key]; if (copied?.styles?.[row]?.[col]) styles[key] = { ...copied.styles[row][col] }; }));
    updateSheet({ ...current, cells, styles }); setRowCount(count => Math.max(count, active.row + pasted.length)); setColCount(count => Math.max(count, Math.min(100, active.col + Math.max(1, ...pasted.map(row => row.length)))));
    setAnchor({ row: Math.min(1999, active.row + pasted.length - 1), col: Math.min(99, active.col + Math.max(1, ...pasted.map(row => row.length)) - 1) });
  };
  const fill = (direction: 'down' | 'right') => {
    commit(); if ((direction === 'down' && range.firstRow === range.lastRow) || (direction === 'right' && range.firstCol === range.lastCol)) { onNotify(`Select a range with multiple ${direction === 'down' ? 'rows' : 'columns'} first.`); return; }
    const current = currentSheet(); const cells = { ...current.cells }; const styles = { ...current.styles };
    for (let row = range.firstRow; row <= range.lastRow; row++) for (let col = range.firstCol; col <= range.lastCol; col++) {
      const source = { row: direction === 'down' ? range.firstRow : row, col: direction === 'right' ? range.firstCol : col }; const sourceKey = cellKey(source); const targetKey = cellKey({ row, col }); const raw = current.cells[sourceKey] || '';
      if (raw) cells[targetKey] = raw.startsWith('=') ? shiftFormulaReferences(raw, row - source.row, col - source.col) : raw; else delete cells[targetKey];
      if (current.styles?.[sourceKey]) styles[targetKey] = { ...current.styles[sourceKey] }; else delete styles[targetKey];
    }
    updateSheet({ ...current, cells, styles });
  };
  const sortRows = (direction: 1 | -1, confirmed = false) => {
    commit(); const current = currentSheet();
    if (!confirmed && workbookRef.current.sheets.some(item => Object.values(item.cells).some(value => value.startsWith('=')))) { setSortConfirmation(direction); return; }
    setSortConfirmation(null); const start = protectHeader ? 1 : 0;
    const usedLastRow = Object.keys(current.cells).reduce((max, key) => { try { return Math.max(max, pointFor(key).row); } catch { return max; } }, 0);
    if (usedLastRow < start) return;
    const getValue = calculateCells(current.cells, { sheets: workbookRef.current.sheets, currentSheetId: current.id });
    const order = Array.from({ length: usedLastRow - start + 1 }, (_, i) => i + start).sort((left, right) => { const a = getValue(cellKey({ row: left, col: active.col })); const b = getValue(cellKey({ row: right, col: active.col })); if (a === '' || b === '') return a === b ? 0 : a === '' ? 1 : -1; return direction * (typeof a === 'number' && typeof b === 'number' ? a - b : String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: 'base' })); });
    const rowMap = new Map(order.map((row, index) => [row, index + start])); const cells: Record<string, string> = {}; const styles: Record<string, CellStyle> = {};
    Object.entries(current.cells).forEach(([key, value]) => { const point = pointFor(key); const row = rowMap.get(point.row) ?? point.row; cells[cellKey({ ...point, row })] = value.startsWith('=') ? shiftFormulaReferences(value, row - point.row, 0) : value; });
    Object.entries(current.styles || {}).forEach(([key, value]) => { const point = pointFor(key); styles[cellKey({ ...point, row: rowMap.get(point.row) ?? point.row })] = value; });
    updateSheet({ ...current, cells, styles }); onNotify(`Rows sorted by column ${columnName(active.col)}${protectHeader ? '; header kept in place' : ''}.`);
  };
  const exportSheet = async (format: 'xlsx' | 'csv') => { commit(); try { const current = workbookRef.current; await exportWorkbook(packWorkbook(current.sheets, current.activeSheetId), title.trim() || file.name, format); setExportOpen(false); onNotify(`${format.toUpperCase()} downloaded`); } catch { onNotify('The spreadsheet could not be exported. Please try again.'); } };
  const insertSum = () => { commit(); const formula = anchor ? `=SUM(${rangeLabel})` : active.row > 0 ? `=SUM(${columnName(active.col)}1:${columnName(active.col)}${active.row})` : '=SUM()'; if (anchor) { const target = { row: range.lastRow + 1, col: range.firstCol }; if (target.row >= 2000) { onNotify('The visible grid supports up to 2,000 rows.'); return; } if (target.row >= rows) setRowCount(Math.min(2000, target.row + 10)); setActive(target); setAnchor(null); editRef.current = { sheetId: sheet.id, key: cellKey(target), value: formula }; setDraft(formula); setEditing(true); } else beginEdit(formula); requestAnimationFrame(() => formulaRef.current?.focus()); };
  const findNext = () => { if (!matches.length) return; const next = matches[(matches.indexOf(activeKey) + 1) % matches.length]; setFilter(null); focus(pointFor(next)); };
  const replace = (all: boolean) => { if (!search) return; commit(); const current = currentSheet(); const keys = all ? Object.keys(current.cells) : [activeKey]; const pattern = new RegExp(escapeRegex(search), 'gi'); let count = 0; const cells = { ...current.cells }; keys.forEach(key => { const value = cells[key]; if (value === undefined) return; const next = value.replace(pattern, () => { count++; return replacement; }); if (next) cells[key] = next; else delete cells[key]; }); if (count) updateSheet({ ...current, cells }); onNotify(`${count} replacement${count === 1 ? '' : 's'} made in cell contents.`); };
  const switchSheet = (id: string) => { commit(); updateWorkbook({ ...workbookRef.current, activeSheetId: id }, false); setActive({ row: 0, col: 0 }); setAnchor(null); setFilter(null); setRowCount(40); setColCount(12); };
  const uniqueName = (base: string) => { const names = new Set(workbookRef.current.sheets.map(item => item.name.toLowerCase())); let name = base.slice(0, 31); let index = 2; while (names.has(name.toLowerCase())) { const suffix = ` (${index++})`; name = `${base.slice(0, 31 - suffix.length)}${suffix}`; } return name; };
  const addSheet = (duplicate = false) => { commit(); const current = currentSheet(); const next: SheetTab = duplicate ? { ...current, id: sheetId(), name: uniqueName(`${current.name} copy`), cells: Object.fromEntries(Object.entries(current.cells).map(([key, value]) => [key, renameSheetReferences(value, current.name, uniqueName(`${current.name} copy`))])) } : { id: sheetId(), name: uniqueName(`Sheet ${workbookRef.current.sheets.length + 1}`), cells: {}, styles: {} }; updateWorkbook({ sheets: [...workbookRef.current.sheets, next], activeSheetId: next.id }); setActive({ row: 0, col: 0 }); setAnchor(null); setFilter(null); setRowCount(40); setColCount(12); };
  const submitSheetDialog = () => {
    if (!sheetDialog) return; commit(); const current = workbookRef.current; const target = current.sheets.find(item => item.id === sheetDialog.id); if (!target) { setSheetDialog(null); return; }
    if (sheetDialog.mode === 'delete') { const sheets = current.sheets.filter(item => item.id !== target.id); if (!sheets.length) return; updateWorkbook({ sheets, activeSheetId: current.activeSheetId === target.id ? sheets[0].id : current.activeSheetId }); setFilter(null); setActive({ row: 0, col: 0 }); setAnchor(null); }
    else { const name = sheetDialog.value.trim(); if (!name || name.length > 31 || /[\\/?*\[\]:]/.test(name) || name.startsWith("'") || name.endsWith("'")) { setDialogError('Use 1–31 characters, without \\ / ? * [ ] : or leading/trailing apostrophes.'); return; } if (current.sheets.some(item => item.id !== target.id && item.name.toLowerCase() === name.toLowerCase())) { setDialogError('A sheet with this name already exists.'); return; } updateWorkbook({ ...current, sheets: current.sheets.map(item => ({ ...item, name: item.id === target.id ? name : item.name, cells: Object.fromEntries(Object.entries(item.cells).map(([key, value]) => [key, renameSheetReferences(value, target.name, name)])) })) }); }
    setSheetDialog(null); setDialogError('');
  };
  const startResize = (event: React.PointerEvent<HTMLSpanElement>, col: number) => {
    event.preventDefault(); event.stopPropagation(); commit(); resizeCleanup.current?.(); const startX = event.clientX; const startWidth = currentSheet().columnWidths?.[columnName(col)] || 122; let width = startWidth;
    const move = (next: PointerEvent) => { width = Math.max(56, Math.min(500, Math.round(startWidth + (next.clientX - startX) / (zoom / 100)))); setResizing({ col, width }); };
    const cleanup = () => { window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', finish); window.removeEventListener('pointercancel', cancel); document.body.style.cursor = ''; document.body.style.userSelect = ''; resizeCleanup.current = null; };
    const finish = () => { cleanup(); setResizing(null); const current = currentSheet(); updateSheet({ ...current, columnWidths: { ...current.columnWidths, [columnName(col)]: width } }); };
    const cancel = () => { cleanup(); setResizing(null); };
    resizeCleanup.current = cleanup; document.body.style.cursor = 'col-resize'; document.body.style.userSelect = 'none'; window.addEventListener('pointermove', move); window.addEventListener('pointerup', finish); window.addEventListener('pointercancel', cancel);
  };
  const autoFit = (col: number) => { commit(); const current = currentSheet(); const length = Object.keys(current.cells).filter(key => pointFor(key).col === col).reduce((longest, key) => Math.max(longest, String(formatValue(evaluate(key), current.styles?.[key])).length), 8); updateSheet({ ...current, columnWidths: { ...current.columnWidths, [columnName(col)]: Math.max(72, Math.min(500, length * 8 + 24)) } }); };
  const status = saveState === 'saving' ? 'Saving…' : saveState === 'error' ? 'Could not save · export a copy' : 'Saved on this device';

  return <div className="sheet-editor">
    <header className="sheet-topbar">
      <button className="sheet-icon-button sheet-back" onClick={() => { commit(); onBack(); }} title="Back to workspace"><ArrowLeft size={20} /></button>
      <div className="sheet-app-icon"><FileSpreadsheet size={23} /></div>
      <div className="sheet-file-info"><input className="sheet-file-title" aria-label="Spreadsheet title" value={title} onChange={event => setTitle(event.target.value)} onBlur={() => { const name = title.trim() || 'Untitled spreadsheet'; setTitle(name); onRename(name); }} onKeyDown={event => { if (event.key === 'Enter') event.currentTarget.blur(); }} /><span className={saveState === 'error' ? 'sheet-save-error' : ''}><Check size={12} /> {status}</span></div>
      <span className="sheet-private-badge">SPREADSHEET</span>
      {onOpenVersions && <button className="sheet-icon-button" title="Version history" onClick={() => { commit(); onOpenVersions(); }}><History size={18} /></button>}
      <div className="sheet-export-wrap"><button className="sheet-export-button" onClick={() => setExportOpen(!exportOpen)}><Download size={16} /> Export <ChevronDown size={14} /></button>{exportOpen && <><button className="sheet-menu-scrim" aria-label="Close export menu" onClick={() => setExportOpen(false)} /><div className="sheet-export-menu"><button onClick={() => void exportSheet('xlsx')}><FileSpreadsheet size={16} /><span>Excel workbook<small>.xlsx · all sheets, styles & formulas</small></span></button><button onClick={() => void exportSheet('csv')}><Download size={16} /><span>Comma-separated values<small>.csv · current sheet, calculated values</small></span></button></div></>}</div>
    </header>
    <div className="sheet-ribbon-tabs"><div className="sheet-history-tools"><button className="sheet-icon-button" disabled={!undoStack.length} title="Undo (Ctrl+Z)" onClick={undo}><Undo2 size={16} /></button><button className="sheet-icon-button" disabled={!redoStack.length} title="Redo (Ctrl+Shift+Z)" onClick={redo}><Redo2 size={16} /></button></div>{(['home', 'data', 'view'] as const).map(tab => <button key={tab} className={ribbon === tab ? 'is-active' : ''} onClick={() => setRibbon(tab)}>{tab[0].toUpperCase() + tab.slice(1)}</button>)}<button className={`sheet-find-toggle ${findOpen ? 'is-active' : ''}`} title="Find and replace (Ctrl+F)" onClick={() => { setFindOpen(!findOpen); requestAnimationFrame(() => findRef.current?.focus()); }}><Search size={15} /><span>Find & replace</span></button><button title="Quick guide" className={helpOpen ? 'is-active' : ''} onClick={() => setHelpOpen(!helpOpen)}><Keyboard size={17} /></button></div>
    <div className="sheet-toolbar sheet-ribbon" onMouseDown={event => { if ((event.target as HTMLElement).closest('button')) event.preventDefault(); }}>
      {ribbon === 'home' && <>
        <div className="sheet-tool-group"><button className={`sheet-icon-button ${activeStyle.bold ? 'is-active' : ''}`} aria-pressed={!!activeStyle.bold} title="Bold (Ctrl+B)" onClick={() => applyStyle({ bold: !activeStyle.bold })}><Bold size={16} /></button><button className={`sheet-icon-button ${activeStyle.italic ? 'is-active' : ''}`} aria-pressed={!!activeStyle.italic} title="Italic (Ctrl+I)" onClick={() => applyStyle({ italic: !activeStyle.italic })}><Italic size={16} /></button><button className={`sheet-icon-button ${activeStyle.underline ? 'is-active' : ''}`} aria-pressed={!!activeStyle.underline} title="Underline (Ctrl+U)" onClick={() => applyStyle({ underline: !activeStyle.underline })}><Underline size={16} /></button><label className="sheet-color-control" title="Text color"><span style={{ borderBottomColor: activeStyle.color || '#354d40' }}>A</span><input type="color" aria-label="Text color" value={activeStyle.color || '#354d40'} onChange={event => applyStyle({ color: event.target.value })} /></label><label className="sheet-color-control" title="Fill color"><PaintBucket size={16} style={{ color: activeStyle.background || '#607968' }} /><input type="color" aria-label="Fill color" value={activeStyle.background || '#ffffff'} onChange={event => applyStyle({ background: event.target.value })} /></label></div>
        <div className="sheet-tool-group">{([{ align: 'left', Icon: AlignLeft }, { align: 'center', Icon: AlignCenter }, { align: 'right', Icon: AlignRight }] as const).map(({ align, Icon }) => <button key={align} className={`sheet-icon-button ${activeStyle.align === align ? 'is-active' : ''}`} aria-pressed={activeStyle.align === align} title={`Align ${align}`} onClick={() => applyStyle({ align })}><Icon size={16} /></button>)}</div>
        <div className="sheet-tool-group"><select aria-label="Number format" value={activeStyle.numberFormat || 'general'} onChange={event => applyStyle({ numberFormat: event.target.value as CellStyle['numberFormat'] })}><option value="general">General</option><option value="number">Number</option><option value="currency">Currency (€)</option><option value="percentage">Percentage</option><option value="date">Date</option></select><button title="Decrease decimal places" onClick={() => applyStyle({ decimals: Math.max(0, (activeStyle.decimals ?? (activeStyle.numberFormat === 'percentage' ? 0 : 2)) - 1), numberFormat: !activeStyle.numberFormat || activeStyle.numberFormat === 'general' ? 'number' : activeStyle.numberFormat })}>.0</button><button title="Increase decimal places" onClick={() => applyStyle({ decimals: Math.min(8, (activeStyle.decimals ?? (activeStyle.numberFormat === 'percentage' ? 0 : 2)) + 1), numberFormat: !activeStyle.numberFormat || activeStyle.numberFormat === 'general' ? 'number' : activeStyle.numberFormat })}>.00</button></div>
        <div className="sheet-tool-group"><button onClick={insertSum} title="Insert SUM formula"><Sigma size={17} /> Auto sum</button><button title="Fill down (Ctrl+D)" onClick={() => fill('down')}><ArrowDown size={15} /><span>Fill down</span></button><button title="Fill right (Ctrl+R)" onClick={() => fill('right')}><ArrowRight size={15} /></button></div><button className="sheet-icon-button" onClick={() => clear()} title="Clear selected cell contents"><Trash2 size={15} /></button><button className="sheet-text-tool" onClick={() => clear(true)} title="Clear selected cell formatting">Clear format</button>
      </>}
      {ribbon === 'data' && <><div className="sheet-tool-group"><button title={`Sort column ${columnName(active.col)} ascending`} onClick={() => sortRows(1)}><ArrowDownAZ size={18} /> Sort A–Z</button><button title={`Sort column ${columnName(active.col)} descending`} onClick={() => sortRows(-1)}><ArrowUpAZ size={18} /> Sort Z–A</button></div><label className="sheet-check-control"><input type="checkbox" checked={protectHeader} onChange={event => setProtectHeader(event.target.checked)} /> First row is a header</label><div className="sheet-filter-control"><label htmlFor="sheet-filter">Filter {columnName(filter?.col ?? active.col)}</label><input id="sheet-filter" aria-label="Filter selected column" placeholder="Contains…" value={filter?.text || ''} onChange={event => setFilter({ col: filter?.col ?? active.col, text: event.target.value })} /><button className="sheet-icon-button" title="Clear filter" disabled={!filter} onClick={() => setFilter(null)}><X size={14} /></button></div><span className="sheet-toolbar-note">Sort moves entire rows</span></>}
      {ribbon === 'view' && <><div className="sheet-tool-group"><button className={sheet.freezeRows ? 'is-active' : ''} aria-pressed={!!sheet.freezeRows} onClick={() => { commit(); const current = currentSheet(); updateSheet({ ...current, freezeRows: current.freezeRows ? 0 : 1 }); }}><Snowflake size={16} /> Freeze first row</button><button className={showFormulas ? 'is-active' : ''} aria-pressed={showFormulas} onClick={() => setShowFormulas(!showFormulas)}>ƒx Show formulas</button></div><label className="sheet-check-control"><input type="checkbox" checked={showGridlines} onChange={event => setShowGridlines(event.target.checked)} /> Gridlines</label><div className="sheet-tool-group"><button disabled={rows >= 2000} onClick={() => setRowCount(Math.min(2000, rows + 10))}><Plus size={15} /> Add 10 rows</button><button disabled={cols >= 100} onClick={() => setColCount(Math.min(100, cols + 5))}><Plus size={15} /> Add 5 columns</button></div><button className="sheet-text-tool" onClick={() => autoFit(active.col)}>Fit column {columnName(active.col)}</button></>}
    </div>
    {findOpen && <div className="sheet-find-panel"><Search size={16} /><input ref={findRef} aria-label="Find in sheet" placeholder="Find in sheet" value={search} onChange={event => setSearch(event.target.value)} onKeyDown={event => { if (event.key === 'Enter') findNext(); }} /><span>{matches.length} matches</span><button disabled={!matches.length} onClick={findNext}>Next</button><input aria-label="Replace with" placeholder="Replace with" value={replacement} onChange={event => setReplacement(event.target.value)} /><button disabled={!search} onClick={() => replace(false)}>Replace cell</button><button disabled={!search} onClick={() => replace(true)}>Replace all</button><button className="sheet-icon-button" title="Close find and replace" onClick={() => { setFindOpen(false); setSearch(''); }}><X size={15} /></button><small>Replaces text in stored cell contents, including formulas.</small></div>}
    {helpOpen && <div className="sheet-help"><span><b>Make yourself at home.</b> Double-click or press F2 to edit. Enter moves down, Tab moves right. Shift-click selects a range. Click a row or column heading to select it. Drag a column edge to resize; double-click to fit.</span><span>Use <b>Fill down / right</b> to copy formulas with relative references. Keep references fixed with <code>$A$1</code>. Try <code>=IF(B2&gt;100,"Over","OK")</code> <code>=SUM(B2:B8)</code> or <code>='Sheet 2'!A1</code>. Right-click a sheet tab to rename it.</span><button className="sheet-icon-button" title="Close guide" onClick={() => setHelpOpen(false)}><X size={16} /></button></div>}
    {filter?.text && <div className="sheet-filter-banner">Showing {visibleRows.length} of {rows} rows · Column {columnName(filter.col)} contains “{filter.text}”<button onClick={() => setFilter(null)}>Clear filter <X size={12} /></button></div>}
    <div className="sheet-formula-bar"><div className="sheet-address">{rangeLabel}</div><span className="sheet-fx">ƒx</span><input ref={formulaRef} aria-label="Formula bar" placeholder="Enter a value or formula" value={editing ? draft : sheet.cells[activeKey] || ''} onFocus={() => { if (!editing) beginEdit(); }} onChange={event => changeDraft(event.target.value)} onBlur={commit} onKeyDown={event => { if (event.key === 'Enter') { event.preventDefault(); commit(); focus({ ...active, row: moveRow(active.row, 1) }); } if (event.key === 'Escape') { editRef.current = null; setEditing(false); inputs.current.get(activeKey)?.focus(); } }} /></div>
    <div className={`sheet-grid-scroll ${showGridlines ? '' : 'sheet-no-gridlines'}`} onPaste={paste} onCopy={copy}>
      <table className="sheet-grid" style={{ zoom: zoom / 100 } as CSSProperties}><colgroup><col className="sheet-row-number-col" />{Array.from({ length: cols }, (_, col) => <col key={col} style={{ width: resizing?.col === col ? resizing.width : sheet.columnWidths?.[columnName(col)] || 122 }} />)}</colgroup><thead><tr><th className="sheet-corner" onClick={() => { commit(); setAnchor({ row: 0, col: 0 }); setActive({ row: Math.min(lastUsed.row, 1999), col: Math.min(lastUsed.col, 99) }); }} title="Select used cells"><span /></th>{Array.from({ length: cols }, (_, col) => <th key={col} className={col >= range.firstCol && col <= range.lastCol ? 'sheet-selected-heading' : ''} onClick={() => { commit(); setAnchor({ row: 0, col }); setActive({ row: rows - 1, col }); }}><span>{columnName(col)}</span>{filter?.col === col && filter.text && <span className="sheet-column-filter">•</span>}<span className="sheet-column-resize" role="separator" aria-label={`Resize column ${columnName(col)}`} aria-orientation="vertical" onPointerDown={event => startResize(event, col)} onClick={event => event.stopPropagation()} onDoubleClick={event => { event.stopPropagation(); autoFit(col); }} /></th>)}</tr></thead><tbody>{visibleRows.map(row => <tr key={row} className={sheet.freezeRows && row === 0 ? 'sheet-frozen-row' : ''}><th className={row >= range.firstRow && row <= range.lastRow ? 'sheet-selected-heading' : ''} onClick={() => { commit(); setAnchor({ row, col: 0 }); setActive({ row, col: cols - 1 }); }}>{row + 1}</th>{Array.from({ length: cols }, (_, col) => { const key = cellKey({ row, col }); const value = evaluate(key); const isActive = activeKey === key; const isSelected = row >= range.firstRow && row <= range.lastRow && col >= range.firstCol && col <= range.lastCol; const style = sheet.styles?.[key] || {}; return <td key={key} style={{ backgroundColor: style.background }} className={`${isActive ? 'sheet-active-cell ' : ''}${isSelected ? 'sheet-selected-cell ' : ''}${matchSet.has(key) ? 'sheet-search-match ' : ''}${String(value).startsWith('#') ? 'sheet-error-cell' : ''}`}><input ref={element => { if (element) inputs.current.set(key, element); else inputs.current.delete(key); }} aria-label={`Cell ${key}`} title={sheet.cells[key]?.startsWith('=') ? `${sheet.cells[key]} → ${value}` : undefined} tabIndex={isActive ? 0 : -1} readOnly={!isActive || !editing} value={isActive && editing ? draft : showFormulas && sheet.cells[key]?.startsWith('=') ? sheet.cells[key] : formatValue(value, style)} style={{ fontWeight: style.bold ? 700 : undefined, fontStyle: style.italic ? 'italic' : undefined, textDecoration: style.underline ? 'underline' : undefined, color: style.color, textAlign: style.align }} className={typeof value === 'number' && !(isActive && editing) ? 'sheet-number' : ''} onMouseDown={event => { if (event.shiftKey) { event.preventDefault(); focus({ row, col }, true); } else if (!isActive) select({ row, col }); }} onFocus={() => { if (!isActive) select({ row, col }); }} onDoubleClick={() => beginEdit(sheet.cells[key] || '')} onChange={event => changeDraft(event.target.value)} onBlur={commit} onKeyDown={event => keyboard(event, { row, col })} />{isActive && <span className="sheet-cell-handle" />}</td>; })}</tr>)}</tbody></table>
      {!visibleRows.length && <div className="sheet-empty-filter">No rows match this filter.<button onClick={() => setFilter(null)}>Clear filter</button></div>}
    </div>
    <footer className="sheet-footer"><div className="sheet-tabs" role="tablist" aria-label="Workbook sheets">{workbook.sheets.map(item => <button key={item.id} role="tab" aria-selected={item.id === sheet.id} className={`sheet-tab ${item.id === sheet.id ? 'is-active' : ''}`} onClick={() => switchSheet(item.id)} onDoubleClick={() => { setDialogError(''); setSheetDialog({ mode: 'rename', id: item.id, value: item.name }); }} onContextMenu={event => { event.preventDefault(); setDialogError(''); setSheetDialog({ mode: 'rename', id: item.id, value: item.name }); }}><FileSpreadsheet size={13} /><span>{item.name}</span></button>)}</div><button className="sheet-icon-button" title="Add sheet" onClick={() => addSheet()}><Plus size={17} /></button><button className="sheet-icon-button" title="Duplicate current sheet" onClick={() => addSheet(true)}><Copy size={15} /></button><button className="sheet-icon-button" title="Rename current sheet" onClick={() => { setDialogError(''); setSheetDialog({ mode: 'rename', id: sheet.id, value: sheet.name }); }}><span className="sheet-rename-icon">A</span></button><button className="sheet-icon-button" title="Delete current sheet" disabled={workbook.sheets.length === 1} onClick={() => { setDialogError(''); setSheetDialog({ mode: 'delete', id: sheet.id, value: sheet.name }); }}><Trash2 size={14} /></button><div className="sheet-selection-info">{selectedKeys.length > 1 ? <><span>Count <b>{selectedKeys.filter(key => sheet.cells[key]).length}</b></span><span>Sum <b>{numericValues.reduce((sum, value) => sum + value, 0).toLocaleString(undefined, { maximumFractionDigits: 6 })}</b></span>{numericValues.length > 0 && <span className="sheet-average">Average <b>{(numericValues.reduce((sum, value) => sum + value, 0) / numericValues.length).toLocaleString(undefined, { maximumFractionDigits: 4 })}</b></span>}</> : <span>{Object.keys(sheet.cells).length} filled cells</span>}<select className="sheet-zoom-select" aria-label="Spreadsheet zoom" value={zoom} onChange={event => setZoom(Number(event.target.value))}>{[75, 90, 100, 110, 125, 150].map(value => <option key={value} value={value}>{value}%</option>)}</select></div></footer>
    {sortConfirmation && <div className="sheet-modal-backdrop"><section className="sheet-modal" role="dialog" aria-modal="true" aria-labelledby="sort-dialog-title"><h2 id="sort-dialog-title">Sort rows with formulas?</h2><p>All data rows will be reordered by column {columnName(active.col)}{protectHeader ? ', keeping the header in place' : ''}. Formulas move with their rows and relative references adjust.</p><p>References from other cells or sheets keep their cell addresses, so their results may change. Summary rows are included. You can undo this sort.</p><div className="sheet-modal-actions"><button onClick={() => setSortConfirmation(null)}>Cancel</button><button className="sheet-primary" onClick={() => sortRows(sortConfirmation, true)}>Sort rows</button></div></section></div>}
    {sheetDialog && <div className="sheet-modal-backdrop" onMouseDown={event => { if (event.target === event.currentTarget) setSheetDialog(null); }}><form className="sheet-modal" role="dialog" aria-modal="true" aria-labelledby="sheet-dialog-title" onSubmit={event => { event.preventDefault(); submitSheetDialog(); }}><h2 id="sheet-dialog-title">{sheetDialog.mode === 'rename' ? 'Rename sheet' : 'Delete sheet?'}</h2>{sheetDialog.mode === 'rename' ? <><p>Give this worksheet a useful name.</p><input autoFocus aria-label="New sheet name" maxLength={31} value={sheetDialog.value} onFocus={event => event.currentTarget.select()} onChange={event => setSheetDialog({ ...sheetDialog, value: event.target.value })} onKeyDown={event => { if (event.key === 'Escape') setSheetDialog(null); }} />{dialogError && <p className="sheet-dialog-error" role="alert">{dialogError}</p>}</> : <p>“{sheetDialog.value}” and its cells will be removed. Formulas in other sheets that reference it will show an error. You can undo this change.</p>}<div className="sheet-modal-actions"><button type="button" onClick={() => setSheetDialog(null)}>Cancel</button><button className={sheetDialog.mode === 'delete' ? 'sheet-danger' : 'sheet-primary'} type="submit">{sheetDialog.mode === 'rename' ? 'Rename' : 'Delete sheet'}</button></div></form></div>}
  </div>;
}
