import { useEffect, useMemo, useRef, useState } from 'react';
import type { ClipboardEvent, KeyboardEvent } from 'react';
import { ArrowLeft, ChevronDown, Check, Download, FileSpreadsheet, Undo2, Redo2, Plus, Search, Sigma, Trash2, X, Keyboard } from 'lucide-react';
import * as XLSX from 'xlsx';
import type { EditorProps, SheetContent } from '../types';
import './spreadsheet-editor.css';

type Point = { row: number; col: number };
const columnName = (n: number): string => { let s = ''; for (n++; n; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
const cellKey = ({ row, col }: Point) => `${columnName(col)}${row + 1}`;
const pointFor = (key: string): Point => { const match = key.replaceAll('$', '').match(/^([A-Z]+)(\d+)$/i); if (!match || Number(match[2]) < 1) throw Error('#REF!'); return { col: [...match[1].toUpperCase()].reduce((v, c) => v * 26 + c.charCodeAt(0) - 64, 0) - 1, row: Number(match[2]) - 1 }; };
const cleanSheet = (content: SheetContent): SheetContent => ({ cells: content?.cells || {}, name: content?.name || 'Sheet 1' });
const isStoredNumber = (value: string) => /^[+-]?(?:0|[1-9]\d*)(?:\.\d*)?(?:e[+-]?\d+)?$/i.test(value) && Number.isFinite(Number(value)) && value.replace(/\D/g, '').length <= 15;

/** Restricted recursive-descent spreadsheet evaluator. Never executes JavaScript. */
export function calculateCells(cells: Record<string, string>) {
  const cache = new Map<string, string | number>();
  const calculate = (key: string, ancestors = new Set<string>()): string | number => {
    if (cache.has(key)) return cache.get(key)!;
    if (ancestors.has(key)) throw Error('#CYCLE!');
    if (ancestors.size > 200) throw Error('#DEPTH!');
    const raw = String(cells[key] ?? '');
    if (!raw.startsWith('=')) return isStoredNumber(raw) ? Number(raw) : raw;
    const visiting = new Set(ancestors); visiting.add(key);
    const source = raw.slice(1).replaceAll('$', '').toUpperCase();
    const tokens = source.match(/(?:\d+\.?\d*|\.\d+)(?:E[+-]?\d+)?|[A-Z_][A-Z_0-9]*|[+\-*/^%(),:]/g) || [];
    if (tokens.join('') !== source.replace(/\s/g, '') || tokens.length > 1000) throw Error('#VALUE!');
    let position = 0;
    const numberFor = (ref: string) => { const value = calculate(ref, visiting); if (typeof value === 'string' && value.startsWith('#')) throw Error(value); return typeof value === 'number' ? value : 0; };
    const primary = (): number => {
      const token = tokens[position++];
      if (token === '(') { const result = expression(); if (tokens[position++] !== ')') throw Error('#VALUE!'); return result; }
      if (token && /^(?:\d|\.)/.test(token)) return Number(token);
      if (token && /^[A-Z]+\d+$/.test(token)) { pointFor(token); return numberFor(token); }
      if (token && ['SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'ABS', 'ROUND'].includes(token)) {
        if (tokens[position++] !== '(') throw Error('#VALUE!');
        const values: number[] = [];
        while (tokens[position] && tokens[position] !== ')') {
          if (/^[A-Z]+\d+$/.test(tokens[position]) && tokens[position + 1] === ':') {
            const start = pointFor(tokens[position]); const end = pointFor(tokens[position + 2] || ''); position += 3;
            if ((Math.abs(end.row - start.row) + 1) * (Math.abs(end.col - start.col) + 1) > 100000 || Math.max(start.row, end.row) > 100000 || Math.max(start.col, end.col) > 1000) throw Error('#REF!');
            for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row++) for (let col = Math.min(start.col, end.col); col <= Math.max(start.col, end.col); col++) {
              const value = calculate(cellKey({ row, col }), visiting);
              if (typeof value === 'string' && value.startsWith('#')) throw Error(value);
              if (typeof value === 'number') values.push(value);
            }
          } else if (/^[A-Z]+\d+$/.test(tokens[position]) && [',', ')'].includes(tokens[position + 1])) {
            const reference = tokens[position++]; pointFor(reference); const value = calculate(reference, visiting);
            if (typeof value === 'string' && value.startsWith('#')) throw Error(value);
            if (typeof value === 'number') values.push(value);
          } else values.push(expression());
          if (tokens[position] !== ',') break;
          position++;
        }
        if (tokens[position++] !== ')') throw Error('#VALUE!');
        if (token === 'SUM') return values.reduce((a, b) => a + b, 0);
        if (token === 'COUNT') return values.length;
        if (token === 'AVERAGE') { if (!values.length) throw Error('#DIV/0!'); return values.reduce((a, b) => a + b, 0) / values.length; }
        if (token === 'MIN') return values.length ? Math.min(...values) : 0;
        if (token === 'MAX') return values.length ? Math.max(...values) : 0;
        if (token === 'ABS') return Math.abs(values[0] || 0);
        const factor = 10 ** (values[1] || 0); return Math.round((values[0] || 0) * factor) / factor;
      }
      throw Error('#NAME?');
    };
    const power = (): number => { let result = primary(); if (tokens[position] === '^') { position++; result **= unary(); } while (tokens[position] === '%') { position++; result /= 100; } return result; };
    const unary = (): number => { if (tokens[position] === '+') { position++; return unary(); } if (tokens[position] === '-') { position++; return -unary(); } return power(); };
    const product = (): number => { let result = unary(); while (tokens[position] === '*' || tokens[position] === '/') { const operator = tokens[position++]; const rhs = unary(); if (operator === '/' && rhs === 0) throw Error('#DIV/0!'); result = operator === '*' ? result * rhs : result / rhs; } return result; };
    const expression = (): number => { let result = product(); while (tokens[position] === '+' || tokens[position] === '-') { const operator = tokens[position++]; const rhs = product(); result = operator === '+' ? result + rhs : result - rhs; } return result; };
    const value = expression();
    if (position !== tokens.length || !Number.isFinite(value)) throw Error('#VALUE!');
    const result = Math.round(value * 1e10) / 1e10;
    cache.set(key, result); return result;
  };
  return (key: string): string | number => { try { return calculate(key); } catch (error) { return error instanceof Error ? error.message : '#VALUE!'; } };
}

export default function SpreadsheetEditor({ file, onChange, onRename, onBack, onNotify }: EditorProps) {
  const [sheet, setSheet] = useState(() => cleanSheet(file.content));
  const sheetRef = useRef(sheet);
  const [title, setTitle] = useState(file.name);
  const [active, setActive] = useState<Point>({ row: 0, col: 0 });
  const [anchor, setAnchor] = useState<Point | null>(null);
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState('');
  const editRef = useRef<{ key: string; value: string } | null>(null);
  const [rowCount, setRowCount] = useState(40);
  const [exportOpen, setExportOpen] = useState(false);
  const [helpOpen, setHelpOpen] = useState(false);
  const [search, setSearch] = useState('');
  const [undoStack, setUndoStack] = useState<SheetContent[]>([]);
  const [redoStack, setRedoStack] = useState<SheetContent[]>([]);
  const inputs = useRef(new Map<string, HTMLInputElement>());
  const formulaRef = useRef<HTMLInputElement>(null);
  const evaluate = useMemo(() => calculateCells(sheet.cells), [sheet.cells]);
  const activeKey = cellKey(active);
  const lastUsed = useMemo(() => Object.keys(sheet.cells).reduce((acc, key) => { try { const point = pointFor(key); return { row: Math.max(acc.row, point.row), col: Math.max(acc.col, point.col) }; } catch { return acc; } }, { row: 39, col: 11 }), [sheet.cells]);
  const rows = Math.min(2000, Math.max(rowCount, lastUsed.row + 1));
  const cols = Math.min(100, Math.max(12, lastUsed.col + 1));
  const range = { firstRow: Math.min(active.row, anchor?.row ?? active.row), lastRow: Math.max(active.row, anchor?.row ?? active.row), firstCol: Math.min(active.col, anchor?.col ?? active.col), lastCol: Math.max(active.col, anchor?.col ?? active.col) };
  const rangeLabel = anchor ? `${cellKey({ row: range.firstRow, col: range.firstCol })}:${cellKey({ row: range.lastRow, col: range.lastCol })}` : activeKey;
  const selectedKeys: string[] = [];
  for (let row = range.firstRow; row <= range.lastRow; row++) for (let col = range.firstCol; col <= range.lastCol; col++) selectedKeys.push(cellKey({ row, col }));
  const numericValues = selectedKeys.map(evaluate).filter((v): v is number => typeof v === 'number');
  const matches = search.trim() ? Object.keys(sheet.cells).filter(key => String(evaluate(key)).toLowerCase().includes(search.toLowerCase())) : [];

  useEffect(() => { const next = cleanSheet(file.content); sheetRef.current = next; setSheet(next); setTitle(file.name); setActive({ row: 0, col: 0 }); setAnchor(null); setUndoStack([]); setRedoStack([]); editRef.current = null; setEditing(false); }, [file.id]);

  const updateSheet = (next: SheetContent, record = true) => {
    const previous = sheetRef.current;
    if (record) { setUndoStack(stack => [...stack.slice(-79), previous]); setRedoStack([]); }
    sheetRef.current = next; setSheet(next); onChange(next);
  };
  const commit = () => {
    const pending = editRef.current;
    if (!pending) return;
    editRef.current = null; setEditing(false);
    if ((sheetRef.current.cells[pending.key] || '') === pending.value) return;
    const cells = { ...sheetRef.current.cells };
    if (pending.value) cells[pending.key] = pending.value; else delete cells[pending.key];
    updateSheet({ ...sheetRef.current, cells });
  };
  const beginEdit = (value = sheetRef.current.cells[activeKey] || '') => { editRef.current = { key: activeKey, value }; setDraft(value); setEditing(true); };
  const changeDraft = (value: string) => { editRef.current = { key: activeKey, value }; setDraft(value); setEditing(true); };
  const select = (point: Point, extend = false) => { commit(); if (extend) setAnchor(previous => previous || active); else setAnchor(null); setActive(point); };
  const focus = (point: Point, extend = false) => { const next = { row: Math.max(0, Math.min(rows - 1, point.row)), col: Math.max(0, Math.min(cols - 1, point.col)) }; select(next, extend); requestAnimationFrame(() => inputs.current.get(cellKey(next))?.focus()); };
  const clear = () => { editRef.current = null; setEditing(false); const cells = { ...sheetRef.current.cells }; selectedKeys.forEach(key => delete cells[key]); updateSheet({ ...sheetRef.current, cells }); };
  const undo = () => { if (!undoStack.length) return; editRef.current = null; setEditing(false); const current = sheetRef.current; const previous = undoStack[undoStack.length - 1]; setRedoStack(stack => [...stack, current]); setUndoStack(stack => stack.slice(0, -1)); updateSheet(previous, false); };
  const redo = () => { if (!redoStack.length) return; editRef.current = null; setEditing(false); const current = sheetRef.current; const next = redoStack[redoStack.length - 1]; setUndoStack(stack => [...stack, current]); setRedoStack(stack => stack.slice(0, -1)); updateSheet(next, false); };
  const keyboard = (event: KeyboardEvent<HTMLInputElement>, point: Point) => {
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'z') { event.preventDefault(); event.shiftKey ? redo() : undo(); return; }
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'y') { event.preventDefault(); redo(); return; }
    if (event.key === 'Escape') { editRef.current = null; setEditing(false); setAnchor(null); return; }
    if (event.key === 'Enter') { event.preventDefault(); focus({ ...point, row: point.row + (event.shiftKey ? -1 : 1) }); return; }
    if (event.key === 'Tab') { event.preventDefault(); const col = point.col + (event.shiftKey ? -1 : 1); focus({ row: point.row + (col >= cols ? 1 : col < 0 ? -1 : 0), col: (col + cols) % cols }); return; }
    if (editing) return;
    if (event.key.startsWith('Arrow')) { event.preventDefault(); const delta = { ArrowDown: [1, 0], ArrowUp: [-1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[event.key]!; focus({ row: point.row + delta[0], col: point.col + delta[1] }, event.shiftKey); }
    else if (event.key === 'Delete' || event.key === 'Backspace') { event.preventDefault(); clear(); }
    else if (event.key === 'F2') { event.preventDefault(); beginEdit(); }
    else if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) { event.preventDefault(); beginEdit(event.key); }
  };
  const paste = (event: ClipboardEvent<HTMLDivElement>) => {
    const text = event.clipboardData.getData('text/plain');
    if (editing && !text.includes('\t') && !text.includes('\n')) return;
    event.preventDefault(); commit();
    const data = text.replace(/\r/g, '').replace(/\n$/, '').split('\n').map(row => row.split('\t'));
    const cells = { ...sheetRef.current.cells };
    data.slice(0, 2000 - active.row).forEach((values, row) => values.slice(0, 100 - active.col).forEach((value, col) => { const key = cellKey({ row: active.row + row, col: active.col + col }); if (value) cells[key] = value; else delete cells[key]; }));
    updateSheet({ ...sheetRef.current, cells }); setRowCount(count => Math.max(count, active.row + data.length));
  };
  const exportSheet = (format: 'xlsx' | 'csv') => {
    commit();
    try {
      const content = sheetRef.current;
      const currentEvaluate = calculateCells(content.cells);
      const workbook = XLSX.utils.book_new();
      const worksheet: XLSX.WorkSheet = {};
      let maxRow = 0; let maxCol = 0;
      Object.entries(content.cells).forEach(([key, raw]) => { const point = pointFor(key); maxRow = Math.max(maxRow, point.row); maxCol = Math.max(maxCol, point.col); const value = currentEvaluate(key); worksheet[key] = { t: typeof value === 'number' ? 'n' : 's', v: value, ...(raw.startsWith('=') && format === 'xlsx' ? { f: raw.slice(1) } : {}) }; });
      worksheet['!ref'] = `A1:${cellKey({ row: maxRow, col: maxCol })}`;
      worksheet['!cols'] = Array.from({ length: maxCol + 1 }, () => ({ wch: 20 }));
      XLSX.utils.book_append_sheet(workbook, worksheet, (content.name.replace(/[\\/?*\[\]:]/g, '') || 'Sheet 1').slice(0, 31));
      XLSX.writeFile(workbook, `${file.name.replace(/\.(xlsx|csv)$/i, '') || 'Spreadsheet'}.${format}`, { bookType: format });
      setExportOpen(false); onNotify(`${format.toUpperCase()} downloaded`);
    } catch { onNotify('The spreadsheet could not be exported. Please try again.'); }
  };
  const insertSum = () => { const formula = anchor ? `=SUM(${rangeLabel})` : active.row > 0 ? `=SUM(${columnName(active.col)}1:${columnName(active.col)}${active.row})` : '=SUM()'; if (anchor) { const target = { row: range.lastRow + 1, col: range.firstCol }; if (target.row >= rows) setRowCount(rows + 10); setActive(target); setAnchor(null); editRef.current = { key: cellKey(target), value: formula }; setDraft(formula); setEditing(true); } else beginEdit(formula); requestAnimationFrame(() => formulaRef.current?.focus()); };
  const findNext = () => { if (!matches.length) return; const next = matches[(matches.indexOf(activeKey) + 1) % matches.length]; focus(pointFor(next)); };

  return <div className="sheet-editor">
    <header className="sheet-topbar">
      <button className="sheet-icon-button sheet-back" onClick={() => { commit(); onBack(); }} title="Back to workspace"><ArrowLeft size={20} /></button>
      <div className="sheet-app-icon"><FileSpreadsheet size={23} /></div>
      <div className="sheet-file-info"><input className="sheet-file-title" aria-label="Spreadsheet title" value={title} onChange={event => setTitle(event.target.value)} onBlur={() => { const name = title.trim() || 'Untitled spreadsheet'; setTitle(name); onRename(name); }} onKeyDown={event => { if (event.key === 'Enter') event.currentTarget.blur(); }} /><span><Check size={12} /> Saved on this device</span></div>
      <span className="sheet-private-badge">SPREADSHEET</span>
      <div className="sheet-export-wrap"><button className="sheet-export-button" onClick={() => setExportOpen(!exportOpen)}><Download size={16} /> Export <ChevronDown size={14} /></button>{exportOpen && <><button className="sheet-menu-scrim" aria-label="Close export menu" onClick={() => setExportOpen(false)} /><div className="sheet-export-menu"><button onClick={() => exportSheet('xlsx')}><FileSpreadsheet size={16} /><span>Excel workbook<small>.xlsx · formulas included</small></span></button><button onClick={() => exportSheet('csv')}><Download size={16} /><span>Comma-separated values<small>.csv · calculated values</small></span></button></div></>}</div>
    </header>
    <div className="sheet-toolbar">
      <div className="sheet-tool-group"><button className="sheet-icon-button" disabled={!undoStack.length} title="Undo (Ctrl+Z)" onClick={undo}><Undo2 size={17} /></button><button className="sheet-icon-button" disabled={!redoStack.length} title="Redo (Ctrl+Shift+Z)" onClick={redo}><Redo2 size={17} /></button></div>
      <div className="sheet-tool-group"><button onClick={insertSum}><Sigma size={17} /> Auto sum</button><button disabled={rows >= 2000} onClick={() => setRowCount(Math.min(2000, rows + 10))}><Plus size={16} /> Add 10 rows</button><button onClick={clear} title="Clear selected cells"><Trash2 size={15} /><span className="sheet-clear-label">Clear</span></button></div>
      <button className={`sheet-help-button ${helpOpen ? 'is-active' : ''}`} onClick={() => setHelpOpen(!helpOpen)}><Keyboard size={16} /><span>Quick guide</span></button>
      <div className="sheet-find"><Search size={15} /><input value={search} placeholder="Find in sheet" aria-label="Find in sheet" onChange={event => setSearch(event.target.value)} onKeyDown={event => { if (event.key === 'Enter') findNext(); }} />{search && <><span>{matches.length}</span><button className="sheet-icon-button" title="Find next match" onClick={findNext}><ChevronDown size={14} /></button><button className="sheet-icon-button" title="Clear search" onClick={() => setSearch('')}><X size={13} /></button></>}</div>
    </div>
    {helpOpen && <div className="sheet-help"><span><b>Make yourself at home.</b> Double-click a cell to edit. Enter moves down, Tab moves right. Shift-click selects a range. Paste a table from Excel.</span><span>Formulas: <code>=SUM(B2:B8)</code> <code>=AVERAGE(C2:C8)</code> <code>=A1*1.2</code> · Also MIN, MAX, COUNT, ABS, ROUND.</span><button className="sheet-icon-button" title="Close guide" onClick={() => setHelpOpen(false)}><X size={16} /></button></div>}
    <div className="sheet-formula-bar"><div className="sheet-address">{rangeLabel}</div><span className="sheet-fx">ƒx</span><input ref={formulaRef} aria-label="Formula bar" placeholder="Enter a value or formula" value={editing ? draft : sheet.cells[activeKey] || ''} onFocus={() => { if (!editing) beginEdit(); }} onChange={event => changeDraft(event.target.value)} onBlur={commit} onKeyDown={event => { if (event.key === 'Enter') { event.preventDefault(); commit(); focus({ ...active, row: active.row + 1 }); } if (event.key === 'Escape') { editRef.current = null; setEditing(false); inputs.current.get(activeKey)?.focus(); } }} /></div>
    <div className="sheet-grid-scroll" onPaste={paste} onCopy={event => { if (editing) return; const lines: string[] = []; for (let row = range.firstRow; row <= range.lastRow; row++) { const values: string[] = []; for (let col = range.firstCol; col <= range.lastCol; col++) values.push(String(evaluate(cellKey({ row, col })))); lines.push(values.join('\t')); } event.clipboardData.setData('text/plain', lines.join('\n')); event.preventDefault(); }}>
      <table className="sheet-grid"><colgroup><col className="sheet-row-number-col" />{Array.from({ length: cols }, (_, col) => <col key={col} />)}</colgroup><thead><tr><th className="sheet-corner"><span /></th>{Array.from({ length: cols }, (_, col) => <th key={col} className={col >= range.firstCol && col <= range.lastCol ? 'sheet-selected-heading' : ''}>{columnName(col)}</th>)}</tr></thead><tbody>{Array.from({ length: rows }, (_, row) => <tr key={row}><th className={row >= range.firstRow && row <= range.lastRow ? 'sheet-selected-heading' : ''}>{row + 1}</th>{Array.from({ length: cols }, (_, col) => { const key = cellKey({ row, col }); const value = evaluate(key); const isActive = activeKey === key; const isSelected = row >= range.firstRow && row <= range.lastRow && col >= range.firstCol && col <= range.lastCol; return <td key={key} className={`${isActive ? 'sheet-active-cell ' : ''}${isSelected ? 'sheet-selected-cell ' : ''}${matches.includes(key) ? 'sheet-search-match ' : ''}${String(value).startsWith('#') ? 'sheet-error-cell' : ''}`}><input ref={element => { if (element) inputs.current.set(key, element); else inputs.current.delete(key); }} aria-label={`Cell ${key}`} title={sheet.cells[key]?.startsWith('=') ? `${sheet.cells[key]} → ${value}` : undefined} tabIndex={isActive ? 0 : -1} readOnly={!isActive || !editing} value={isActive && editing ? draft : value} className={typeof value === 'number' && !(isActive && editing) ? 'sheet-number' : ''} onMouseDown={event => { if (event.shiftKey) { event.preventDefault(); focus({ row, col }, true); } else if (!isActive) select({ row, col }); }} onFocus={() => { if (!isActive) select({ row, col }); }} onDoubleClick={() => beginEdit(sheet.cells[key] || '')} onChange={event => changeDraft(event.target.value)} onBlur={commit} onKeyDown={event => keyboard(event, { row, col })} />{isActive && <span className="sheet-cell-handle" />}</td>; })}</tr>)}</tbody></table>
    </div>
    <footer className="sheet-footer"><div className="sheet-tab"><FileSpreadsheet size={14} /><input aria-label="Sheet name" value={sheet.name} onChange={event => updateSheet({ ...sheetRef.current, name: event.target.value })} onBlur={() => { if (!sheet.name.trim()) updateSheet({ ...sheetRef.current, name: 'Sheet 1' }); }} /></div><span className="sheet-ready"><span />Ready</span><div className="sheet-selection-info">{selectedKeys.length > 1 ? <><span>Count <b>{selectedKeys.filter(key => sheet.cells[key]).length}</b></span><span>Sum <b>{numericValues.reduce((sum, value) => sum + value, 0).toLocaleString(undefined, { maximumFractionDigits: 6 })}</b></span></> : <span>{Object.keys(sheet.cells).length} filled cells</span>}<span className="sheet-zoom">100%</span></div></footer>
  </div>;
}
