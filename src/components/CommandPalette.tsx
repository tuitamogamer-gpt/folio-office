import { useEffect, useId, useMemo, useRef, useState } from 'react';
import { ArrowDownToLine, ArrowUpRight, FileText, History, LayoutGrid, Plus, Presentation, Search, Settings2, Table2, Upload, X } from 'lucide-react';
import type { FileKind, OfficeFile } from '../types';
import './workspace-productivity.css';

interface Props {
  open: boolean;
  files: OfficeFile[];
  canOpenHistory: boolean;
  onClose: () => void;
  onOpenFile: (file: OfficeFile) => void;
  onCreate: (kind: FileKind) => void;
  onUpload: () => void;
  onTemplates: () => void;
  onSettings: () => void;
  onBackup: () => void;
  onHistory: () => void;
}

export default function CommandPalette(props: Props) {
  const [query, setQuery] = useState('');
  const [active, setActive] = useState(0);
  const input = useRef<HTMLInputElement>(null);
  const close = useRef<HTMLButtonElement>(null);
  const list = useRef<HTMLDivElement>(null);
  const id = useId();
  const actions = [
    { id: 'document', label: 'New document', detail: 'Start writing', icon: FileText, run: () => props.onCreate('document') },
    { id: 'spreadsheet', label: 'New spreadsheet', detail: 'Make room for your numbers', icon: Table2, run: () => props.onCreate('spreadsheet') },
    { id: 'presentation', label: 'New presentation', detail: 'Bring an idea to life', icon: Presentation, run: () => props.onCreate('presentation') },
    { id: 'upload', label: 'Import files', detail: 'Word, Excel, PowerPoint and more', icon: Upload, run: props.onUpload },
    { id: 'templates', label: 'Browse templates', detail: 'A head start for your next project', icon: LayoutGrid, run: props.onTemplates },
    { id: 'backup', label: 'Back up workspace', detail: 'Download every file and saved version', icon: ArrowDownToLine, run: props.onBackup },
    ...(props.canOpenHistory ? [{ id: 'history', label: 'Version history', detail: 'Save or restore a draft of this file', icon: History, run: props.onHistory }] : []),
    { id: 'settings', label: 'Workspace settings', detail: 'Your profile and backups', icon: Settings2, run: props.onSettings },
  ];
  const matches = (text: string) => query.trim().toLocaleLowerCase().split(/\s+/).every(word => text.toLocaleLowerCase().includes(word));
  const recent = useMemo(() => props.files.filter(file => !file.trashed).sort((a, b) => b.updatedAt - a.updatedAt), [props.files]);
  const items = [
    ...recent.filter(file => matches(`${file.name} ${file.kind}`)).slice(0, query.trim() ? 30 : 5).map(file => ({
      id: file.id, label: file.name, detail: file.kind === 'document' ? 'Document' : file.kind === 'spreadsheet' ? 'Spreadsheet' : 'Presentation',
      icon: file.kind === 'document' ? FileText : file.kind === 'spreadsheet' ? Table2 : Presentation,
      run: () => props.onOpenFile(file), group: query.trim() ? 'Files' : 'Recent files',
    })),
    ...actions.filter(action => matches(`${action.label} ${action.detail}`)).map(action => ({ ...action, group: 'Actions' })),
  ];
  const selected = Math.min(active, Math.max(0, items.length - 1));
  useEffect(() => { setActive(0); }, [query]);
  useEffect(() => {
    if (!props.open) return;
    const previous = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    setQuery(''); setActive(0);
    const timer = window.setTimeout(() => input.current?.focus(), 0);
    return () => { clearTimeout(timer); requestAnimationFrame(() => {
      const dialog = Array.from(document.querySelectorAll<HTMLElement>('[role="dialog"]')).find(element => !element.closest('[inert]'));
      if (dialog) { if (!dialog.contains(document.activeElement)) dialog.querySelector<HTMLElement>('input:not([hidden]), button:not([disabled])')?.focus({ preventScroll: true }); }
      else if (previous?.isConnected && !previous.closest('[inert]')) previous.focus({ preventScroll: true });
    }); };
  }, [props.open]);
  useEffect(() => { if (props.open) list.current?.querySelector<HTMLElement>('[aria-selected="true"]')?.scrollIntoView({ block: 'nearest' }); }, [selected, props.open, query]);
  if (!props.open) return null;
  const run = (index: number) => { const item = items[index]; if (item) { props.onClose(); item.run(); } };
  return <div className="command-backdrop" onMouseDown={props.onClose}>
    <section className="command-palette" role="dialog" aria-modal="true" aria-label="Search files and commands" onMouseDown={event => event.stopPropagation()} onKeyDown={event => {
      event.stopPropagation();
      if (event.key === 'Escape' || ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'k')) { event.preventDefault(); props.onClose(); return; }
      if (event.key === 'ArrowDown' || event.key === 'ArrowUp') { event.preventDefault(); setActive(items.length ? (selected + (event.key === 'ArrowDown' ? 1 : -1) + items.length) % items.length : 0); }
      if (event.key === 'Enter' && event.target !== close.current) { event.preventDefault(); run(selected); }
      if (event.key === 'Tab') { event.preventDefault(); if (document.activeElement === input.current) close.current?.focus(); else input.current?.focus(); }
    }} onKeyUp={event => event.stopPropagation()}>
      <div className="command-input-row"><Search size={21}/><input ref={input} role="combobox" aria-label="Search files and actions" aria-expanded="true" aria-controls={`${id}-results`} aria-autocomplete="list" aria-activedescendant={items.length ? `${id}-${selected}` : undefined} value={query} onChange={event => setQuery(event.target.value)} placeholder="Find a file, or do something…" autoComplete="off"/><button ref={close} className="icon-button" aria-label="Close command menu" onClick={props.onClose}><X size={18}/></button></div>
      <div className="command-results" role="listbox" aria-label="Files and actions" id={`${id}-results`} ref={list}>
        {items.map((item, index) => <div key={item.group + item.id}>
          {items[index - 1]?.group !== item.group && <div className="command-group" role="presentation">{item.group}</div>}
          <button role="option" aria-selected={selected === index} id={`${id}-${index}`} tabIndex={-1} className="command-option" onMouseMove={() => setActive(index)} onClick={() => run(index)}>
            <span className={`command-item-icon ${item.group === 'Actions' ? 'action' : ''}`}><item.icon size={18}/></span><span><strong>{item.label}</strong><small>{item.detail}</small></span>{item.group === 'Actions' && item.id.startsWith('doc') ? <Plus size={15}/> : <ArrowUpRight size={15}/>}
          </button>
        </div>)}
        {!items.length && <div className="command-empty"><Search size={26}/><strong>No matches yet</strong><p>Try a file name or an action like “import” or “new”.</p></div>}
      </div>
      <footer className="command-footer"><span><kbd>↑</kbd><kbd>↓</kbd> to move</span><span><kbd>↵</kbd> to open</span><span><kbd>esc</kbd> to close</span></footer>
    </section>
  </div>;
}
