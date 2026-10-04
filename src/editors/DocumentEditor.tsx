import { useEffect, useRef, useState, type CSSProperties, type ReactNode } from 'react';
import { EditorContent, useEditor, useEditorState } from '@tiptap/react';
import StarterKit from '@tiptap/starter-kit';
import TextAlign from '@tiptap/extension-text-align';
import { TextStyleKit } from '@tiptap/extension-text-style';
import Highlight from '@tiptap/extension-highlight';
import { TableKit } from '@tiptap/extension-table';
import Image from '@tiptap/extension-image';
import { ArrowLeft, FileText, ChevronDown, Download, Undo2, Redo2, Bold, Italic, Underline, Strikethrough, AlignLeft, AlignCenter, AlignRight, AlignJustify, List, ListOrdered, Link, ImagePlus, Table2, Printer, Search, X, Type, Highlighter, RemoveFormatting, Minus, Plus, CheckCheck, Quote, Maximize, PanelLeft, Eye, Pencil, Trash2, Columns3, Rows3 } from 'lucide-react';
import DOMPurify from 'dompurify';
import type { EditorProps } from '../types';
import { exportDocument } from './documentExport';
import './document-editor.css';

type Tab = 'Home' | 'Insert' | 'Layout' | 'Review' | 'View';
type Margin = 'normal' | 'narrow' | 'wide';
type Modal = 'link' | 'table' | null;

function Tool({ children, label, active = false, disabled = false, onClick, wide = false }: { children: ReactNode; label: string; active?: boolean; disabled?: boolean; onClick: () => void; wide?: boolean }) {
  return <button type="button" className={`doc-tool ${active ? 'is-active' : ''} ${wide ? 'doc-tool-wide' : ''}`} title={label} aria-label={label} aria-pressed={active} disabled={disabled} onMouseDown={e => e.preventDefault()} onClick={onClick}>{children}{wide && <span>{label}</span>}</button>;
}

export default function DocumentEditor({ file, onChange, onRename, onBack, onNotify, onPageSetupChange }: EditorProps) {
  const [tab, setTab] = useState<Tab>('Home');
  const [exportOpen, setExportOpen] = useState(false);
  const [findOpen, setFindOpen] = useState(false);
  const [find, setFind] = useState('');
  const [replacement, setReplacement] = useState('');
  const [findIndex, setFindIndex] = useState(-1);
  const [zoom, setZoom] = useState(90);
  const [landscape, setLandscape] = useState(file.pageSetup?.landscape ?? false);
  const [margin, setMargin] = useState<Margin>(file.pageSetup?.margin ?? 'normal');
  const [reading, setReading] = useState(false);
  const [outlineOpen, setOutlineOpen] = useState(false);
  const [modal, setModal] = useState<Modal>(null);
  const [linkValue, setLinkValue] = useState('');
  const [tableRows, setTableRows] = useState(3);
  const [tableCols, setTableCols] = useState(3);
  const [exporting, setExporting] = useState(false);
  const [title, setTitle] = useState(file.name);
  const imageInput = useRef<HTMLInputElement>(null);
  const findInput = useRef<HTMLInputElement>(null);
  const latestOnChange = useRef(onChange);
  latestOnChange.current = onChange;

  const editor = useEditor({
    immediatelyRender: false,
    extensions: [
      StarterKit.configure({ link: { openOnClick: false, HTMLAttributes: { target: '_blank', rel: 'noopener noreferrer' } } }),
      TextAlign.configure({ types: ['heading', 'paragraph'] }),
      TextStyleKit,
      Highlight.configure({ multicolor: true }),
      TableKit.configure({ table: { resizable: true } }),
      Image.configure({ allowBase64: true }),
    ],
    content: DOMPurify.sanitize(typeof file.content === 'string' ? file.content : '<p></p>'),
    editorProps: { attributes: { class: 'doc-prose', spellcheck: 'true', 'aria-label': 'Document content', role: 'textbox', 'aria-multiline': 'true' } },
    onUpdate: ({ editor: current }) => latestOnChange.current(current.getHTML()),
  });

  const editorState = useEditorState({
    editor,
    selector: ({ editor: current }) => current && !current.isDestroyed && current.schema ? {
      bold: current.isActive('bold'), italic: current.isActive('italic'), underline: current.isActive('underline'), strike: current.isActive('strike'),
      bullet: current.isActive('bulletList'), ordered: current.isActive('orderedList'), table: current.isActive('table'),
      align: ['left', 'center', 'right', 'justify'].find(a => current.isActive({ textAlign: a })) || 'left',
      heading: current.isActive('heading', { level: 1 }) ? '1' : current.isActive('heading', { level: 2 }) ? '2' : current.isActive('heading', { level: 3 }) ? '3' : '0',
      family: current.getAttributes('textStyle').fontFamily || 'Arial', size: current.getAttributes('textStyle').fontSize || '11pt',
      text: current.getText(), canUndo: current.can().undo(), canRedo: current.can().redo(),
      headings: current.getJSON().content?.filter(n => n.type === 'heading').map(n => ({ text: n.content?.map(c => 'text' in c ? c.text || '' : '').join('') || '', level: n.attrs?.level || 1 })) || [],
    } : null,
  });

  useEffect(() => { setTitle(file.name); }, [file.name]);
  useEffect(() => { if (findOpen) findInput.current?.focus(); }, [findOpen]);
  useEffect(() => { if (editor && !editor.isDestroyed) editor.setEditable(!reading); }, [editor, reading]);
  useEffect(() => {
    function onKey(event: KeyboardEvent) {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'f') { event.preventDefault(); setFindOpen(true); }
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's') { event.preventDefault(); onNotify('All changes are saved on this device.'); }
      if (event.key === 'Escape') { setModal(null); setExportOpen(false); setFindOpen(false); }
    }
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [onNotify]);

  const words = editorState?.text.trim() ? editorState.text.trim().split(/\s+/).length : 0;
  const characters = editorState?.text.length || 0;
  const getMatches = () => {
    const matches: { from: number; to: number }[] = [];
    if (!editor || !find) return matches;
    editor.state.doc.descendants((node, pos) => {
      if (!node.isText || !node.text) return;
      const text = node.text.toLowerCase();
      let index = text.indexOf(find.toLowerCase());
      while (index !== -1) { matches.push({ from: pos + index, to: pos + index + find.length }); index = text.indexOf(find.toLowerCase(), index + Math.max(find.length, 1)); }
    });
    return matches;
  };
  const findNext = () => {
    const matches = getMatches();
    if (!matches.length) { onNotify('No matching text found.'); return; }
    const next = (findIndex + 1) % matches.length;
    setFindIndex(next);
    editor?.chain().focus().setTextSelection(matches[next]).scrollIntoView().run();
  };
  const replaceAll = () => {
    if (!editor) return;
    const matches = getMatches();
    const chain = editor.chain();
    matches.slice().reverse().forEach(range => replacement ? chain.insertContentAt(range, { type: 'text', text: replacement }) : chain.deleteRange(range));
    chain.run(); setFindIndex(-1); onNotify(`Replaced ${matches.length} ${matches.length === 1 ? 'match' : 'matches'}.`);
  };
  const download = async (format: 'docx' | 'html' | 'txt') => {
    if (!editor) return;
    setExporting(true); setExportOpen(false);
    try { await exportDocument(editor.getHTML(), title || file.name, format, { landscape, margin }); onNotify(`Your ${format.toUpperCase()} file is ready.`); }
    catch (error) { console.error(error); onNotify('The export could not be completed. Please try again.'); }
    finally { setExporting(false); }
  };
  const commitTitle = () => { const value = title.trim() || 'Untitled document'; setTitle(value); if (value !== file.name) onRename(value); };
  const insertLink = () => {
    let url = linkValue.trim();
    if (!url) { editor?.chain().focus().extendMarkRange('link').unsetLink().run(); setModal(null); return; }
    if (!/^(https?:|mailto:|tel:)/i.test(url)) url = `https://${url}`;
    if (editor?.state.selection.empty) editor.chain().focus().insertContent({ type: 'text', text: linkValue, marks: [{ type: 'link', attrs: { href: url } }] }).run();
    else editor?.chain().focus().extendMarkRange('link').setLink({ href: url }).run();
    setModal(null);
  };
  const importImage = (selected?: File) => {
    if (!selected) return;
    if (!selected.type.startsWith('image/')) { onNotify('Choose an image file.'); return; }
    if (selected.size > 5 * 1024 * 1024) { onNotify('Please choose an image smaller than 5 MB.'); return; }
    const reader = new FileReader(); reader.onload = () => editor?.chain().focus().setImage({ src: String(reader.result), alt: selected.name }).run(); reader.readAsDataURL(selected);
    if (imageInput.current) imageInput.current.value = '';
  };
  const showLink = () => { setLinkValue(editor?.getAttributes('link').href || ''); setModal('link'); };
  const print = () => { setExportOpen(false); window.print(); };

  if (!editor || editor.isDestroyed || !editor.schema) return <div className="doc-loading">Opening your document…</div>;

  return <div className="doc-editor">
    <style media="print">{`@page folio-document { size: A4 ${landscape ? 'landscape' : 'portrait'}; margin: ${margin === 'narrow' ? '.5' : margin === 'wide' ? '1.5' : '1'}in; }`}</style>
    <header className="doc-topbar">
      <button className="doc-back" aria-label="Back to workspace" title="Back to workspace" onClick={onBack}><ArrowLeft size={20} /></button>
      <div className="doc-app-icon"><FileText size={23} strokeWidth={1.6} /></div>
      <div className="doc-title-block"><input aria-label="Document name" value={title} onChange={e => setTitle(e.target.value)} onBlur={commitTitle} onKeyDown={e => { if (e.key === 'Enter') e.currentTarget.blur(); }} /><span><CheckCheck size={13} /> Saved on this device</span></div>
      <div className="doc-header-actions"><button className="doc-mode" onClick={() => setReading(!reading)} title={reading ? 'Switch to editing' : 'Switch to reading'}>{reading ? <Eye size={15} /> : <Pencil size={15} />}<span>{reading ? 'Reading' : 'Editing'}</span><ChevronDown size={13} /></button>
        <div className="doc-export-wrap"><button className="doc-export" disabled={exporting} onClick={() => setExportOpen(!exportOpen)}><Download size={16} /><span>{exporting ? 'Exporting…' : 'Export'}</span><ChevronDown size={14} /></button>
          {exportOpen && <><button className="doc-menu-scrim" aria-label="Close export menu" onClick={() => setExportOpen(false)} /><div className="doc-export-menu"><div className="doc-menu-label">DOWNLOAD A COPY</div><button onClick={() => download('docx')}><FileText size={17} /><span>Word document<small>.docx · Microsoft Word compatible</small></span></button><button onClick={print}><Printer size={17} /><span>Print or save as PDF<small>Choose “Save as PDF” in print settings</small></span></button><button onClick={() => download('html')}><Columns3 size={17} /><span>Web page<small>.html · Keep text and formatting</small></span></button><button onClick={() => download('txt')}><Type size={17} /><span>Plain text<small>.txt · Text only</small></span></button></div></>}
        </div>
      </div>
    </header>
    <nav className="doc-tabs" aria-label="Document tools"><span className="doc-brand">folio<span> / </span>write</span>{(['Home', 'Insert', 'Layout', 'Review', 'View'] as Tab[]).map(t => <button key={t} className={tab === t ? 'is-active' : ''} onClick={() => setTab(t)}>{t}</button>)}<button className="doc-find-shortcut" onClick={() => setFindOpen(!findOpen)}><Search size={15} /><span>Find in document</span><kbd>⌘ F</kbd></button></nav>
    <fieldset className={`doc-ribbon ${reading ? 'doc-ribbon-reading' : ''}`} disabled={reading && ['Home', 'Insert', 'Layout'].includes(tab)} aria-label={`${tab} ribbon`}>
      {tab === 'Home' && <>
        <div className="doc-ribbon-group"><div className="doc-tools-row"><Tool label="Undo" disabled={!editorState?.canUndo} onClick={() => editor.chain().focus().undo().run()}><Undo2 size={18} /></Tool><Tool label="Redo" disabled={!editorState?.canRedo} onClick={() => editor.chain().focus().redo().run()}><Redo2 size={18} /></Tool></div><span className="doc-group-label">History</span></div>
        <div className="doc-ribbon-group doc-font-group"><div className="doc-tools-row"><select aria-label="Font family" value={editorState?.family} onChange={e => editor.chain().focus().setFontFamily(e.target.value).run()}>{['Arial', 'Georgia', 'Times New Roman', 'Verdana', 'Courier New', 'Inter'].map(f => <option key={f}>{f}</option>)}</select><select className="doc-size-select" aria-label="Font size" value={editorState?.size} onChange={e => editor.chain().focus().setFontSize(e.target.value).run()}>{['8', '9', '10', '11', '12', '14', '16', '18', '20', '24', '28', '32', '36', '48', '72'].map(s => <option value={`${s}pt`} key={s}>{s}</option>)}</select></div><div className="doc-tools-row"><Tool label="Bold (Ctrl+B)" active={editorState?.bold} onClick={() => editor.chain().focus().toggleBold().run()}><Bold size={16} /></Tool><Tool label="Italic (Ctrl+I)" active={editorState?.italic} onClick={() => editor.chain().focus().toggleItalic().run()}><Italic size={16} /></Tool><Tool label="Underline (Ctrl+U)" active={editorState?.underline} onClick={() => editor.chain().focus().toggleUnderline().run()}><Underline size={16} /></Tool><Tool label="Strikethrough" active={editorState?.strike} onClick={() => editor.chain().focus().toggleStrike().run()}><Strikethrough size={16} /></Tool><label className="doc-color-tool" title="Text color"><Type size={17} /><span style={{ background: editor.getAttributes('textStyle').color || '#253a34' }} /><input aria-label="Text color" type="color" defaultValue="#253a34" onChange={e => editor.chain().focus().setColor(e.target.value).run()} /></label><label className="doc-color-tool" title="Highlight color"><Highlighter size={16} /><span style={{ background: editor.getAttributes('highlight').color || '#fae797' }} /><input aria-label="Highlight color" type="color" defaultValue="#fae797" onChange={e => editor.chain().focus().setHighlight({ color: e.target.value }).run()} /></label><Tool label="Clear formatting" onClick={() => editor.chain().focus().clearNodes().unsetAllMarks().run()}><RemoveFormatting size={16} /></Tool></div><span className="doc-group-label">Font</span></div>
        <div className="doc-ribbon-group"><div className="doc-tools-row"><Tool label="Bulleted list" active={editorState?.bullet} onClick={() => editor.chain().focus().toggleBulletList().run()}><List size={18} /></Tool><Tool label="Numbered list" active={editorState?.ordered} onClick={() => editor.chain().focus().toggleOrderedList().run()}><ListOrdered size={18} /></Tool><select className="doc-line-select" title="Line spacing" aria-label="Line spacing" defaultValue="1.65" onChange={e => editor.chain().focus().setLineHeight(e.target.value).run()}><option value="1">1.0</option><option value="1.15">1.15</option><option value="1.5">1.5</option><option value="1.65">1.65</option><option value="2">2.0</option></select></div><div className="doc-tools-row">{[{ value: 'left', Icon: AlignLeft }, { value: 'center', Icon: AlignCenter }, { value: 'right', Icon: AlignRight }, { value: 'justify', Icon: AlignJustify }].map(({ value, Icon }) => <Tool key={value} label={`Align ${value}`} active={editorState?.align === value} onClick={() => editor.chain().focus().setTextAlign(value).run()}><Icon size={17} /></Tool>)}</div><span className="doc-group-label">Paragraph</span></div>
        <div className="doc-ribbon-group doc-styles-group"><div className="doc-style-options">{[{ value: '0', name: 'Normal', sample: 'Aa' }, { value: '1', name: 'Heading 1', sample: 'Aa' }, { value: '2', name: 'Heading 2', sample: 'Aa' }, { value: '3', name: 'Heading 3', sample: 'Aa' }].map(s => <button key={s.value} className={`doc-style-button doc-style-${s.value} ${editorState?.heading === s.value ? 'is-active' : ''}`} onMouseDown={e => e.preventDefault()} onClick={() => s.value === '0' ? editor.chain().focus().setParagraph().run() : editor.chain().focus().setHeading({ level: Number(s.value) as 1 | 2 | 3 }).run()}><strong>{s.sample}</strong><span>{s.name}</span></button>)}</div><span className="doc-group-label">Styles</span></div>
      </>}
      {tab === 'Insert' && <>
        <div className="doc-ribbon-group"><div className="doc-large-tools"><Tool wide label="Table" onClick={() => setModal('table')}><Table2 size={23} /></Tool><Tool wide label="Image" onClick={() => imageInput.current?.click()}><ImagePlus size={23} /></Tool><Tool wide label="Link" onClick={showLink}><Link size={23} /></Tool></div><span className="doc-group-label">Add to your document</span></div>
        <div className="doc-ribbon-group"><div className="doc-large-tools"><Tool wide label="Divider" onClick={() => editor.chain().focus().setHorizontalRule().run()}><Minus size={23} /></Tool><Tool wide label="Quote" onClick={() => editor.chain().focus().toggleBlockquote().run()}><Quote size={23} /></Tool><Tool wide label="Date" onClick={() => editor.chain().focus().insertContent(new Date().toLocaleDateString(undefined, { year: 'numeric', month: 'long', day: 'numeric' })).run()}><Type size={23} /></Tool></div><span className="doc-group-label">Elements</span></div>
        {editorState?.table && <div className="doc-ribbon-group"><div className="doc-large-tools"><Tool wide label="Add row" onClick={() => editor.chain().focus().addRowAfter().run()}><Rows3 size={23} /></Tool><Tool wide label="Add column" onClick={() => editor.chain().focus().addColumnAfter().run()}><Columns3 size={23} /></Tool><Tool wide label="Delete row" onClick={() => editor.chain().focus().deleteRow().run()}><Rows3 size={23} /></Tool><Tool wide label="Delete column" onClick={() => editor.chain().focus().deleteColumn().run()}><Columns3 size={23} /></Tool><Tool wide label="Delete table" onClick={() => editor.chain().focus().deleteTable().run()}><Trash2 size={23} /></Tool></div><span className="doc-group-label">Table tools</span></div>}
      </>}
      {tab === 'Layout' && <>
        <div className="doc-ribbon-group"><div className="doc-layout-control"><span>Page orientation</span><select aria-label="Page orientation" value={landscape ? 'landscape' : 'portrait'} onChange={e => { const value = e.target.value === 'landscape'; setLandscape(value); onPageSetupChange?.({ landscape: value, margin }); }}><option value="portrait">Portrait</option><option value="landscape">Landscape</option></select></div><span className="doc-group-label">Page setup</span></div>
        <div className="doc-ribbon-group"><div className="doc-layout-control"><span>Page margins</span><select aria-label="Page margins" value={margin} onChange={e => { const value = e.target.value as Margin; setMargin(value); onPageSetupChange?.({ landscape, margin: value }); }}><option value="normal">Normal · 1 inch</option><option value="narrow">Narrow · 0.5 inch</option><option value="wide">Wide · 1.5 inch</option></select></div><span className="doc-group-label">Spacing</span></div><p className="doc-ribbon-hint">Give your ideas a little room.<br />Page settings are included in your Word export.</p>
      </>}
      {tab === 'Review' && <><div className="doc-ribbon-group"><div className="doc-large-tools"><Tool wide label="Find & replace" onClick={() => setFindOpen(!findOpen)}><Search size={23} /></Tool><Tool wide label="Spellcheck" onClick={() => { const enabled = editor.view.dom.getAttribute('spellcheck') !== 'false'; editor.view.dom.setAttribute('spellcheck', String(!enabled)); onNotify(`Browser spellcheck ${enabled ? 'disabled' : 'enabled'}.`); }}><CheckCheck size={23} /></Tool></div><span className="doc-group-label">Proofing</span></div><div className="doc-word-stats"><strong>{words.toLocaleString()}<span>words</span></strong><strong>{characters.toLocaleString()}<span>characters</span></strong><strong>{editorState?.text.split('\n').filter(Boolean).length || 0}<span>paragraphs</span></strong></div></>}
      {tab === 'View' && <><div className="doc-ribbon-group"><div className="doc-large-tools"><Tool wide label="Outline" active={outlineOpen} onClick={() => setOutlineOpen(!outlineOpen)}><PanelLeft size={23} /></Tool><Tool wide label={reading ? 'Edit document' : 'Reading mode'} active={reading} onClick={() => setReading(!reading)}><Eye size={23} /></Tool><Tool wide label="100%" onClick={() => setZoom(100)}><Maximize size={23} /></Tool></div><span className="doc-group-label">Workspace</span></div><div className="doc-ribbon-group"><div className="doc-large-tools"><Tool wide label="Print / PDF" onClick={print}><Printer size={23} /></Tool></div><span className="doc-group-label">Print</span></div></>}
    </fieldset>
    {findOpen && <div className="doc-find-panel"><Search size={17} /><input ref={findInput} aria-label="Find text" placeholder="Find in document" value={find} onChange={e => { setFind(e.target.value); setFindIndex(-1); }} onKeyDown={e => { if (e.key === 'Enter') findNext(); }} /><button onClick={findNext} disabled={!find}>Find next</button><span className="doc-find-separator" /><input aria-label="Replace text" placeholder="Replace with" value={replacement} onChange={e => setReplacement(e.target.value)} /><button onClick={replaceAll} disabled={!find || reading}>Replace all</button><small>{find ? `${getMatches().length} matches` : ''}</small><button className="doc-find-close" aria-label="Close find and replace" onClick={() => setFindOpen(false)}><X size={17} /></button></div>}
    <main className="doc-work-area">
      {outlineOpen && <aside className="doc-outline"><div><strong>Document outline</strong><button aria-label="Close outline" onClick={() => setOutlineOpen(false)}><X size={16} /></button></div>{editorState?.headings.length ? editorState.headings.map((h, i) => <button className="doc-outline-item" style={{ paddingLeft: 12 + (h.level - 1) * 12 }} key={i} onClick={() => { const headings = editor.view.dom.querySelectorAll('h1,h2,h3,h4,h5,h6'); headings[i]?.scrollIntoView({ behavior: 'smooth', block: 'center' }); }}>{h.text || 'Untitled heading'}</button>) : <p>Add headings to your document to see its outline here.</p>}</aside>}
      <div className="doc-paper-scroll"><div className="doc-page-caption"><span>{reading ? 'READING VIEW' : 'YOUR WORDS, YOUR SPACE'}</span><span>A4 · {landscape ? 'Landscape' : 'Portrait'}</span></div><div className={`doc-sheet doc-margin-${margin} ${landscape ? 'doc-landscape' : ''}`} style={{ zoom: zoom / 100, '--doc-zoom': zoom / 100 } as CSSProperties}><EditorContent editor={editor} /></div><div className="doc-page-end">Made with folio. Made by you.</div></div>
    </main>
    <footer className="doc-status"><span><span className="doc-status-dot" /> All changes saved</span><span>{words.toLocaleString()} words</span><span className="doc-status-local">Private · Stored on this device</span><div className="doc-zoom"><button aria-label="Zoom out" disabled={zoom <= 50} onClick={() => setZoom(Math.max(50, zoom - 10))}><Minus size={14} /></button><input aria-label="Zoom level" type="range" min="50" max="150" step="10" value={zoom} onChange={e => setZoom(Number(e.target.value))} /><button aria-label="Zoom in" disabled={zoom >= 150} onClick={() => setZoom(Math.min(150, zoom + 10))}><Plus size={14} /></button><button className="doc-zoom-value" onClick={() => setZoom(100)}>{zoom}%</button></div></footer>
    <input type="file" ref={imageInput} accept="image/*" hidden onChange={e => importImage(e.target.files?.[0])} />
    {modal && <div className="doc-modal-scrim" onMouseDown={e => { if (e.target === e.currentTarget) setModal(null); }}><section className="doc-modal" role="dialog" aria-modal="true" aria-label={modal === 'link' ? 'Insert link' : 'Insert table'}><button className="doc-modal-close" aria-label="Close dialog" onClick={() => setModal(null)}><X size={19} /></button><div className="doc-modal-icon">{modal === 'link' ? <Link size={25} /> : <Table2 size={25} />}</div><h2>{modal === 'link' ? 'Add a link' : 'Make room for details'}</h2><p>{modal === 'link' ? 'Connect your words to something useful.' : 'Insert a table. You can add more rows and columns later.'}</p>{modal === 'link' ? <form onSubmit={e => { e.preventDefault(); insertLink(); }}><label>Web address<input autoFocus value={linkValue} onChange={e => setLinkValue(e.target.value)} placeholder="https://example.com" /></label><div className="doc-modal-actions"><button type="button" onClick={() => setModal(null)}>Cancel</button><button className="doc-primary" type="submit">{linkValue ? 'Apply link' : 'Remove link'}</button></div></form> : <form onSubmit={e => { e.preventDefault(); editor.chain().focus().insertTable({ rows: tableRows, cols: tableCols, withHeaderRow: true }).run(); setModal(null); }}><div className="doc-table-inputs"><label>Rows<input type="number" min="1" max="30" value={tableRows} onChange={e => setTableRows(Math.min(30, Math.max(1, Number(e.target.value))))} /></label><label>Columns<input type="number" min="1" max="12" value={tableCols} onChange={e => setTableCols(Math.min(12, Math.max(1, Number(e.target.value))))} /></label></div><div className="doc-modal-actions"><button type="button" onClick={() => setModal(null)}>Cancel</button><button className="doc-primary" type="submit">Insert table</button></div></form>}</section></div>}
  </div>;
}
