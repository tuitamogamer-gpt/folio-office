import { useEffect, useRef, useState, type CSSProperties, type ReactNode } from 'react';
import { EditorContent, useEditor, useEditorState } from '@tiptap/react';
import StarterKit from '@tiptap/starter-kit';
import TextAlign from '@tiptap/extension-text-align';
import { TextStyleKit } from '@tiptap/extension-text-style';
import Highlight from '@tiptap/extension-highlight';
import { TableKit } from '@tiptap/extension-table';
import Subscript from '@tiptap/extension-subscript';
import Superscript from '@tiptap/extension-superscript';
import type { Mark } from '@tiptap/pm/model';
import { ArrowLeft, FileText, ChevronDown, Download, Undo2, Redo2, Bold, Italic, Underline, Strikethrough, AlignLeft, AlignCenter, AlignRight, AlignJustify, List, ListOrdered, Link, ImagePlus, Table2, Printer, Search, X, Type, Highlighter, RemoveFormatting, Minus, Plus, CheckCheck, Quote, Maximize, PanelLeft, Eye, Pencil, Trash2, Columns3, Rows3, Subscript as SubscriptIcon, Superscript as SuperscriptIcon, IndentIncrease, IndentDecrease, Paintbrush, MessageSquare, History, FilePlus2, BetweenHorizontalStart, Merge, Split, Check, RotateCcw } from 'lucide-react';
import DOMPurify from 'dompurify';
import type { EditorProps, DocumentPageSetup } from '../types';
import { exportDocument } from './documentExport';
import { DocumentAttributes, DocumentImage, PageBreak, CommentAnchor } from './documentExtensions';
import './document-editor.css';

type Tab = 'Home' | 'Insert' | 'Layout' | 'Review' | 'View';
type Margin = 'normal' | 'narrow' | 'wide';
type Modal = 'link' | 'table' | 'headerFooter' | null;

function Tool({ children, label, active = false, disabled = false, onClick, wide = false }: { children: ReactNode; label: string; active?: boolean; disabled?: boolean; onClick: () => void; wide?: boolean }) {
  return <button type="button" className={`doc-tool ${active ? 'is-active' : ''} ${wide ? 'doc-tool-wide' : ''}`} title={label} aria-label={label} aria-pressed={active} disabled={disabled} onMouseDown={e => e.preventDefault()} onClick={onClick}>{children}{wide && <span>{label}</span>}</button>;
}

export default function DocumentEditor({ file, onChange, onRename, onBack, onNotify, onPageSetupChange, onCommentsChange, onOpenVersions, saveState = 'saved' }: EditorProps) {
  const [tab, setTab] = useState<Tab>('Home');
  const [exportOpen, setExportOpen] = useState(false);
  const [findOpen, setFindOpen] = useState(false);
  const [find, setFind] = useState('');
  const [replacement, setReplacement] = useState('');
  const [findIndex, setFindIndex] = useState(-1);
  const [matchCase, setMatchCase] = useState(false);
  const [wholeWord, setWholeWord] = useState(false);
  const [zoom, setZoom] = useState(90);
  const [landscape, setLandscape] = useState(file.pageSetup?.landscape ?? false);
  const [margin, setMargin] = useState<Margin>(file.pageSetup?.margin ?? 'normal');
  const [pageSetup, setPageSetup] = useState<DocumentPageSetup>(file.pageSetup || { landscape: false, margin: 'normal', size: 'a4' });
  const [headerDraft, setHeaderDraft] = useState('');
  const [footerDraft, setFooterDraft] = useState('');
  const [numbersDraft, setNumbersDraft] = useState(false);
  const [reading, setReading] = useState(false);
  const [outlineOpen, setOutlineOpen] = useState(false);
  const [commentsOpen, setCommentsOpen] = useState(false);
  const [commentDraft, setCommentDraft] = useState('');
  const [commentRange, setCommentRange] = useState<{ from: number; to: number; quote: string } | null>(null);
  const [showResolved, setShowResolved] = useState(false);
  const [formatMarks, setFormatMarks] = useState<Mark[] | null>(null);
  const [imageWidthDraft, setImageWidthDraft] = useState('');
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
      DocumentImage, DocumentAttributes, PageBreak, CommentAnchor,
      Subscript.extend({ excludes: 'subscript superscript' }), Superscript.extend({ excludes: 'subscript superscript' }),
    ],
    content: DOMPurify.sanitize(typeof file.content === 'string' ? file.content : '<p></p>'),
    editorProps: { attributes: { class: 'doc-prose', spellcheck: 'true', 'aria-label': 'Document content', role: 'textbox', 'aria-multiline': 'true' } },
    onUpdate: ({ editor: current }) => latestOnChange.current(current.getHTML()),
  });

  const editorState = useEditorState({
    editor,
    selector: ({ editor: current }) => current && !current.isDestroyed && current.schema ? {
      bold: current.isActive('bold'), italic: current.isActive('italic'), underline: current.isActive('underline'), strike: current.isActive('strike'),
      subscript: current.isActive('subscript'), superscript: current.isActive('superscript'), image: current.isActive('image'), imageAttrs: current.getAttributes('image'),
      bullet: current.isActive('bulletList'), ordered: current.isActive('orderedList'), table: current.isActive('table'),
      align: ['left', 'center', 'right', 'justify'].find(a => current.isActive({ textAlign: a })) || 'left',
      heading: current.isActive('heading', { level: 1 }) ? '1' : current.isActive('heading', { level: 2 }) ? '2' : current.isActive('heading', { level: 3 }) ? '3' : '0',
      family: current.getAttributes('textStyle').fontFamily || 'Arial', size: current.getAttributes('textStyle').fontSize || '11pt',
      text: current.getText(), canUndo: current.can().undo(), canRedo: current.can().redo(),
      selectionFrom: current.state.selection.from, selectionTo: current.state.selection.to,
      lineHeight: current.getAttributes('paragraph').lineHeight || current.getAttributes('heading').lineHeight || '1.65',
      indent: current.getAttributes('paragraph').indent || current.getAttributes('heading').indent || 0,
      canMerge: current.can().mergeCells(), canSplit: current.can().splitCell(),
      headings: current.getJSON().content?.filter(n => n.type === 'heading').map(n => ({ text: n.content?.map(c => 'text' in c ? c.text || '' : '').join('') || '', level: n.attrs?.level || 1 })) || [],
    } : null,
  });

  useEffect(() => { setTitle(file.name); }, [file.name]);
  useEffect(() => { setImageWidthDraft(editorState?.imageAttrs.width ? String(editorState.imageAttrs.width) : ''); }, [editorState?.imageAttrs.width, editorState?.selectionFrom]);
  useEffect(() => {
    const next = file.pageSetup || { landscape: false, margin: 'normal' as const, size: 'a4' as const };
    setPageSetup(next); setLandscape(next.landscape); setMargin(next.margin);
  }, [file.pageSetup]);
  useEffect(() => {
    if (!editor || editor.isDestroyed || !editor.schema || typeof file.content !== 'string' || file.content === editor.getHTML()) return;
    editor.commands.setContent(DOMPurify.sanitize(file.content), { emitUpdate: false });
  }, [file.content, editor]);
  useEffect(() => { if (findOpen) findInput.current?.focus(); }, [findOpen]);
  useEffect(() => { if (editor && !editor.isDestroyed) editor.setEditable(!reading); }, [editor, reading]);
  useEffect(() => {
    function onKey(event: KeyboardEvent) {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'f') { event.preventDefault(); setFindOpen(true); }
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's') { event.preventDefault(); onNotify(saveState === 'saved' ? 'All changes are saved on this device.' : saveState === 'error' ? 'Saving failed. Export a copy to keep your work.' : 'Saving your changes…'); }
      if (event.key === 'Escape') { setModal(null); setExportOpen(false); setFindOpen(false); }
    }
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [onNotify, saveState]);

  const updatePageSetup = (values: Partial<DocumentPageSetup>) => {
    const next = { ...pageSetup, ...values }; setPageSetup(next); setLandscape(next.landscape); setMargin(next.margin); onPageSetupChange?.(next);
  };
  const saveLabel = saveState === 'saving' ? 'Saving changes…' : saveState === 'error' ? 'Could not save · Export a copy' : 'Saved on this device';

  const words = editorState?.text.trim() ? editorState.text.trim().split(/\s+/).length : 0;
  const characters = editorState?.text.length || 0;
  const getMatches = () => {
    const matches: { from: number; to: number }[] = [];
    if (!editor || !find) return matches;
    editor.state.doc.descendants((node, pos) => {
      if (!node.isTextblock) return;
      const original = node.textBetween(0, node.content.size, '', '\ufffc');
      const text = matchCase ? original : original.toLocaleLowerCase();
      const needle = matchCase ? find : find.toLocaleLowerCase();
      let index = text.indexOf(needle);
      while (index !== -1) {
        const isWord = (character: string) => /[\p{L}\p{N}_]/u.test(character);
        if (!wholeWord || ((!index || !isWord(original[index - 1])) && (index + find.length >= original.length || !isWord(original[index + find.length])))) matches.push({ from: pos + 1 + index, to: pos + 1 + index + find.length });
        index = text.indexOf(needle, index + Math.max(find.length, 1));
      }
      return false;
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
    if (!editor || reading) return;
    const matches = getMatches();
    const transaction = editor.state.tr;
    matches.slice().reverse().forEach(range => transaction.insertText(replacement, range.from, range.to));
    editor.view.dispatch(transaction); setFindIndex(-1); onNotify(`Replaced ${matches.length} ${matches.length === 1 ? 'match' : 'matches'}.`);
  };
  const replaceSelected = () => {
    if (!editor || reading) return;
    const { from, to } = editor.state.selection;
    const match = getMatches().find(range => range.from === from && range.to === to);
    if (!match) { findNext(); return; }
    editor.view.dispatch(editor.state.tr.insertText(replacement, from, to)); setFindIndex(-1);
  };
  const indentParagraph = (delta: number) => {
    if (!editor || reading) return;
    const { from, to } = editor.state.selection;
    const transaction = editor.state.tr;
    editor.state.doc.nodesBetween(from, to, (node, pos) => {
      if (node.type.name === 'paragraph' || node.type.name === 'heading') transaction.setNodeMarkup(pos, undefined, { ...node.attrs, indent: Math.max(0, Math.min(8, (Number(node.attrs.indent) || 0) + delta)) });
    });
    editor.view.dispatch(transaction); editor.commands.focus();
  };
  const changeCase = (mode: string) => {
    if (!editor || reading || editor.state.selection.empty) return;
    const { from, to } = editor.state.selection;
    const edits: { from: number; to: number; text: string; marks: readonly Mark[] }[] = [];
    editor.state.doc.nodesBetween(from, to, (node, pos) => {
      if (!node.isText || !node.text) return;
      const start = Math.max(from, pos), end = Math.min(to, pos + node.nodeSize);
      const selected = node.text.slice(start - pos, end - pos);
      const before = editor.state.doc.textBetween(Math.max(0, start - 1), start, ' ');
      const text = mode === 'upper' ? selected.toLocaleUpperCase() : mode === 'lower' ? selected.toLocaleLowerCase() : (before + selected.toLocaleLowerCase()).replace(/(^|[^\p{L}\p{N}])(\p{L})/gu, (_, prefix: string, letter: string) => prefix + letter.toLocaleUpperCase()).slice(before.length);
      edits.push({ from: start, to: end, text, marks: node.marks });
    });
    const transaction = editor.state.tr;
    edits.reverse().forEach(edit => transaction.replaceWith(edit.from, edit.to, editor.schema.text(edit.text, edit.marks)));
    editor.view.dispatch(transaction); editor.commands.focus();
  };
  const formatPainter = () => {
    if (!editor || reading) return;
    const { from, to, $from } = editor.state.selection;
    const names = ['bold', 'italic', 'underline', 'strike', 'textStyle', 'highlight', 'subscript', 'superscript'];
    if (formatMarks === null) {
      let marks = editor.state.storedMarks || $from.marks();
      // A selection can start at a mark boundary; resolve the first selected text explicitly.
      if (from !== to) { let found = false; editor.state.doc.nodesBetween(from, to, node => { if (!found && node.isText) { marks = node.marks; found = true; } }); }
      setFormatMarks(marks.filter(mark => names.includes(mark.type.name))); onNotify('Formatting copied. Select text, then click Apply formatting.');
    } else {
      if (from === to) { onNotify('Select the text you want to format.'); return; }
      const transaction = editor.state.tr;
      names.forEach(name => { const type = editor.schema.marks[name]; if (type) transaction.removeMark(from, to, type); });
      formatMarks.forEach(mark => transaction.addMark(from, to, mark));
      editor.view.dispatch(transaction); setFormatMarks(null); editor.commands.focus();
    }
  };
  const beginComment = () => {
    if (!editor || reading || !onCommentsChange) return;
    const { from, to } = editor.state.selection;
    setCommentsOpen(true);
    if (from === to) { setCommentRange(null); onNotify('Select text in the document, then click New comment.'); return; }
    setCommentRange({ from, to, quote: editor.state.doc.textBetween(from, to, ' ') }); setCommentDraft('');
  };
  const addComment = () => {
    if (!editor || !commentRange || !commentDraft.trim() || !onCommentsChange || reading) return;
    const { from, to, quote } = commentRange;
    if (to > editor.state.doc.content.size || editor.state.doc.textBetween(from, to, ' ') !== quote) { setCommentRange(null); onNotify('The selection changed. Select text again to add your comment.'); return; }
    const id = crypto.randomUUID();
    editor.chain().focus().setTextSelection({ from, to }).setMark('commentAnchor', { commentId: id }).run();
    onCommentsChange([...(file.comments || []), { id, text: commentDraft.trim(), quote, createdAt: Date.now(), resolved: false }]); setCommentDraft(''); setCommentRange(null);
  };
  const focusComment = (id: string) => {
    if (!editor) return;
    let first = -1, last = -1;
    editor.state.doc.descendants((node, pos) => { if (node.marks.some(mark => mark.type.name === 'commentAnchor' && mark.attrs.commentId === id)) { if (first < 0) first = pos; last = pos + node.nodeSize; } });
    if (first >= 0) editor.chain().focus().setTextSelection({ from: first, to: last }).scrollIntoView().run();
    else onNotify('The text for this comment has been removed.');
  };
  const deleteComment = (id: string) => {
    if (!editor || reading || !onCommentsChange) return;
    const transaction = editor.state.tr;
    editor.state.doc.descendants((node, pos) => { const mark = node.marks.find(item => item.type.name === 'commentAnchor' && item.attrs.commentId === id); if (mark) transaction.removeMark(pos, pos + node.nodeSize, mark); });
    editor.view.dispatch(transaction); onCommentsChange((file.comments || []).filter(comment => comment.id !== id));
  };
  const openHeaderFooter = () => { setHeaderDraft(pageSetup.header || ''); setFooterDraft(pageSetup.footer || ''); setNumbersDraft(!!pageSetup.pageNumbers); setModal('headerFooter'); };
  const download = async (format: 'docx' | 'html' | 'txt') => {
    if (!editor) return;
    setExporting(true); setExportOpen(false);
    try { await exportDocument(editor.getHTML(), title || file.name, format, { ...pageSetup, comments: file.comments }); onNotify(`Your ${format.toUpperCase()} file is ready.`); }
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
  const pageSize = pageSetup.size || 'a4';
  const paperDimensions = pageSize === 'letter' ? [816, 1056] : pageSize === 'legal' ? [816, 1344] : [794, 1123];
  const comments = file.comments || [];
  const selectedText = editorState?.selectionFrom !== editorState?.selectionTo;

  return <div className="doc-editor">
    <style media="print">{`@page folio-document { size: ${pageSize} ${landscape ? 'landscape' : 'portrait'}; margin: ${margin === 'narrow' ? '.5' : margin === 'wide' ? '1.5' : '1'}in; @top-center { content: ${JSON.stringify(pageSetup.header || '')}; font: 10pt Arial, sans-serif; color: #626c62; } @bottom-left { content: ${JSON.stringify(pageSetup.footer || '')}; font: 10pt Arial, sans-serif; color: #626c62; } @bottom-right { content: ${pageSetup.pageNumbers ? 'counter(page)' : 'none'}; font: 10pt Arial, sans-serif; color: #626c62; } }`}</style>
    <header className="doc-topbar">
      <button className="doc-back" aria-label="Back to workspace" title="Back to workspace" onClick={onBack}><ArrowLeft size={20} /></button>
      <div className="doc-app-icon"><FileText size={23} strokeWidth={1.6} /></div>
      <div className="doc-title-block"><input aria-label="Document name" value={title} disabled={reading} onChange={e => setTitle(e.target.value)} onBlur={commitTitle} onKeyDown={e => { if (e.key === 'Enter') e.currentTarget.blur(); }} /><span className={saveState === 'error' ? 'doc-save-error' : ''}><CheckCheck size={13} /> {saveLabel}</span></div>
      <div className="doc-header-actions"><button className="doc-mode doc-comments-toggle" aria-label="Show comments" onClick={() => setCommentsOpen(!commentsOpen)}><MessageSquare size={16} /><span>Comments{comments.filter(comment => !comment.resolved).length ? ` (${comments.filter(comment => !comment.resolved).length})` : ''}</span></button><button className="doc-mode" onClick={() => setReading(!reading)} title={reading ? 'Switch to editing' : 'Switch to reading'}>{reading ? <Eye size={15} /> : <Pencil size={15} />}<span>{reading ? 'Reading' : 'Editing'}</span><ChevronDown size={13} /></button>
        <div className="doc-export-wrap"><button className="doc-export" disabled={exporting} onClick={() => setExportOpen(!exportOpen)}><Download size={16} /><span>{exporting ? 'Exporting…' : 'Export'}</span><ChevronDown size={14} /></button>
          {exportOpen && <><button className="doc-menu-scrim" aria-label="Close export menu" onClick={() => setExportOpen(false)} /><div className="doc-export-menu"><div className="doc-menu-label">DOWNLOAD A COPY</div><button onClick={() => download('docx')}><FileText size={17} /><span>Word document<small>.docx · Microsoft Word compatible</small></span></button><button onClick={print}><Printer size={17} /><span>Print or save as PDF<small>Choose “Save as PDF” in print settings</small></span></button><button onClick={() => download('html')}><Columns3 size={17} /><span>Web page<small>.html · Keep text and formatting</small></span></button><button onClick={() => download('txt')}><Type size={17} /><span>Plain text<small>.txt · Text only</small></span></button></div></>}
        </div>
      </div>
    </header>
    <nav className="doc-tabs" aria-label="Document tools"><span className="doc-brand">folio<span> / </span>write</span>{(['Home', 'Insert', 'Layout', 'Review', 'View'] as Tab[]).map(t => <button key={t} className={tab === t ? 'is-active' : ''} onClick={() => setTab(t)}>{t}</button>)}<button className="doc-find-shortcut" onClick={() => setFindOpen(!findOpen)}><Search size={15} /><span>Find in document</span><kbd>⌘ F</kbd></button></nav>
    <fieldset className={`doc-ribbon ${reading ? 'doc-ribbon-reading' : ''}`} disabled={reading && ['Home', 'Insert', 'Layout'].includes(tab)} aria-label={`${tab} ribbon`}>
      {tab === 'Home' && <>
        <div className="doc-ribbon-group"><div className="doc-tools-row"><Tool label="Undo" disabled={!editorState?.canUndo} onClick={() => editor.chain().focus().undo().run()}><Undo2 size={18} /></Tool><Tool label="Redo" disabled={!editorState?.canRedo} onClick={() => editor.chain().focus().redo().run()}><Redo2 size={18} /></Tool></div><span className="doc-group-label">History</span></div>
        <div className="doc-ribbon-group doc-font-group"><div className="doc-tools-row"><select aria-label="Font family" value={editorState?.family} onChange={e => editor.chain().focus().setFontFamily(e.target.value).run()}>{['Arial', 'Georgia', 'Times New Roman', 'Verdana', 'Courier New', 'Inter'].map(f => <option key={f}>{f}</option>)}</select><select className="doc-size-select" aria-label="Font size" value={editorState?.size} onChange={e => editor.chain().focus().setFontSize(e.target.value).run()}>{['8', '9', '10', '11', '12', '14', '16', '18', '20', '24', '28', '32', '36', '48', '72'].map(s => <option value={`${s}pt`} key={s}>{s}</option>)}</select></div><div className="doc-tools-row"><Tool label="Bold (Ctrl+B)" active={editorState?.bold} onClick={() => editor.chain().focus().toggleBold().run()}><Bold size={16} /></Tool><Tool label="Italic (Ctrl+I)" active={editorState?.italic} onClick={() => editor.chain().focus().toggleItalic().run()}><Italic size={16} /></Tool><Tool label="Underline (Ctrl+U)" active={editorState?.underline} onClick={() => editor.chain().focus().toggleUnderline().run()}><Underline size={16} /></Tool><Tool label="Strikethrough" active={editorState?.strike} onClick={() => editor.chain().focus().toggleStrike().run()}><Strikethrough size={16} /></Tool><label className="doc-color-tool" title="Text color"><Type size={17} /><span style={{ background: editor.getAttributes('textStyle').color || '#253a34' }} /><input aria-label="Text color" type="color" defaultValue="#253a34" onChange={e => editor.chain().focus().setColor(e.target.value).run()} /></label><label className="doc-color-tool" title="Highlight color"><Highlighter size={16} /><span style={{ background: editor.getAttributes('highlight').color || '#fae797' }} /><input aria-label="Highlight color" type="color" defaultValue="#fae797" onChange={e => editor.chain().focus().setHighlight({ color: e.target.value }).run()} /></label><Tool label="Subscript" active={editorState?.subscript} onClick={() => editor.chain().focus().unsetSuperscript().toggleSubscript().run()}><SubscriptIcon size={16} /></Tool><Tool label="Superscript" active={editorState?.superscript} onClick={() => editor.chain().focus().unsetSubscript().toggleSuperscript().run()}><SuperscriptIcon size={16} /></Tool><Tool label="Clear formatting" onClick={() => editor.chain().focus().clearNodes().unsetAllMarks().run()}><RemoveFormatting size={16} /></Tool></div><span className="doc-group-label">Font</span></div>
        <div className="doc-ribbon-group"><div className="doc-tools-row"><Tool label="Bulleted list" active={editorState?.bullet} onClick={() => editor.chain().focus().toggleBulletList().run()}><List size={18} /></Tool><Tool label="Numbered list" active={editorState?.ordered} onClick={() => editor.chain().focus().toggleOrderedList().run()}><ListOrdered size={18} /></Tool><Tool label="Decrease indent" disabled={!editorState?.indent} onClick={() => indentParagraph(-1)}><IndentDecrease size={17} /></Tool><Tool label="Increase indent" disabled={(editorState?.indent || 0) >= 8} onClick={() => indentParagraph(1)}><IndentIncrease size={17} /></Tool><select className="doc-line-select" title="Line spacing" aria-label="Line spacing" value={editorState?.lineHeight} onChange={e => editor.chain().focus().setLineHeight(e.target.value).run()}><option value="1">1.0</option><option value="1.15">1.15</option><option value="1.5">1.5</option><option value="1.65">1.65</option><option value="2">2.0</option></select></div><div className="doc-tools-row">{[{ value: 'left', Icon: AlignLeft }, { value: 'center', Icon: AlignCenter }, { value: 'right', Icon: AlignRight }, { value: 'justify', Icon: AlignJustify }].map(({ value, Icon }) => <Tool key={value} label={`Align ${value}`} active={editorState?.align === value} onClick={() => editor.chain().focus().setTextAlign(value).run()}><Icon size={17} /></Tool>)}</div><span className="doc-group-label">Paragraph</span></div>
        <div className="doc-ribbon-group"><div className="doc-tools-row"><select aria-label="Change case" value="" disabled={!selectedText} onChange={e => changeCase(e.target.value)}><option value="" disabled>Change case</option><option value="upper">UPPERCASE</option><option value="lower">lowercase</option><option value="title">Title Case</option></select></div><div className="doc-tools-row"><Tool label={formatMarks === null ? "Copy formatting" : "Apply formatting"} active={formatMarks !== null} onClick={formatPainter}><Paintbrush size={17} /><span>{formatMarks === null ? "Format painter" : "Apply formatting"}</span></Tool>{formatMarks !== null && <Tool label="Cancel format painter" onClick={() => setFormatMarks(null)}><X size={14} /></Tool>}</div><span className="doc-group-label">Text tools</span></div><div className="doc-ribbon-group doc-styles-group"><div className="doc-style-options">{[{ value: '0', name: 'Normal', sample: 'Aa' }, { value: '1', name: 'Heading 1', sample: 'Aa' }, { value: '2', name: 'Heading 2', sample: 'Aa' }, { value: '3', name: 'Heading 3', sample: 'Aa' }].map(s => <button key={s.value} className={`doc-style-button doc-style-${s.value} ${editorState?.heading === s.value ? 'is-active' : ''}`} onMouseDown={e => e.preventDefault()} onClick={() => s.value === '0' ? editor.chain().focus().setParagraph().run() : editor.chain().focus().setHeading({ level: Number(s.value) as 1 | 2 | 3 }).run()}><strong>{s.sample}</strong><span>{s.name}</span></button>)}</div><span className="doc-group-label">Styles</span></div>
      </>}
      {tab === 'Insert' && <>
        <div className="doc-ribbon-group"><div className="doc-large-tools"><Tool wide label="Table" onClick={() => setModal('table')}><Table2 size={23} /></Tool><Tool wide label="Image" onClick={() => imageInput.current?.click()}><ImagePlus size={23} /></Tool><Tool wide label="Link" onClick={showLink}><Link size={23} /></Tool></div><span className="doc-group-label">Add to your document</span></div>
        <div className="doc-ribbon-group"><div className="doc-large-tools"><Tool wide label="Page break" onClick={() => editor.chain().focus().insertContent([{ type: "pageBreak" }, { type: "paragraph" }]).run()}><FilePlus2 size={23} /></Tool><Tool wide label="Divider" onClick={() => editor.chain().focus().setHorizontalRule().run()}><Minus size={23} /></Tool><Tool wide label="Quote" onClick={() => editor.chain().focus().toggleBlockquote().run()}><Quote size={23} /></Tool><Tool wide label="Date" onClick={() => editor.chain().focus().insertContent(new Date().toLocaleDateString(undefined, { year: 'numeric', month: 'long', day: 'numeric' })).run()}><Type size={23} /></Tool></div><span className="doc-group-label">Elements</span></div>
        {editorState?.table && <><div className="doc-ribbon-group"><div className="doc-tools-row"><Tool label="Add row above" onClick={() => editor.chain().focus().addRowBefore().run()}><Rows3 size={17} /><span>Row above</span></Tool><Tool label="Add row below" onClick={() => editor.chain().focus().addRowAfter().run()}><Rows3 size={17} /><span>Row below</span></Tool><Tool label="Delete row" onClick={() => editor.chain().focus().deleteRow().run()}><Trash2 size={16} /><span>Row</span></Tool></div><div className="doc-tools-row"><Tool label="Add column before" onClick={() => editor.chain().focus().addColumnBefore().run()}><Columns3 size={17} /><span>Column before</span></Tool><Tool label="Add column after" onClick={() => editor.chain().focus().addColumnAfter().run()}><Columns3 size={17} /><span>Column after</span></Tool><Tool label="Delete column" onClick={() => editor.chain().focus().deleteColumn().run()}><Trash2 size={16} /><span>Column</span></Tool></div><span className="doc-group-label">Rows & columns</span></div><div className="doc-ribbon-group"><div className="doc-tools-row"><Tool label="Merge cells" disabled={!editorState.canMerge} onClick={() => editor.chain().focus().mergeCells().run()}><Merge size={17} /><span>Merge</span></Tool><Tool label="Split cell" disabled={!editorState.canSplit} onClick={() => editor.chain().focus().splitCell().run()}><Split size={17} /><span>Split</span></Tool><Tool label="Toggle header row" onClick={() => editor.chain().focus().toggleHeaderRow().run()}><Rows3 size={17} /><span>Header</span></Tool></div><div className="doc-tools-row"><label className="doc-cell-color">Cell fill<input type="color" aria-label="Cell background" defaultValue="#e6efe9" onChange={e => editor.chain().focus().setCellAttribute('backgroundColor', e.target.value).run()} /></label><Tool label="Remove cell fill" onClick={() => editor.chain().focus().setCellAttribute('backgroundColor', null).run()}><RemoveFormatting size={16} /></Tool><Tool label="Delete table" onClick={() => editor.chain().focus().deleteTable().run()}><Trash2 size={16} /><span>Table</span></Tool></div><span className="doc-group-label">Table tools</span></div></>}
        {editorState?.image && <div className="doc-ribbon-group"><div className="doc-tools-row"><label className="doc-image-width">Width <input aria-label="Image width" type="number" min="40" max="1200" value={imageWidthDraft} placeholder="Auto" onChange={e => setImageWidthDraft(e.target.value)} onBlur={() => { const width = imageWidthDraft ? Math.max(40, Math.min(1200, Number(imageWidthDraft) || 40)) : null; editor.commands.updateAttributes('image', { width }); setImageWidthDraft(width ? String(width) : ''); }} onKeyDown={e => { if (e.key === 'Enter') e.currentTarget.blur(); }} /> px</label><Tool label="Reset image width" onClick={() => editor.chain().focus().updateAttributes('image', { width: null }).run()}><RotateCcw size={15} /></Tool></div><div className="doc-tools-row">{[{ value: 'left', Icon: AlignLeft }, { value: 'center', Icon: AlignCenter }, { value: 'right', Icon: AlignRight }].map(({ value, Icon }) => <Tool key={value} label={`Image align ${value}`} active={editorState.imageAttrs.align === value} onClick={() => editor.chain().focus().updateAttributes('image', { align: value }).run()}><Icon size={17} /></Tool>)}<Tool label="Delete image" onClick={() => editor.chain().focus().deleteSelection().run()}><Trash2 size={17} /></Tool></div><span className="doc-group-label">Picture tools</span></div>}
      </>}
      {tab === 'Layout' && <>
        <div className="doc-ribbon-group"><div className="doc-layout-control"><span>Page orientation</span><select aria-label="Page orientation" value={landscape ? 'landscape' : 'portrait'} onChange={e => updatePageSetup({ landscape: e.target.value === 'landscape' })}><option value="portrait">Portrait</option><option value="landscape">Landscape</option></select></div><span className="doc-group-label">Page setup</span></div>
        <div className="doc-ribbon-group"><div className="doc-layout-control"><span>Page margins</span><select aria-label="Page margins" value={margin} onChange={e => updatePageSetup({ margin: e.target.value as Margin })}><option value="normal">Normal · 1 inch</option><option value="narrow">Narrow · 0.5 inch</option><option value="wide">Wide · 1.5 inch</option></select></div><span className="doc-group-label">Spacing</span></div><div className="doc-ribbon-group"><div className="doc-layout-control"><span>Paper size</span><select aria-label="Paper size" value={pageSize} onChange={e => updatePageSetup({ size: e.target.value as DocumentPageSetup["size"] })}><option value="a4">A4 · 210 × 297 mm</option><option value="letter">Letter · 8.5 × 11 in</option><option value="legal">Legal · 8.5 × 14 in</option></select></div><span className="doc-group-label">Paper</span></div><div className="doc-ribbon-group"><div className="doc-large-tools"><Tool wide label="Header & footer" onClick={openHeaderFooter}><BetweenHorizontalStart size={23} /></Tool><Tool wide label="Page numbers" active={!!pageSetup.pageNumbers} onClick={() => updatePageSetup({ pageNumbers: !pageSetup.pageNumbers })}><FileText size={23} /></Tool></div><span className="doc-group-label">Page details</span></div>
      </>}
      {tab === 'Review' && <><div className="doc-ribbon-group"><div className="doc-large-tools"><Tool wide label="Find & replace" onClick={() => setFindOpen(!findOpen)}><Search size={23} /></Tool><Tool wide label="New comment" disabled={reading || !onCommentsChange || !selectedText} onClick={beginComment}><MessageSquare size={23} /></Tool><Tool wide label="Version history" disabled={!onOpenVersions} onClick={() => onOpenVersions?.()}><History size={23} /></Tool><Tool wide label="Spellcheck" onClick={() => { const enabled = editor.view.dom.getAttribute('spellcheck') !== 'false'; editor.view.dom.setAttribute('spellcheck', String(!enabled)); onNotify(`Browser spellcheck ${enabled ? 'disabled' : 'enabled'}.`); }}><CheckCheck size={23} /></Tool></div><span className="doc-group-label">Proofing</span></div><div className="doc-word-stats"><strong>{words.toLocaleString()}<span>words</span></strong><strong>{characters.toLocaleString()}<span>characters</span></strong><strong>{editorState?.text.split('\n').filter(Boolean).length || 0}<span>paragraphs</span></strong></div></>}
      {tab === 'View' && <><div className="doc-ribbon-group"><div className="doc-large-tools"><Tool wide label="Outline" active={outlineOpen} onClick={() => setOutlineOpen(!outlineOpen)}><PanelLeft size={23} /></Tool><Tool wide label={reading ? 'Edit document' : 'Reading mode'} active={reading} onClick={() => setReading(!reading)}><Eye size={23} /></Tool><Tool wide label="100%" onClick={() => setZoom(100)}><Maximize size={23} /></Tool></div><span className="doc-group-label">Workspace</span></div><div className="doc-ribbon-group"><div className="doc-large-tools"><Tool wide label="Print / PDF" onClick={print}><Printer size={23} /></Tool></div><span className="doc-group-label">Print</span></div></>}
    </fieldset>
    {findOpen && <div className="doc-find-panel"><Search size={17} /><input ref={findInput} aria-label="Find text" placeholder="Find in document" value={find} onChange={e => { setFind(e.target.value); setFindIndex(-1); }} onKeyDown={e => { if (e.key === 'Enter') findNext(); }} /><button onClick={findNext} disabled={!find}>Find next</button><span className="doc-find-separator" /><input aria-label="Replace text" placeholder="Replace with" value={replacement} onChange={e => setReplacement(e.target.value)} /><button onClick={replaceSelected} disabled={!find || reading}>Replace</button><button onClick={replaceAll} disabled={!find || reading}>Replace all</button><label className="doc-find-option"><input type="checkbox" checked={matchCase} onChange={e => { setMatchCase(e.target.checked); setFindIndex(-1); }} />Match case</label><label className="doc-find-option"><input type="checkbox" checked={wholeWord} onChange={e => { setWholeWord(e.target.checked); setFindIndex(-1); }} />Whole word</label><small>{find ? `${getMatches().length} matches` : ''}</small><button className="doc-find-close" aria-label="Close find and replace" onClick={() => setFindOpen(false)}><X size={17} /></button></div>}
    <main className="doc-work-area">
      {outlineOpen && <aside className="doc-outline"><div><strong>Document outline</strong><button aria-label="Close outline" onClick={() => setOutlineOpen(false)}><X size={16} /></button></div>{editorState?.headings.length ? editorState.headings.map((h, i) => <button className="doc-outline-item" style={{ paddingLeft: 12 + (h.level - 1) * 12 }} key={i} onClick={() => { const headings = editor.view.dom.querySelectorAll('h1,h2,h3,h4,h5,h6'); headings[i]?.scrollIntoView({ behavior: 'smooth', block: 'center' }); }}>{h.text || 'Untitled heading'}</button>) : <p>Add headings to your document to see its outline here.</p>}</aside>}
      <div className="doc-paper-scroll"><div className="doc-page-caption"><span>{reading ? 'READING VIEW' : 'YOUR WORDS, YOUR SPACE'}</span><span>{pageSize.toUpperCase()} · {landscape ? 'Landscape' : 'Portrait'}</span></div><div className={`doc-sheet doc-margin-${margin} ${landscape ? 'doc-landscape' : ''}`} style={{ zoom: zoom / 100, '--doc-zoom': zoom / 100, '--paper-width': `${paperDimensions[landscape ? 1 : 0]}px`, '--paper-height': `${paperDimensions[landscape ? 0 : 1]}px` } as CSSProperties}>{pageSetup.header && <div className="doc-page-header">{pageSetup.header}</div>}<EditorContent editor={editor} />{(pageSetup.footer || pageSetup.pageNumbers) && <div className="doc-page-footer"><span>{pageSetup.footer}</span>{pageSetup.pageNumbers && <span className="doc-page-number" title="Page numbers update automatically in your Word document">1</span>}</div>}</div><div className="doc-page-end">Made with folio. Made by you.</div></div>
      {commentsOpen && <aside className="doc-comments" aria-label="Comments"><div className="doc-comments-heading"><h3>Comments <span>{comments.filter(comment => !comment.resolved).length}</span></h3><button aria-label="Close comments" onClick={() => setCommentsOpen(false)}><X size={17} /></button></div><button className="doc-new-comment" disabled={reading || !onCommentsChange || !selectedText} onMouseDown={e => e.preventDefault()} onClick={beginComment}><Plus size={15} />New comment</button>{commentRange && <form className="doc-comment-compose" onSubmit={event => { event.preventDefault(); addComment(); }}><blockquote>{commentRange.quote}</blockquote><textarea autoFocus aria-label="Comment text" placeholder="Add your comment…" value={commentDraft} disabled={reading} onChange={event => setCommentDraft(event.target.value)} /><div><button type="button" onClick={() => setCommentRange(null)}>Cancel</button><button type="submit" disabled={reading || !commentDraft.trim()}>Add comment</button></div></form>}<label className="doc-resolved-toggle"><input type="checkbox" checked={showResolved} onChange={event => setShowResolved(event.target.checked)} />Show resolved</label><div className="doc-comment-list">{comments.filter(comment => showResolved || !comment.resolved).map(comment => <article key={comment.id} className={`doc-comment-card ${comment.resolved ? 'is-resolved' : ''}`}><button className="doc-comment-quote" onClick={() => focusComment(comment.id)} title="Go to commented text">“{comment.quote}”</button><p>{comment.text}</p><time>{new Date(comment.createdAt).toLocaleDateString(undefined, { month: 'short', day: 'numeric' })}{comment.resolved ? ' · Resolved' : ''}</time><div className="doc-comment-actions"><button disabled={reading || !onCommentsChange} onClick={() => onCommentsChange?.(comments.map(item => item.id === comment.id ? { ...item, resolved: !item.resolved } : item))}>{comment.resolved ? <RotateCcw size={13} /> : <Check size={13} />}{comment.resolved ? 'Reopen' : 'Resolve'}</button><button aria-label="Delete comment" disabled={reading || !onCommentsChange} onClick={() => deleteComment(comment.id)}><Trash2 size={13} /></button></div></article>)}{!comments.some(comment => showResolved || !comment.resolved) && <p className="doc-comment-empty">{comments.length ? 'All comments are resolved.' : 'Select a passage and add a comment. Keep your notes close to your words.'}</p>}</div></aside>}
    </main>
    <footer className="doc-status"><span className={saveState === "error" ? "doc-save-error" : ""}><span className={`doc-status-dot doc-status-${saveState}`} /> {saveState === "saved" ? "All changes saved" : saveLabel}</span><span>{words.toLocaleString()} words</span><span className="doc-status-local">Private · Stored on this device</span><div className="doc-zoom"><button aria-label="Zoom out" disabled={zoom <= 50} onClick={() => setZoom(Math.max(50, zoom - 10))}><Minus size={14} /></button><input aria-label="Zoom level" type="range" min="50" max="150" step="10" value={zoom} onChange={e => setZoom(Number(e.target.value))} /><button aria-label="Zoom in" disabled={zoom >= 150} onClick={() => setZoom(Math.min(150, zoom + 10))}><Plus size={14} /></button><button className="doc-zoom-value" onClick={() => setZoom(100)}>{zoom}%</button></div></footer>
    <input type="file" ref={imageInput} accept="image/*" hidden onChange={e => importImage(e.target.files?.[0])} />
    {modal && <div className="doc-modal-scrim" onMouseDown={e => { if (e.target === e.currentTarget) setModal(null); }}><section className="doc-modal" role="dialog" aria-modal="true" aria-label={modal === 'link' ? 'Insert link' : modal === 'table' ? 'Insert table' : 'Header and footer'}><button className="doc-modal-close" aria-label="Close dialog" onClick={() => setModal(null)}><X size={19} /></button><div className="doc-modal-icon">{modal === 'link' ? <Link size={25} /> : modal === 'table' ? <Table2 size={25} /> : <BetweenHorizontalStart size={25} />}</div><h2>{modal === 'link' ? 'Add a link' : modal === 'table' ? 'Make room for details' : 'The finishing touches'}</h2><p>{modal === 'link' ? 'Connect your words to something useful.' : modal === 'table' ? 'Insert a table. You can add more rows and columns later.' : 'Repeat a header, footer, and page numbers throughout your Word document.'}</p>{modal === 'link' ? <form onSubmit={e => { e.preventDefault(); insertLink(); }}><label>Web address<input autoFocus value={linkValue} onChange={e => setLinkValue(e.target.value)} placeholder="https://example.com" /></label><div className="doc-modal-actions"><button type="button" onClick={() => setModal(null)}>Cancel</button><button className="doc-primary" type="submit">{linkValue ? 'Apply link' : 'Remove link'}</button></div></form> : modal === 'headerFooter' ? <form className="doc-header-footer-form" onSubmit={e => { e.preventDefault(); updatePageSetup({ header: headerDraft, footer: footerDraft, pageNumbers: numbersDraft }); setModal(null); }}><label>Header<input autoFocus aria-label="Header text" maxLength={300} value={headerDraft} onChange={e => setHeaderDraft(e.target.value)} placeholder="Company or document title" /></label><label>Footer<input aria-label="Footer text" maxLength={300} value={footerDraft} onChange={e => setFooterDraft(e.target.value)} placeholder="Add a footer" /></label><label className="doc-checkbox-label"><input type="checkbox" checked={numbersDraft} onChange={e => setNumbersDraft(e.target.checked)} />Include page numbers</label><div className="doc-modal-actions"><button type="button" onClick={() => setModal(null)}>Cancel</button><button className="doc-primary" type="submit">Apply</button></div></form> : <form onSubmit={e => { e.preventDefault(); editor.chain().focus().insertTable({ rows: tableRows, cols: tableCols, withHeaderRow: true }).run(); setModal(null); }}><div className="doc-table-inputs"><label>Rows<input type="number" min="1" max="30" value={tableRows} onChange={e => setTableRows(Math.min(30, Math.max(1, Number(e.target.value))))} /></label><label>Columns<input type="number" min="1" max="12" value={tableCols} onChange={e => setTableCols(Math.min(12, Math.max(1, Number(e.target.value))))} /></label></div><div className="doc-modal-actions"><button type="button" onClick={() => setModal(null)}>Cancel</button><button className="doc-primary" type="submit">Insert table</button></div></form>}</section></div>}
  </div>;
}
