import { useEffect, useMemo, useRef, useState } from 'react';
import type { CSSProperties, PointerEvent as ReactPointerEvent } from 'react';
import { AlignCenter, AlignLeft, AlignRight, ArrowDown, ArrowLeft, ArrowUp, Bold, Check, ChevronDown, ChevronLeft, ChevronRight, Circle, Copy, Download, EyeOff, History, ImagePlus, Layers, LayoutTemplate, Maximize2, Minus, Move, Plus, Presentation, Redo2, Square, StickyNote, Trash2, Type, Undo2, X } from 'lucide-react';
import type { EditorProps, SlideData, SlideElement } from '../types';
import { exportPresentationPptx, slideAccent, slideLayout, slideTextColor, SLIDE_THEMES } from './slidesPptx';
import './presentation-editor.css';

export { exportPresentationPptx } from './slidesPptx';

function createSlide(): SlideData {
  return { id: crypto.randomUUID(), title: 'Untitled slide', body: 'Add your ideas here.', background: '#ffffff', layout: 'content', elements: [], notes: '', transition: 'none' };
}
const clamp = (value: number, min: number, max: number) => Math.max(min, Math.min(max, value));
const fontSize = (points: number) => `${points / 9.6}cqw`;
const isTextTarget = (target: EventTarget | null) => target instanceof HTMLElement && (target.matches('input,textarea,select') || target.isContentEditable);
const boxStyle = (box: { x: number; y: number; w: number; h: number }): CSSProperties => ({ left: `${box.x}%`, top: `${box.y}%`, width: `${box.w}%`, height: `${box.h}%` });

interface SlidePreviewProps {
  slide: SlideData;
  index: number;
  editable?: boolean;
  selectedElementId?: string | null;
  onSelect?: (id: string | null) => void;
  onEdit?: (patch: Partial<SlideData>, group?: string) => void;
  onElementEdit?: (id: string, patch: Partial<SlideElement>, group?: string) => void;
}
function SlidePreview({ slide, index, editable, selectedElementId, onSelect, onEdit, onElementEdit }: SlidePreviewProps) {
  const paperRef = useRef<HTMLDivElement>(null);
  const dragRef = useRef<{ id: string; x: number; y: number; box: SlideElement; mode: 'move' | 'resize'; rect: DOMRect; patch: Partial<SlideElement> } | null>(null);
  const [dragPreview, setDragPreview] = useState<{ id: string; patch: Partial<SlideElement> } | null>(null);
  const positions = slideLayout(slide);
  const beginDrag = (event: ReactPointerEvent<HTMLElement>, element: SlideElement, mode: 'move' | 'resize') => {
    if (!editable || event.button !== 0) return;
    event.stopPropagation();
    event.preventDefault();
    onSelect?.(element.id);
    const rect = paperRef.current!.getBoundingClientRect();
    dragRef.current = { id: element.id, x: event.clientX, y: event.clientY, box: element, mode, rect, patch: {} };
    event.currentTarget.setPointerCapture(event.pointerId);
  };
  const moveDrag = (event: ReactPointerEvent<HTMLElement>) => {
    const drag = dragRef.current;
    if (!drag) return;
    const dx = (event.clientX - drag.x) / drag.rect.width * 100;
    const dy = (event.clientY - drag.y) / drag.rect.height * 100;
    if (drag.mode === 'move') drag.patch = { x: clamp(drag.box.x + dx, 0, 100 - drag.box.width), y: clamp(drag.box.y + dy, 0, 100 - drag.box.height) };
    else {
      const angle = (drag.box.rotation || 0) * Math.PI / 180;
      const cos = Math.cos(angle); const sin = Math.sin(angle);
      const localDx = (event.clientX - drag.x) * cos + (event.clientY - drag.y) * sin;
      const localDy = -(event.clientX - drag.x) * sin + (event.clientY - drag.y) * cos;
      let width = clamp(drag.box.width + localDx / drag.rect.width * 100, 2, 100 - drag.box.x);
      let height = clamp(drag.box.height + localDy / drag.rect.height * 100, 1, 100 - drag.box.y);
      if (event.shiftKey) {
        const scaleX = (width - drag.box.width) / drag.box.width;
        const scaleY = (height - drag.box.height) / drag.box.height;
        const factor = clamp(1 + (Math.abs(scaleX) > Math.abs(scaleY) ? scaleX : scaleY), Math.max(2 / drag.box.width, 1 / drag.box.height), Math.min((100 - drag.box.x) / drag.box.width, (100 - drag.box.y) / drag.box.height));
        width = drag.box.width * factor;
        height = drag.box.height * factor;
      }
      const dw = (width - drag.box.width) * drag.rect.width / 100;
      const dh = (height - drag.box.height) * drag.rect.height / 100;
      drag.patch = { width, height, x: clamp(drag.box.x + (dw * cos - dh * sin - dw) / 2 / drag.rect.width * 100, 0, 100 - width), y: clamp(drag.box.y + (dw * sin + dh * cos - dh) / 2 / drag.rect.height * 100, 0, 100 - height) };
    }
    setDragPreview({ id: drag.id, patch: drag.patch });
  };
  const finishDrag = (event: ReactPointerEvent<HTMLElement>) => {
    const drag = dragRef.current;
    if (!drag) return;
    event.stopPropagation();
    dragRef.current = null;
    setDragPreview(null);
    if (event.type !== 'pointercancel') onElementEdit?.(drag.id, drag.patch);
  };
  return <div ref={paperRef} className={`slides-paper slides-layout-${slide.layout} ${editable ? 'slides-paper-editable' : ''}`} style={{ backgroundColor: slide.background, color: slideTextColor(slide), fontFamily: slide.fontFamily || 'Aptos, Arial, sans-serif', '--slide-accent': slideAccent(slide) } as CSSProperties} onPointerDown={() => editable && onSelect?.(null)}>
    {slide.layout !== 'blank' && <>
      <div className="slides-paper-mark"><span /><span /></div>
      <div className="slides-paper-title-area" style={boxStyle(positions.title)}>
        {editable ? <textarea aria-label="Slide title" className="slides-title-field" style={{ fontSize: fontSize(slide.titleSize || (slide.layout === 'title' ? 42 : 34)) }} value={slide.title} onFocus={() => onSelect?.(null)} onChange={event => onEdit?.({ title: event.target.value }, 'title')} placeholder="Add a title" spellCheck /> : <h2 style={{ fontSize: fontSize(slide.titleSize || (slide.layout === 'title' ? 42 : 34)) }}>{slide.title}</h2>}
      </div>
      {slide.layout === 'title' && <div className="slides-paper-subtitle-area" style={boxStyle(positions.subtitle)}>{editable ? <textarea aria-label="Slide subtitle" className="slides-subtitle-field" value={slide.subtitle || ''} onFocus={() => onSelect?.(null)} onChange={event => onEdit?.({ subtitle: event.target.value }, 'subtitle')} placeholder="Add a subtitle" /> : <p className="slides-subtitle-text">{slide.subtitle}</p>}</div>}
      <div className="slides-paper-body-area" style={boxStyle(positions.body)}>
        {editable ? <textarea aria-label="Slide body" className="slides-body-field" style={{ fontSize: fontSize(slide.bodySize || 23) }} value={slide.body} onFocus={() => onSelect?.(null)} onChange={event => onEdit?.({ body: event.target.value }, 'body')} placeholder="Add your ideas here" spellCheck /> : <p style={{ fontSize: fontSize(slide.bodySize || 23) }}>{slide.body}</p>}
      </div>
      <span className="slides-paper-number">{String(index + 1).padStart(2, '0')}</span><div className="slides-paper-rule" />
    </>}
    {(slide.elements || []).map(original => {
      const element = dragPreview?.id === original.id ? { ...original, ...dragPreview.patch } : original;
      const selected = editable && selectedElementId === element.id;
      return <div key={element.id} data-element-id={element.id} data-element-type={element.type} className={`slides-element slides-element-${element.type} ${selected ? 'slides-element-selected' : ''}`} style={{ left: `${element.x}%`, top: `${element.y}%`, width: `${element.width}%`, height: `${element.height}%`, transform: `rotate(${element.rotation || 0}deg)`, color: element.color || slideTextColor(slide), background: element.type === 'text' ? element.fill || 'transparent' : undefined, fontSize: fontSize(element.fontSize || 18), fontWeight: element.bold ? 700 : 400, textAlign: element.align || 'left' }} onPointerDown={event => beginDrag(event, element, 'move')} onPointerMove={moveDrag} onPointerUp={finishDrag} onPointerCancel={finishDrag}>
        {element.type === 'text' && (editable ? <textarea aria-label="Text box content" spellCheck value={element.text || ''} onPointerDown={event => { event.stopPropagation(); onSelect?.(element.id); }} onFocus={() => onSelect?.(element.id)} onChange={event => onElementEdit?.(element.id, { text: event.target.value }, `text-${element.id}`)} /> : <div className="slides-element-text-content">{element.text}</div>)}
        {element.type === 'image' && <img src={element.src} alt="Slide image" draggable={false} />}
        {element.type === 'shape' && (element.shape === 'line' ? <svg className="slides-shape-line" viewBox="0 0 100 100" preserveAspectRatio="none"><line x1="0" y1="0" x2="100" y2="100" stroke={element.color || element.fill || slideAccent(slide)} strokeWidth="2" vectorEffect="non-scaling-stroke" /></svg> : <div className={`slides-shape-${element.shape || 'rectangle'}`} style={{ background: element.fill || slideAccent(slide), borderColor: element.color || element.fill || slideAccent(slide) }} />)}
        {selected && <><button className="slides-element-move" aria-label="Move selected object" title="Drag to move" onPointerDown={event => beginDrag(event, element, 'move')}><Move size={14} /></button><button className="slides-element-resize" aria-label="Resize selected object" title="Drag to resize; hold Shift to keep proportions" onPointerDown={event => beginDrag(event, element, 'resize')} /></>}
      </div>;
    })}
  </div>;
}

export default function PresentationEditor({ file, onChange, onRename, onBack, onNotify, onOpenVersions, saveState = 'saved' }: EditorProps) {
  const fallback = useMemo(() => [createSlide()], [file.id]);
  const slides: SlideData[] = Array.isArray(file.content) && file.content.length ? file.content : fallback;
  const [selectedId, setSelectedId] = useState(slides[0].id);
  const [selectedElementId, setSelectedElementId] = useState<string | null>(null);
  const [name, setName] = useState(file.name);
  const [exportOpen, setExportOpen] = useState(false);
  const [presenting, setPresenting] = useState(false);
  const [presentationIndex, setPresentationIndex] = useState(0);
  const [presenterNotes, setPresenterNotes] = useState(false);
  const [notesOpen, setNotesOpen] = useState(true);
  const [exporting, setExporting] = useState(false);
  const [designOpen, setDesignOpen] = useState(() => window.innerWidth > 760);
  const [history, setHistory] = useState<{ past: SlideData[][]; future: SlideData[][] }>({ past: [], future: [] });
  const historyGroup = useRef({ key: '', time: 0 });
  const imageInput = useRef<HTMLInputElement>(null);
  const activeIndex = Math.max(0, slides.findIndex(slide => slide.id === selectedId));
  const activeSlide = slides[activeIndex];
  const selectedElement = activeSlide.elements?.find(element => element.id === selectedElementId);
  const visibleSlides = slides.filter(slide => !slide.hidden);
  const presentedSlide = visibleSlides[Math.min(presentationIndex, visibleSlides.length - 1)];

  useEffect(() => { setName(file.name); }, [file.name]);
  useEffect(() => { setHistory({ past: [], future: [] }); setSelectedElementId(null); historyGroup.current = { key: '', time: 0 }; }, [file.id]);
  const commit = (next: SlideData[], group?: string) => {
    if (JSON.stringify(next) === JSON.stringify(slides)) return;
    const now = Date.now();
    const key = group ? `${activeSlide.id}:${group}` : '';
    const grouped = !!key && historyGroup.current.key === key && now - historyGroup.current.time < 800;
    setHistory(previous => ({ past: grouped ? previous.past : [...previous.past, structuredClone(slides)].slice(-60), future: [] }));
    historyGroup.current = { key, time: now };
    onChange(next);
  };
  const latestState = useRef({ slides, activeSlide, commit });
  latestState.current = { slides, activeSlide, commit };
  const undo = () => {
    if (!history.past.length) return;
    const next = history.past[history.past.length - 1];
    setHistory({ past: history.past.slice(0, -1), future: [structuredClone(slides), ...history.future] });
    historyGroup.current = { key: '', time: 0 };
    onChange(next);
  };
  const redo = () => {
    if (!history.future.length) return;
    setHistory({ past: [...history.past, structuredClone(slides)], future: history.future.slice(1) });
    historyGroup.current = { key: '', time: 0 };
    onChange(history.future[0]);
  };
  const updateSlide = (patch: Partial<SlideData>, group?: string) => commit(slides.map(slide => slide.id === activeSlide.id ? { ...slide, ...patch } : slide), group);
  const updateElement = (id: string, patch: Partial<SlideElement>, group?: string) => updateSlide({ elements: (activeSlide.elements || []).map(element => element.id === id ? { ...element, ...patch } : element) }, group);
  const selectElement = (id: string | null) => { setSelectedElementId(id); if (id && window.innerWidth > 850) setDesignOpen(true); };
  const insertElement = (element: SlideElement) => { updateSlide({ elements: [...(activeSlide.elements || []), element] }); selectElement(element.id); };
  const addText = () => insertElement({ id: crypto.randomUUID(), type: 'text', x: 12, y: 35, width: 40, height: 18, text: 'Your text here', fontSize: 24, color: slideTextColor(activeSlide) });
  const addShape = (shape: 'rectangle' | 'ellipse' | 'line') => insertElement({ id: crypto.randomUUID(), type: 'shape', shape, x: 35, y: 35, width: 25, height: shape === 'line' ? 1 : 28, fill: slideAccent(activeSlide), color: slideAccent(activeSlide) });
  const insertImage = async (fileToRead: File) => {
    const targetSlideId = activeSlide.id;
    if (!/^image\/(png|jpeg|webp|gif)$/.test(fileToRead.type)) { onNotify('Choose a PNG, JPG, WebP, or GIF image.'); return; }
    if (fileToRead.size > 12 * 1024 * 1024) { onNotify('Choose an image smaller than 12 MB.'); return; }
    try {
      const src = await new Promise<string>((resolve, reject) => { const reader = new FileReader(); reader.onload = () => resolve(String(reader.result)); reader.onerror = () => reject(reader.error); reader.readAsDataURL(fileToRead); });
      const dimensions = await new Promise<HTMLImageElement>((resolve, reject) => { const image = new Image(); image.onload = () => resolve(image); image.onerror = reject; image.src = src; });
      // Decode to PNG so exported pictures also open in older PowerPoint versions.
      const canvas = document.createElement('canvas');
      const factor = Math.min(1, 2000 / Math.max(dimensions.naturalWidth, dimensions.naturalHeight));
      canvas.width = Math.max(1, Math.round(dimensions.naturalWidth * factor)); canvas.height = Math.max(1, Math.round(dimensions.naturalHeight * factor));
      canvas.getContext('2d')!.drawImage(dimensions, 0, 0, canvas.width, canvas.height);
      let width = 42; let height = width * 16 / 9 * dimensions.naturalHeight / dimensions.naturalWidth;
      if (height > 68) { width *= 68 / height; height = 68; }
      const element: SlideElement = { id: crypto.randomUUID(), type: 'image', x: (100 - width) / 2, y: (100 - height) / 2, width, height, src: canvas.toDataURL('image/png') };
      const current = latestState.current;
      if (!current.slides.some(slide => slide.id === targetSlideId)) return;
      current.commit(current.slides.map(slide => slide.id === targetSlideId ? { ...slide, elements: [...(slide.elements || []), element] } : slide));
      if (current.activeSlide.id === targetSlideId) selectElement(element.id);
    } catch { onNotify('This image could not be opened. Please choose a different file.'); }
  };
  const deleteElement = () => { if (!selectedElement) return; updateSlide({ elements: activeSlide.elements!.filter(element => element.id !== selectedElement.id) }); setSelectedElementId(null); };
  const duplicateElement = () => { if (!selectedElement) return; insertElement({ ...selectedElement, id: crypto.randomUUID(), x: clamp(selectedElement.x + 3, 0, 100 - selectedElement.width), y: clamp(selectedElement.y + 3, 0, 100 - selectedElement.height) }); };
  const reorderElement = (direction: number) => {
    const next = [...(activeSlide.elements || [])];
    const index = next.findIndex(element => element.id === selectedElementId);
    const destination = index + direction;
    if (index < 0 || destination < 0 || destination >= next.length) return;
    [next[index], next[destination]] = [next[destination], next[index]];
    updateSlide({ elements: next });
  };
  const addSlide = () => {
    const slide = { ...createSlide(), background: activeSlide.background, fontFamily: activeSlide.fontFamily, textColor: activeSlide.textColor, titleSize: activeSlide.titleSize, bodySize: activeSlide.bodySize };
    const next = [...slides]; next.splice(activeIndex + 1, 0, slide); commit(next); setSelectedId(slide.id); setSelectedElementId(null);
  };
  const duplicateSlide = () => {
    const slide = { ...structuredClone(activeSlide), id: crypto.randomUUID(), elements: activeSlide.elements?.map(element => ({ ...element, id: crypto.randomUUID() })) };
    const next = [...slides]; next.splice(activeIndex + 1, 0, slide); commit(next); setSelectedId(slide.id); setSelectedElementId(null); onNotify('Slide duplicated');
  };
  const deleteSlide = () => {
    if (slides.length === 1) return;
    const next = slides.filter(slide => slide.id !== activeSlide.id); setSelectedId(next[Math.min(activeIndex, next.length - 1)].id); setSelectedElementId(null); commit(next);
  };
  const moveSlide = (direction: number) => {
    const destination = activeIndex + direction; if (destination < 0 || destination >= slides.length) return;
    const next = [...slides]; [next[activeIndex], next[destination]] = [next[destination], next[activeIndex]]; commit(next);
  };
  const commitName = () => { const value = name.trim() || 'Untitled presentation'; setName(value); if (value !== file.name) onRename(value); };
  const startPresenting = () => {
    if (!visibleSlides.length) { onNotify('Unhide at least one slide to start the presentation.'); return; }
    setPresentationIndex(Math.max(0, visibleSlides.findIndex(slide => slide.id === activeSlide.id))); setPresenting(true);
  };
  useEffect(() => {
    const handler = (event: KeyboardEvent) => {
      if (document.querySelector('.modal-backdrop')) return;
      if (presenting) {
        if (event.key === 'Escape') setPresenting(false);
        if (['ArrowRight', 'ArrowDown', ' ', 'PageDown'].includes(event.key)) { event.preventDefault(); setPresentationIndex(index => Math.min(visibleSlides.length - 1, index + 1)); }
        if (['ArrowLeft', 'ArrowUp', 'PageUp'].includes(event.key)) { event.preventDefault(); setPresentationIndex(index => Math.max(0, index - 1)); }
        if (event.key === 'Home') setPresentationIndex(0);
        if (event.key === 'End') setPresentationIndex(visibleSlides.length - 1);
        if (event.key.toLowerCase() === 'n') setPresenterNotes(current => !current);
        return;
      }
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'z') { event.preventDefault(); event.shiftKey ? redo() : undo(); return; }
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'y') { event.preventDefault(); redo(); return; }
      if (isTextTarget(event.target)) return;
      if (event.key === 'Escape') setSelectedElementId(null);
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'd') { event.preventDefault(); selectedElement ? duplicateElement() : duplicateSlide(); }
      if (selectedElement && ['Delete', 'Backspace'].includes(event.key)) { event.preventDefault(); deleteElement(); }
      if (selectedElement && ['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(event.key)) {
        event.preventDefault(); const step = event.shiftKey ? 5 : .5;
        updateElement(selectedElement.id, { x: clamp(selectedElement.x + (event.key === 'ArrowLeft' ? -step : event.key === 'ArrowRight' ? step : 0), 0, 100 - selectedElement.width), y: clamp(selectedElement.y + (event.key === 'ArrowUp' ? -step : event.key === 'ArrowDown' ? step : 0), 0, 100 - selectedElement.height) }, 'nudge');
      }
    };
    document.addEventListener('keydown', handler); return () => document.removeEventListener('keydown', handler);
  });
  const exportPowerPoint = async () => {
    setExporting(true); setExportOpen(false);
    try { await exportPresentationPptx(slides, file.name); onNotify('PowerPoint presentation exported'); }
    catch (error) { console.error(error); onNotify('The presentation could not be exported. Please try again.'); }
    finally { setExporting(false); }
  };

  return <div className="slides-editor">
    <header className="slides-header">
      <button className="slides-icon-button slides-back" title="Back to workspace" aria-label="Back to workspace" onClick={onBack}><ArrowLeft size={19} /></button>
      <div className="slides-app-icon"><Presentation size={23} /></div>
      <div className="slides-file-info"><input aria-label="Presentation name" value={name} onChange={event => setName(event.target.value)} onBlur={commitName} onKeyDown={event => { if (event.key === 'Enter') event.currentTarget.blur(); }} /><span className={saveState === 'error' ? 'slides-save-error' : ''} role="status">{saveState === 'saved' && <Check size={12} />}{saveState === 'saving' ? 'Saving…' : saveState === 'error' ? 'Could not save · export a copy' : 'Saved on this device'}</span></div>
      <div className="slides-header-actions">
        {onOpenVersions && <button className="slides-icon-button" aria-label="Version history" title="Version history" onClick={onOpenVersions}><History size={17} /></button>}
        <button className="slides-button" onClick={startPresenting}><Maximize2 size={15} /><span>Present</span></button>
        <div className="slides-export-wrap"><button className="slides-button slides-button-primary" disabled={exporting} onClick={() => setExportOpen(!exportOpen)}><Download size={15} /><span>{exporting ? 'Exporting…' : 'Export'}</span><ChevronDown size={14} /></button>
          {exportOpen && <><button className="slides-menu-backdrop" aria-label="Close export menu" onClick={() => setExportOpen(false)} /><div className="slides-export-menu"><p>DOWNLOAD PRESENTATION</p><button onClick={exportPowerPoint}><Presentation size={17} /><span>PowerPoint <small>.pptx</small></span></button><button onClick={() => { setExportOpen(false); window.print(); }}><Download size={17} /><span>Print / Save as PDF <small>.pdf</small></span></button></div></>}
        </div>
      </div>
    </header>
    <div className="slides-toolbar">
      <div className="slides-toolbar-group"><button className="slides-icon-button" title="Undo (Ctrl+Z)" aria-label="Undo" disabled={!history.past.length} onClick={undo}><Undo2 size={17} /></button><button className="slides-icon-button" title="Redo (Ctrl+Shift+Z)" aria-label="Redo" disabled={!history.future.length} onClick={redo}><Redo2 size={17} /></button><span className="slides-toolbar-divider" /><button className="slides-button slides-new-slide" onClick={addSlide}><Plus size={16} /> New slide</button><button className="slides-icon-button" title="Duplicate slide" aria-label="Duplicate slide" onClick={duplicateSlide}><Copy size={17} /></button><button className="slides-icon-button" title="Delete slide" aria-label="Delete slide" disabled={slides.length === 1} onClick={deleteSlide}><Trash2 size={17} /></button><button className="slides-icon-button" title="Move slide up" aria-label="Move slide up" disabled={activeIndex === 0} onClick={() => moveSlide(-1)}><ArrowUp size={17} /></button><button className="slides-icon-button" title="Move slide down" aria-label="Move slide down" disabled={activeIndex === slides.length - 1} onClick={() => moveSlide(1)}><ArrowDown size={17} /></button></div>
      <button className={`slides-button slides-design-toggle ${designOpen ? 'slides-button-active' : ''}`} onClick={() => setDesignOpen(!designOpen)}><LayoutTemplate size={16} /> Design</button>
    </div>
    <div className="slides-insert-toolbar" aria-label="Insert objects">
      <span>INSERT</span><button className="slides-button" onClick={addText}><Type size={15} /> Text box</button><button className="slides-button" onClick={() => imageInput.current?.click()}><ImagePlus size={15} /> Picture</button><button className="slides-button" onClick={() => addShape('rectangle')}><Square size={14} /> Rectangle</button><button className="slides-button" onClick={() => addShape('ellipse')}><Circle size={14} /> Ellipse</button><button className="slides-button" onClick={() => addShape('line')}><Minus size={15} /> Line</button><input ref={imageInput} type="file" accept="image/png,image/jpeg,image/webp,image/gif" aria-label="Insert slide picture" hidden onChange={event => { const picked = event.target.files?.[0]; if (picked) void insertImage(picked); event.target.value = ''; }} />
      <button className={`slides-button slides-notes-toggle ${notesOpen ? 'slides-button-active' : ''}`} aria-pressed={notesOpen} onClick={() => setNotesOpen(!notesOpen)}><StickyNote size={15} /> Notes</button>
    </div>
    <div className={`slides-workspace ${designOpen ? '' : 'slides-design-hidden'}`}>
      <aside className="slides-filmstrip" aria-label="Slides">
        <div className="slides-filmstrip-heading">SLIDES <span>{slides.length}</span></div>
        <div className="slides-thumbnail-list">{slides.map((slide, index) => <button key={slide.id} onClick={() => { setSelectedId(slide.id); setSelectedElementId(null); }} className={`slides-thumbnail ${activeSlide.id === slide.id ? 'slides-thumbnail-selected' : ''} ${slide.hidden ? 'slides-thumbnail-hidden' : ''}`} aria-label={`Select slide ${index + 1}: ${slide.title}${slide.hidden ? ' (hidden)' : ''}`} aria-pressed={activeSlide.id === slide.id}><span className="slides-thumbnail-number">{index + 1}</span><div className="slides-thumbnail-paper"><SlidePreview slide={slide} index={index} /></div><span className="slides-thumbnail-label">{slide.hidden && <EyeOff size={10} />} {slide.title || 'Untitled slide'}</span></button>)}</div>
        <button className="slides-add-thumbnail" onClick={addSlide}><Plus size={16} /> Add slide</button>
      </aside>
      <main className="slides-stage">
        <div className="slides-stage-heading"><span>Slide {activeIndex + 1}{activeSlide.hidden ? ' · Hidden during presentation' : ''}</span><span>{selectedElement ? 'Drag handle to move · corner to resize' : 'Click text to edit · insert objects above'}</span></div>
        <div className="slides-canvas"><SlidePreview slide={activeSlide} index={activeIndex} editable selectedElementId={selectedElementId} onSelect={selectElement} onEdit={updateSlide} onElementEdit={updateElement} /></div>
        <div className="slides-stage-footer"><span><Presentation size={14} /> Widescreen · 16:9</span><span>{slides.length} {slides.length === 1 ? 'slide' : 'slides'} · {activeSlide.elements?.length || 0} objects</span></div>
        {notesOpen && <label className="slides-speaker-notes"><span><StickyNote size={14} /> Speaker notes <small>Visible in presenter notes and PowerPoint</small></span><textarea aria-label="Speaker notes" placeholder="Add reminders, talking points, or your script…" value={activeSlide.notes || ''} onChange={event => updateSlide({ notes: event.target.value }, 'notes')} /></label>}
      </main>
      {designOpen && <aside className="slides-design-panel"><div className="slides-design-title"><h3>{selectedElement ? 'Format object' : 'Slide design'}</h3><button className="slides-icon-button" aria-label="Close design panel" onClick={() => setDesignOpen(false)}><X size={16} /></button></div>
        {selectedElement && <section className="slides-object-controls"><h4>{selectedElement.type.toUpperCase()} OBJECT</h4><div className="slides-object-actions"><button className="slides-icon-button" aria-label="Duplicate selected object" title="Duplicate object (Ctrl+D)" onClick={duplicateElement}><Copy size={16} /></button><button className="slides-icon-button" aria-label="Delete selected object" title="Delete object" onClick={deleteElement}><Trash2 size={16} /></button><button className="slides-icon-button" aria-label="Send backward" title="Send backward" disabled={activeSlide.elements?.[0]?.id === selectedElement.id} onClick={() => reorderElement(-1)}><ArrowDown size={16} /></button><button className="slides-icon-button" aria-label="Bring forward" title="Bring forward" disabled={activeSlide.elements?.at(-1)?.id === selectedElement.id} onClick={() => reorderElement(1)}><ArrowUp size={16} /></button></div>
          {selectedElement.type === 'text' && <><label className="slides-property">Font size <input aria-label="Object font size" type="number" min="6" max="144" value={selectedElement.fontSize || 18} onChange={event => updateElement(selectedElement.id, { fontSize: clamp(Number(event.target.value) || 18, 6, 144) })} /></label><div className="slides-format-row"><button className={`slides-icon-button ${selectedElement.bold ? 'slides-button-active' : ''}`} aria-label="Bold object text" aria-pressed={!!selectedElement.bold} onClick={() => updateElement(selectedElement.id, { bold: !selectedElement.bold })}><Bold size={16} /></button>{(['left', 'center', 'right'] as const).map((align, index) => { const Icon = [AlignLeft, AlignCenter, AlignRight][index]; return <button key={align} className={`slides-icon-button ${(selectedElement.align || 'left') === align ? 'slides-button-active' : ''}`} aria-label={`Align object ${align}`} aria-pressed={(selectedElement.align || 'left') === align} onClick={() => updateElement(selectedElement.id, { align })}><Icon size={16} /></button>; })}</div></>}
          {selectedElement.type !== 'image' && <><label className="slides-property">{selectedElement.type === 'text' ? 'Text color' : 'Outline / line'}<input type="color" aria-label="Object color" value={selectedElement.color || slideTextColor(activeSlide)} onChange={event => updateElement(selectedElement.id, { color: event.target.value }, 'object-color')} /></label>{selectedElement.shape !== 'line' && <><label className="slides-property">Fill<input type="color" aria-label="Object fill" value={selectedElement.fill && selectedElement.fill !== 'transparent' ? selectedElement.fill : '#ffffff'} onChange={event => updateElement(selectedElement.id, { fill: event.target.value }, 'object-fill')} /></label><label className="slides-check-property"><input type="checkbox" checked={selectedElement.fill === 'transparent' || (selectedElement.type === 'text' && !selectedElement.fill)} onChange={event => updateElement(selectedElement.id, { fill: event.target.checked ? 'transparent' : '#ffffff' })} /> No fill</label></>}</>}
          <div className="slides-geometry-grid">{(['x', 'y', 'width', 'height'] as const).map(key => <label key={key}>{({ x: 'X (%)', y: 'Y (%)', width: 'Width (%)', height: 'Height (%)' })[key]}<input aria-label={`Object ${key}`} type="number" step=".5" min={key === 'width' || key === 'height' ? 1 : 0} max="100" value={Math.round(selectedElement[key] * 10) / 10} onChange={event => { const value = Number(event.target.value); const patch: Partial<SlideElement> = key === 'x' ? { x: clamp(value, 0, 100 - selectedElement.width) } : key === 'y' ? { y: clamp(value, 0, 100 - selectedElement.height) } : key === 'width' ? { width: clamp(value, 1, 100 - selectedElement.x) } : { height: clamp(value, 1, 100 - selectedElement.y) }; updateElement(selectedElement.id, patch); }} /></label>)}</div>
          <label className="slides-property">Rotation<input aria-label="Object rotation" type="number" min="-360" max="360" value={selectedElement.rotation || 0} onChange={event => updateElement(selectedElement.id, { rotation: clamp(Number(event.target.value), -360, 360) })} /></label>
          <button className="slides-apply-theme" onClick={() => setSelectedElementId(null)}>Done formatting object</button>
        </section>}
        <section><h4>LAYOUT</h4><div className="slides-layout-options">{(['title', 'content', 'split', 'blank'] as const).map(layout => <button key={layout} aria-label={`${layout === 'split' ? 'Two columns' : layout[0].toUpperCase() + layout.slice(1)} layout`} className={`slides-layout-option ${activeSlide.layout === layout ? 'slides-option-selected' : ''}`} onClick={() => updateSlide({ layout })}><div className={`slides-layout-mini slides-layout-mini-${layout}`}>{layout !== 'blank' && <><i /><b /><b /></>}</div><span>{layout === 'title' ? 'Title' : layout === 'content' ? 'Content' : layout === 'split' ? 'Two columns' : 'Blank'}</span></button>)}</div></section>
        <section><h4>SLIDE TEXT</h4><label className="slides-property slides-property-stack">Font family<select aria-label="Slide font family" value={activeSlide.fontFamily || 'Aptos'} onChange={event => updateSlide({ fontFamily: event.target.value })}>{['Aptos', 'Arial', 'Georgia', 'Times New Roman', 'Verdana', 'Courier New'].map(font => <option key={font} value={font}>{font}</option>)}</select></label><label className="slides-property">Title size<input type="number" aria-label="Slide title size" min="8" max="96" value={activeSlide.titleSize || (activeSlide.layout === 'title' ? 42 : 34)} onChange={event => updateSlide({ titleSize: clamp(Number(event.target.value) || 34, 8, 96) })} /></label><label className="slides-property">Body size<input type="number" aria-label="Slide body size" min="8" max="72" value={activeSlide.bodySize || 23} onChange={event => updateSlide({ bodySize: clamp(Number(event.target.value) || 23, 8, 72) })} /></label><label className="slides-property">Text color<input type="color" aria-label="Slide text color" value={slideTextColor(activeSlide)} onChange={event => updateSlide({ textColor: event.target.value }, 'slide-color')} /></label></section>
        <section><h4>COLOR THEME</h4><div className="slides-theme-options">{SLIDE_THEMES.map(theme => <button key={theme.name} aria-label={theme.name} aria-pressed={activeSlide.background === theme.background} className={`slides-theme-option ${activeSlide.background === theme.background ? 'slides-option-selected' : ''}`} onClick={() => updateSlide({ background: theme.background })}><div className="slides-theme-swatch" style={{ background: theme.background, color: theme.accent }}><span>Aa</span>{activeSlide.background === theme.background && <Check size={13} />}</div><span>{theme.name}</span></button>)}</div><label className="slides-property slides-background-color">Background<input type="color" aria-label="Slide background color" value={activeSlide.background} onChange={event => updateSlide({ background: event.target.value }, 'background')} /></label><button className="slides-apply-theme" onClick={() => { commit(slides.map(slide => ({ ...slide, background: activeSlide.background }))); onNotify('Background applied to all slides'); }}>Apply background to all slides</button></section>
        <section><h4>SLIDE SHOW</h4><label className="slides-property slides-property-stack">Transition<select aria-label="Slide transition" value={activeSlide.transition || 'none'} onChange={event => updateSlide({ transition: event.target.value as SlideData['transition'] })}><option value="none">None</option><option value="fade">Fade</option><option value="slide">Push</option></select></label><label className="slides-check-property"><input type="checkbox" checked={!!activeSlide.hidden} onChange={event => updateSlide({ hidden: event.target.checked })} /> Hide during presentation</label></section>
        {!!activeSlide.elements?.length && <section><h4><Layers size={12} /> OBJECTS · TOP FIRST</h4><div className="slides-layer-list">{[...activeSlide.elements].reverse().map((element, index) => <button key={element.id} className={element.id === selectedElementId ? 'slides-layer-selected' : ''} onClick={() => selectElement(element.id)}><span>{element.type === 'text' ? element.text?.slice(0, 24) || 'Text box' : element.type === 'image' ? 'Picture' : element.shape || 'Shape'}</span><small>{activeSlide.elements!.length - index}</small></button>)}</div></section>}
      </aside>}
    </div>
    <div className="slides-print-pages">{slides.filter(slide => !slide.hidden).map(slide => <div className="slides-print-page" key={slide.id}><SlidePreview slide={slide} index={slides.findIndex(item => item.id === slide.id)} /></div>)}</div>
    {presenting && presentedSlide && <div className={`slides-presenter ${presenterNotes ? 'slides-presenter-with-notes' : ''}`} role="dialog" aria-modal="true" aria-label="Presentation"><button className="slides-presenter-notes-toggle" onClick={() => setPresenterNotes(!presenterNotes)}><StickyNote size={16} /> {presenterNotes ? 'Hide' : 'Show'} speaker notes</button><button className="slides-presenter-close" onClick={() => setPresenting(false)}><X size={20} /> Exit presentation</button><div key={presentedSlide.id} className={`slides-presenter-screen slides-transition-${presentedSlide.transition || 'none'}`}><SlidePreview slide={presentedSlide} index={slides.findIndex(slide => slide.id === presentedSlide.id)} /></div>{presenterNotes && <div className="slides-presenter-notes" aria-label="Presenter notes"><strong>Speaker notes · visible on this screen</strong><p>{presentedSlide.notes || 'No notes for this slide.'}</p></div>}<div className="slides-presenter-controls"><button aria-label="Previous slide" disabled={presentationIndex === 0} onClick={() => setPresentationIndex(index => Math.max(0, index - 1))}><ChevronLeft size={22} /></button><span>{presentationIndex + 1} / {visibleSlides.length}</span><button aria-label="Next slide" disabled={presentationIndex === visibleSlides.length - 1} onClick={() => setPresentationIndex(index => Math.min(visibleSlides.length - 1, index + 1))}><ChevronRight size={22} /></button></div><span className="slides-presenter-hint">Use ← → to navigate · N for notes · Esc to exit</span></div>}
  </div>;
}
