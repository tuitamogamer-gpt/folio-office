import { useEffect, useMemo, useState } from 'react';
import { ArrowDown, ArrowLeft, ArrowUp, Check, ChevronDown, ChevronLeft, ChevronRight, Copy, Download, LayoutTemplate, Maximize2, Plus, Presentation, Trash2, X } from 'lucide-react';
import type { EditorProps, SlideData } from '../types';
import './presentation-editor.css';

const THEMES = [
  { name: 'Paper', background: '#ffffff', accent: '#2d6554' },
  { name: 'Forest', background: '#244e40', accent: '#d4e6bb' },
  { name: 'Sage', background: '#e6ece2', accent: '#355544' },
  { name: 'Sand', background: '#f4eadd', accent: '#8d5e3d' },
  { name: 'Ink', background: '#25374b', accent: '#bdd5ec' },
  { name: 'Lavender', background: '#ece7f5', accent: '#6f5689' },
];

function createSlide(): SlideData {
  return { id: crypto.randomUUID(), title: 'Untitled slide', body: 'Add your ideas here.', background: '#ffffff', layout: 'content' };
}

function isDark(hex: string) {
  const value = hex.replace('#', '');
  if (value.length !== 6) return false;
  return parseInt(value.slice(0, 2), 16) * .299 + parseInt(value.slice(2, 4), 16) * .587 + parseInt(value.slice(4, 6), 16) * .114 < 140;
}

function SlidePreview({ slide, index, editable, onEdit }: { slide: SlideData; index: number; editable?: boolean; onEdit?: (patch: Partial<SlideData>) => void }) {
  const dark = isDark(slide.background);
  const accent = THEMES.find(theme => theme.background === slide.background)?.accent || (dark ? '#d4e6bb' : '#2d6554');
  return <div className={`slides-paper slides-layout-${slide.layout} ${dark ? 'slides-paper-dark' : ''}`} style={{ backgroundColor: slide.background, '--slide-accent': accent } as React.CSSProperties}>
    <div className="slides-paper-mark"><span /><span /></div>
    <div className="slides-paper-title-area">
      {editable ? <textarea aria-label="Slide title" className="slides-title-field" value={slide.title} onChange={event => onEdit?.({ title: event.target.value })} placeholder="Add a title" spellCheck /> : <h2>{slide.title}</h2>}
      {slide.layout === 'title' && (editable ? <textarea aria-label="Slide subtitle" className="slides-subtitle-field" value={slide.subtitle || ''} onChange={event => onEdit?.({ subtitle: event.target.value })} placeholder="Add a subtitle" /> : slide.subtitle && <p className="slides-subtitle-text">{slide.subtitle}</p>)}
    </div>
    <div className="slides-paper-body-area">
      {editable ? <textarea aria-label="Slide body" className="slides-body-field" value={slide.body} onChange={event => onEdit?.({ body: event.target.value })} placeholder="Add your ideas here" spellCheck /> : <p>{slide.body}</p>}
    </div>
    <span className="slides-paper-number">{String(index + 1).padStart(2, '0')}</span>
    <div className="slides-paper-rule" />
  </div>;
}


export async function exportPresentationPptx(slides: SlideData[], name: string): Promise<void> {
  const { default: PptxGenJS } = await import('pptxgenjs');
  const presentation = new PptxGenJS();
  presentation.layout = 'LAYOUT_WIDE';
  presentation.author = 'Folio';
  presentation.subject = name;
  presentation.title = name;
  presentation.company = 'Folio';
  presentation.theme = { headFontFace: 'Aptos Display', bodyFontFace: 'Aptos' };
  slides.forEach((data, index) => {
    const slide = presentation.addSlide();
    const dark = isDark(data.background);
    const color = dark ? 'F6F7F3' : '253A34';
    const accent = (THEMES.find(theme => theme.background === data.background)?.accent || '#2d6554').replace('#', '');
    slide.background = { color: data.background.replace('#', '') };
    slide.addShape(presentation.ShapeType.rect, { x: .8, y: .65, w: .23, h: .23, fill: { color: accent }, line: { color: accent } });
    const isTitle = data.layout === 'title';
    const isSplit = data.layout === 'split';
    slide.addText(data.title, { objectName: `folio-title-${data.layout}`, x: .8, y: isTitle ? 1.65 : 1.2, w: isSplit ? 5.25 : 11.7, h: isTitle ? 1.7 : isSplit ? 3.5 : 1.2, fontFace: 'Aptos Display', fontSize: isTitle ? 42 : 34, bold: true, color, breakLine: false, margin: 0, fit: 'shrink', valign: isSplit ? 'middle' : 'top', align: isTitle ? 'center' : 'left' });
    if (isTitle && data.subtitle) slide.addText(data.subtitle, { objectName: 'folio-subtitle', x: 1, y: 3.45, w: 11.3, h: .7, fontSize: 22, color, align: 'center', margin: 0, fit: 'shrink' });
    slide.addText(data.body, { objectName: 'folio-body', x: isSplit ? 7.2 : 1, y: isTitle ? 4.45 : isSplit ? 1.6 : 2.7, w: isSplit ? 5.1 : 11.3, h: isTitle ? 1.4 : 3.7, fontSize: 23, color, margin: 0, breakLine: false, fit: 'shrink', align: isTitle ? 'center' : 'left', valign: 'top', paraSpaceAfter: 12 });
    slide.addShape(presentation.ShapeType.line, { x: .8, y: 6.8, w: 11.7, h: 0, line: { color: accent, width: 1 } });
    slide.addText(String(index + 1).padStart(2, '0'), { objectName: 'folio-page-number', x: 11.8, y: 6.98, w: .6, h: .2, fontSize: 10, color, align: 'right', margin: 0 });
  });
  await presentation.writeFile({ fileName: `${name.replace(/[<>:"/\\|?*]/g, '-')}.pptx` });
}

export default function PresentationEditor({ file, onChange, onRename, onBack, onNotify }: EditorProps) {
  const fallback = useMemo(() => [createSlide()], [file.id]);
  const slides: SlideData[] = Array.isArray(file.content) && file.content.length ? file.content : fallback;
  const [selectedId, setSelectedId] = useState(slides[0].id);
  const [name, setName] = useState(file.name);
  const [exportOpen, setExportOpen] = useState(false);
  const [presenting, setPresenting] = useState(false);
  const [presentationIndex, setPresentationIndex] = useState(0);
  const [exporting, setExporting] = useState(false);
  const [designOpen, setDesignOpen] = useState(() => window.innerWidth > 760);
  const activeIndex = Math.max(0, slides.findIndex(slide => slide.id === selectedId));
  const activeSlide = slides[activeIndex];

  useEffect(() => { setName(file.name); }, [file.name]);
  useEffect(() => {
    if (!presenting) return;
    const handler = (event: KeyboardEvent) => {
      if (event.key === 'Escape') setPresenting(false);
      if (['ArrowRight', 'ArrowDown', ' ', 'PageDown'].includes(event.key)) {
        event.preventDefault();
        setPresentationIndex(index => Math.min(slides.length - 1, index + 1));
      }
      if (['ArrowLeft', 'ArrowUp', 'PageUp'].includes(event.key)) {
        event.preventDefault();
        setPresentationIndex(index => Math.max(0, index - 1));
      }
      if (event.key === 'Home') setPresentationIndex(0);
      if (event.key === 'End') setPresentationIndex(slides.length - 1);
    };
    document.addEventListener('keydown', handler);
    return () => document.removeEventListener('keydown', handler);
  }, [presenting, slides.length]);

  const updateSlide = (patch: Partial<SlideData>) => onChange(slides.map(slide => slide.id === activeSlide.id ? { ...slide, ...patch } : slide));
  const addSlide = () => {
    const slide = { ...createSlide(), background: activeSlide.background };
    const next = [...slides];
    next.splice(activeIndex + 1, 0, slide);
    onChange(next);
    setSelectedId(slide.id);
  };
  const duplicateSlide = () => {
    const slide = { ...activeSlide, id: crypto.randomUUID() };
    const next = [...slides];
    next.splice(activeIndex + 1, 0, slide);
    onChange(next);
    setSelectedId(slide.id);
    onNotify('Slide duplicated');
  };
  const deleteSlide = () => {
    if (slides.length === 1) return;
    const next = slides.filter(slide => slide.id !== activeSlide.id);
    setSelectedId(next[Math.min(activeIndex, next.length - 1)].id);
    onChange(next);
  };
  const moveSlide = (direction: number) => {
    const destination = activeIndex + direction;
    if (destination < 0 || destination >= slides.length) return;
    const next = [...slides];
    [next[activeIndex], next[destination]] = [next[destination], next[activeIndex]];
    onChange(next);
  };
  const commitName = () => {
    const value = name.trim() || 'Untitled presentation';
    setName(value);
    if (value !== file.name) onRename(value);
  };
  const exportPowerPoint = async () => {
    setExporting(true);
    setExportOpen(false);
    try {
      await exportPresentationPptx(slides, file.name);
      onNotify('PowerPoint presentation exported');
    } catch (error) {
      console.error(error);
      onNotify('The presentation could not be exported. Please try again.');
    } finally {
      setExporting(false);
    }
  };

  return <div className="slides-editor">
    <header className="slides-header">
      <button className="slides-icon-button slides-back" title="Back to workspace" aria-label="Back to workspace" onClick={onBack}><ArrowLeft size={19} /></button>
      <div className="slides-app-icon"><Presentation size={23} /></div>
      <div className="slides-file-info"><input aria-label="Presentation name" value={name} onChange={event => setName(event.target.value)} onBlur={commitName} onKeyDown={event => { if (event.key === 'Enter') event.currentTarget.blur(); }} /><span><Check size={12} /> Saved on this device</span></div>
      <div className="slides-header-actions">
        <button className="slides-button" onClick={() => { setPresentationIndex(activeIndex); setPresenting(true); }}><Maximize2 size={15} /><span>Present</span></button>
        <div className="slides-export-wrap"><button className="slides-button slides-button-primary" disabled={exporting} onClick={() => setExportOpen(!exportOpen)}><Download size={15} /><span>{exporting ? 'Exporting…' : 'Export'}</span><ChevronDown size={14} /></button>
          {exportOpen && <><button className="slides-menu-backdrop" aria-label="Close export menu" onClick={() => setExportOpen(false)} /><div className="slides-export-menu"><p>DOWNLOAD PRESENTATION</p><button onClick={exportPowerPoint}><Presentation size={17} /><span>PowerPoint <small>.pptx</small></span></button><button onClick={() => { setExportOpen(false); window.print(); }}><Download size={17} /><span>Print / Save as PDF <small>.pdf</small></span></button></div></>}
        </div>
      </div>
    </header>
    <div className="slides-toolbar">
      <div className="slides-toolbar-group"><button className="slides-button slides-new-slide" onClick={addSlide}><Plus size={16} /> New slide</button><span className="slides-toolbar-divider" /><button className="slides-icon-button" title="Duplicate slide" aria-label="Duplicate slide" onClick={duplicateSlide}><Copy size={17} /></button><button className="slides-icon-button" title="Delete slide" aria-label="Delete slide" disabled={slides.length === 1} onClick={deleteSlide}><Trash2 size={17} /></button><span className="slides-toolbar-divider" /><button className="slides-icon-button" title="Move slide up" aria-label="Move slide up" disabled={activeIndex === 0} onClick={() => moveSlide(-1)}><ArrowUp size={17} /></button><button className="slides-icon-button" title="Move slide down" aria-label="Move slide down" disabled={activeIndex === slides.length - 1} onClick={() => moveSlide(1)}><ArrowDown size={17} /></button></div>
      <button className={`slides-button slides-design-toggle ${designOpen ? 'slides-button-active' : ''}`} onClick={() => setDesignOpen(!designOpen)}><LayoutTemplate size={16} /> Design</button>
    </div>
    <div className={`slides-workspace ${designOpen ? '' : 'slides-design-hidden'}`}>
      <aside className="slides-filmstrip" aria-label="Slides">
        <div className="slides-filmstrip-heading">SLIDES <span>{slides.length}</span></div>
        <div className="slides-thumbnail-list">{slides.map((slide, index) => <button key={slide.id} onClick={() => setSelectedId(slide.id)} className={`slides-thumbnail ${activeSlide.id === slide.id ? 'slides-thumbnail-selected' : ''}`} aria-label={`Select slide ${index + 1}: ${slide.title}`} aria-pressed={activeSlide.id === slide.id}><span className="slides-thumbnail-number">{index + 1}</span><div className="slides-thumbnail-paper"><SlidePreview slide={slide} index={index} /></div><span className="slides-thumbnail-label">{slide.title || 'Untitled slide'}</span></button>)}</div>
        <button className="slides-add-thumbnail" onClick={addSlide}><Plus size={16} /> Add slide</button>
      </aside>
      <main className="slides-stage">
        <div className="slides-stage-heading"><span>Slide {activeIndex + 1}</span><span>Click any text to edit</span></div>
        <div className="slides-canvas"><SlidePreview slide={activeSlide} index={activeIndex} editable onEdit={updateSlide} /></div>
        <div className="slides-stage-footer"><span><Presentation size={14} /> Widescreen · 16:9</span><span>{slides.length} {slides.length === 1 ? 'slide' : 'slides'}</span></div>
      </main>
      {designOpen && <aside className="slides-design-panel"><div className="slides-design-title"><h3>Slide design</h3><button className="slides-icon-button" aria-label="Close design panel" onClick={() => setDesignOpen(false)}><X size={16} /></button></div><section><h4>LAYOUT</h4><div className="slides-layout-options">{(['title', 'content', 'split'] as const).map(layout => <button key={layout} className={`slides-layout-option ${activeSlide.layout === layout ? 'slides-option-selected' : ''}`} onClick={() => updateSlide({ layout })}><div className={`slides-layout-mini slides-layout-mini-${layout}`}><i /><b /><b /></div><span>{layout === 'title' ? 'Title' : layout === 'content' ? 'Content' : 'Two columns'}</span></button>)}</div></section><section><h4>COLOR THEME</h4><div className="slides-theme-options">{THEMES.map(theme => <button key={theme.name} aria-label={theme.name} aria-pressed={activeSlide.background === theme.background} className={`slides-theme-option ${activeSlide.background === theme.background ? 'slides-option-selected' : ''}`} onClick={() => updateSlide({ background: theme.background })}><div className="slides-theme-swatch" style={{ background: theme.background, color: theme.accent }}><span>Aa</span>{activeSlide.background === theme.background && <Check size={13} />}</div><span>{theme.name}</span></button>)}</div><button className="slides-apply-theme" onClick={() => { onChange(slides.map(slide => ({ ...slide, background: activeSlide.background }))); onNotify('Theme applied to all slides'); }}>Apply theme to all slides</button></section><div className="slides-design-tip"><Presentation size={20} /><p>A little space goes a long way.<br /><span>Keep each slide focused on one idea.</span></p></div></aside>}
    </div>
    <div className="slides-print-pages">{slides.map((slide, index) => <div className="slides-print-page" key={slide.id}><SlidePreview slide={slide} index={index} /></div>)}</div>
    {presenting && <div className="slides-presenter" role="dialog" aria-modal="true" aria-label="Presentation"><button className="slides-presenter-close" onClick={() => setPresenting(false)}><X size={20} /> Exit presentation</button><div className="slides-presenter-screen"><SlidePreview slide={slides[Math.min(presentationIndex, slides.length - 1)]} index={presentationIndex} /></div><div className="slides-presenter-controls"><button aria-label="Previous slide" disabled={presentationIndex === 0} onClick={() => setPresentationIndex(index => Math.max(0, index - 1))}><ChevronLeft size={22} /></button><span>{presentationIndex + 1} / {slides.length}</span><button aria-label="Next slide" disabled={presentationIndex === slides.length - 1} onClick={() => setPresentationIndex(index => Math.min(slides.length - 1, index + 1))}><ChevronRight size={22} /></button></div><span className="slides-presenter-hint">Use ← → to navigate · Esc to exit</span></div>}
  </div>;
}
