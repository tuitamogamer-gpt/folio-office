import type { SlideData } from '../types';

export const SLIDE_WIDTH = 13.333333;
export const SLIDE_HEIGHT = 7.5;
export const SLIDE_THEMES = [
  { name: 'Paper', background: '#ffffff', accent: '#2d6554' },
  { name: 'Forest', background: '#244e40', accent: '#d4e6bb' },
  { name: 'Sage', background: '#e6ece2', accent: '#355544' },
  { name: 'Sand', background: '#f4eadd', accent: '#8d5e3d' },
  { name: 'Ink', background: '#25374b', accent: '#bdd5ec' },
  { name: 'Lavender', background: '#ece7f5', accent: '#6f5689' },
];

export function slideIsDark(hex: string) {
  const value = hex.replace('#', '');
  return value.length === 6 && parseInt(value.slice(0, 2), 16) * .299 + parseInt(value.slice(2, 4), 16) * .587 + parseInt(value.slice(4, 6), 16) * .114 < 140;
}

export function slideTextColor(slide: SlideData) {
  return slide.textColor || (slideIsDark(slide.background) ? '#f6f7f3' : '#253a34');
}

export function slideAccent(slide: SlideData) {
  return SLIDE_THEMES.find(theme => theme.background === slide.background)?.accent || (slideIsDark(slide.background) ? '#d4e6bb' : '#2d6554');
}

export function slideLayout(slide: SlideData) {
  if (slide.layout === 'title') return { title: { x: 7, y: 25, w: 86, h: 21 }, body: { x: 7, y: 63, w: 86, h: 21 }, subtitle: { x: 7, y: 48, w: 86, h: 10 } };
  if (slide.layout === 'split') return { title: { x: 7, y: 27, w: 39, h: 48 }, body: { x: 54, y: 26, w: 39, h: 53 }, subtitle: { x: 7, y: 48, w: 86, h: 10 } };
  return { title: { x: 7, y: 19, w: 86, h: 23 }, body: { x: 7, y: 43, w: 86, h: 40 }, subtitle: { x: 7, y: 48, w: 86, h: 10 } };
}

const hex = (value: string) => /^#[a-f0-9]{6}$/i.test(value) ? value.slice(1) : '253A34';

export async function buildPresentationBlob(slides: SlideData[], name: string): Promise<Blob> {
  const [{ default: PptxGenJS }, { default: JSZip }] = await Promise.all([import('pptxgenjs'), import('jszip')]);
  const presentation = new PptxGenJS();
  presentation.layout = 'LAYOUT_WIDE';
  presentation.author = 'Folio';
  presentation.subject = name;
  presentation.title = name;
  presentation.company = 'Folio';
  presentation.theme = { headFontFace: 'Aptos Display', bodyFontFace: 'Aptos' };
  slides.forEach((data, index) => {
    const slide = presentation.addSlide();
    const color = hex(slideTextColor(data));
    const accent = hex(slideAccent(data));
    const positions = slideLayout(data);
    const position = ({ x, y, w, h }: { x: number; y: number; w: number; h: number }) => ({ x: x * SLIDE_WIDTH / 100, y: y * SLIDE_HEIGHT / 100, w: w * SLIDE_WIDTH / 100, h: h * SLIDE_HEIGHT / 100 });
    slide.background = { color: hex(data.background) };
    slide.hidden = !!data.hidden;
    if (data.notes) slide.addNotes(data.notes);
    // Store Folio-only settings in an invisible, named object. PowerPoint still
    // receives normal editable shapes, pictures and text for all visible content.
    const { elements: _elements, notes: _notes, id: _id, ...metadata } = data;
    slide.addShape(presentation.ShapeType.rect, { objectName: `folio-meta-${encodeURIComponent(JSON.stringify({ version: 1, ...metadata }))}`, x: 0, y: 0, w: .001, h: .001, fill: { color: 'FFFFFF', transparency: 100 }, line: { color: 'FFFFFF', transparency: 100 } });
    if (data.layout !== 'blank') {
      slide.addShape(presentation.ShapeType.rect, { objectName: 'folio-decoration-mark', x: .8, y: .6375, w: .24, h: .24, fill: { color: accent }, line: { color: accent } });
      slide.addShape(presentation.ShapeType.rect, { objectName: 'folio-decoration-mark-secondary', x: 1.08, y: .6375, w: .14, h: .24, fill: { color: accent, transparency: 65 }, line: { color: accent, transparency: 100 } });
      slide.addText(data.title, { objectName: `folio-title-${data.layout}`, ...position(positions.title), fontFace: data.fontFamily || 'Aptos Display', fontSize: data.titleSize || (data.layout === 'title' ? 42 : 34), bold: true, color, margin: 0, fit: 'shrink', valign: 'top', align: data.layout === 'title' ? 'center' : 'left' });
      if (data.layout === 'title' && data.subtitle) slide.addText(data.subtitle, { objectName: 'folio-subtitle', ...position(positions.subtitle), fontFace: data.fontFamily || 'Aptos', fontSize: 22, color, align: 'center', margin: 0, fit: 'shrink' });
      slide.addText(data.body, { objectName: 'folio-body', ...position(positions.body), fontFace: data.fontFamily || 'Aptos', fontSize: data.bodySize || 23, color, margin: 0, fit: 'shrink', align: data.layout === 'title' ? 'center' : 'left', valign: 'top', paraSpaceAfter: 12 });
      slide.addShape(presentation.ShapeType.line, { objectName: 'folio-decoration-rule', x: .8, y: 6.825, w: 11.733, h: 0, line: { color: accent, width: 1, transparency: 55 } });
      slide.addText(String(index + 1).padStart(2, '0'), { objectName: 'folio-page-number', x: 11.8, y: 6.98, w: .733, h: .2, fontFace: data.fontFamily || 'Aptos', fontSize: 10, color, align: 'right', margin: 0 });
    }
    for (const element of data.elements || []) {
      const box = { objectName: `folio-element-${element.id}`, x: element.x * SLIDE_WIDTH / 100, y: element.y * SLIDE_HEIGHT / 100, w: element.width * SLIDE_WIDTH / 100, h: element.height * SLIDE_HEIGHT / 100, rotate: element.rotation || 0 };
      if (element.type === 'image' && element.src) slide.addImage({ ...box, data: element.src });
      if (element.type === 'text') slide.addText(element.text || '', { ...box, fontFace: data.fontFamily || 'Aptos', fontSize: element.fontSize || 18, bold: !!element.bold, color: hex(element.color || slideTextColor(data)), align: element.align || 'left', margin: 0, valign: 'top', fit: 'resize', breakLine: false, fill: element.fill && element.fill !== 'transparent' ? { color: hex(element.fill) } : { color: 'FFFFFF', transparency: 100 } });
      if (element.type === 'shape') {
        slide.addShape(element.shape === 'ellipse' ? presentation.ShapeType.ellipse : element.shape === 'line' ? presentation.ShapeType.line : presentation.ShapeType.rect, { ...box, line: { color: hex(element.color || element.fill || slideAccent(data)), width: element.shape === 'line' ? 2 : 1 }, fill: element.fill === 'transparent' ? { color: 'FFFFFF', transparency: 100 } : { color: hex(element.fill || slideAccent(data)) } });
      }
    }
  });
  const raw = await presentation.write({ outputType: 'arraybuffer' });
  const zip = await JSZip.loadAsync(raw as ArrayBuffer);
  for (let index = 0; index < slides.length; index++) {
    const transition = slides[index].transition;
    if (!transition || transition === 'none') continue;
    const path = `ppt/slides/slide${index + 1}.xml`;
    const xml = await zip.file(path)!.async('string');
    // p:transition follows clrMapOvr (or cSld when no override is present).
    const effect = `<p:transition spd="med">${transition === 'slide' ? '<p:push dir="l"/>' : '<p:fade/>'}</p:transition>`;
    zip.file(path, xml.includes('</p:clrMapOvr>') ? xml.replace('</p:clrMapOvr>', `</p:clrMapOvr>${effect}`) : xml.replace('</p:cSld>', `</p:cSld>${effect}`));
  }
  return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', compression: 'DEFLATE' });
}

export async function exportPresentationPptx(slides: SlideData[], name: string): Promise<void> {
  const blob = await buildPresentationBlob(slides, name);
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `${name.replace(/[<>:"/\\|?*]/g, '-')}.pptx`;
  link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}
