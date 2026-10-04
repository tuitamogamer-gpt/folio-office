import DOMPurify from 'dompurify';
import {
  AlignmentType, BorderStyle, Document, ExternalHyperlink, HeadingLevel,
  ImageRun, LevelFormat, Packer, PageOrientation, Paragraph, ShadingType, Table, TableCell,
  TableRow, TextRun, UnderlineType, WidthType,
  type IParagraphOptions, type IRunOptions, type INumberingOptions,
  type ParagraphChild,
} from 'docx';

type Block = Paragraph | Table;
type ListContext = { reference: string; level: number };

function download(blob: Blob, name: string) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = name;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30_000);
}

function color(value: string): string | undefined {
  if (!value || value === 'transparent') return undefined;
  if (/^#[\da-f]{6}$/i.test(value)) return value.slice(1).toUpperCase();
  if (/^#[\da-f]{3}$/i.test(value)) return value.slice(1).split('').map(c => c + c).join('').toUpperCase();
  const canvas = document.createElement('canvas');
  const context = canvas.getContext('2d');
  if (!context) return undefined;
  context.fillStyle = value;
  const normalized = context.fillStyle;
  if (/^#[\da-f]{6}$/i.test(normalized)) return normalized.slice(1).toUpperCase();
  const channels = normalized.match(/[\d.]+/g);
  return channels && channels.length >= 3
    ? channels.slice(0, 3).map(v => Math.round(Number(v)).toString(16).padStart(2, '0')).join('').toUpperCase()
    : undefined;
}

function runStyle(element: HTMLElement, inherited: IRunOptions): IRunOptions {
  const style = element.style;
  const tag = element.tagName.toLowerCase();
  const result = { ...inherited };
  if (['b', 'strong'].includes(tag) || style.fontWeight === 'bold' || Number(style.fontWeight) >= 600) result.bold = true;
  if (['em', 'i'].includes(tag) || style.fontStyle === 'italic') result.italics = true;
  if (tag === 'u' || style.textDecoration.includes('underline')) result.underline = { type: UnderlineType.SINGLE };
  if (['s', 'del', 'strike'].includes(tag) || style.textDecoration.includes('line-through')) result.strike = true;
  if (tag === 'sub') result.subScript = true;
  if (tag === 'sup') result.superScript = true;
  if (tag === 'code') result.font = 'Consolas';
  if (style.fontFamily) result.font = style.fontFamily.split(',')[0].replace(/["']/g, '').trim();
  if (style.fontSize) {
    const size = parseFloat(style.fontSize);
    if (Number.isFinite(size)) result.size = Math.round(size * (style.fontSize.endsWith('pt') ? 2 : 1.5));
  }
  const foreground = color(style.color);
  if (foreground) result.color = foreground;
  const background = color(style.backgroundColor || (tag === 'mark' ? '#ffff00' : ''));
  if (background) result.shading = { type: ShadingType.CLEAR, fill: background, color: 'auto' };
  return result;
}

async function imageRun(element: HTMLImageElement): Promise<ParagraphChild[]> {
  const source = element.getAttribute('src') || '';
  if (!/^data:image\//i.test(source)) return [new TextRun(element.alt ? `[Image: ${element.alt}]` : '[Image]')];
  try {
    const image = new Image();
    image.src = source;
    await image.decode();
    const match = source.match(/^data:image\/(png|jpeg|jpg|gif|bmp);base64,([\s\S]+)$/i);
    let bytes: Uint8Array;
    let type: 'png' | 'jpg' | 'gif' | 'bmp';
    if (match) {
      bytes = Uint8Array.from(atob(match[2]), c => c.charCodeAt(0));
      type = match[1].toLowerCase().replace('jpeg', 'jpg') as typeof type;
    } else {
      // Word does not support every browser image format. Rasterize decoded
      // WebP, AVIF, SVG, and other supported data images to a portable PNG.
      const canvas = document.createElement('canvas');
      canvas.width = image.naturalWidth || 480;
      canvas.height = image.naturalHeight || 320;
      const context = canvas.getContext('2d');
      if (!context) throw new Error('Image conversion unavailable');
      context.drawImage(image, 0, 0, canvas.width, canvas.height);
      bytes = Uint8Array.from(atob(canvas.toDataURL('image/png').split(',')[1]), c => c.charCodeAt(0));
      type = 'png';
    }
    const originalWidth = Number(element.getAttribute('width')) || image.naturalWidth || 480;
    const originalHeight = Number(element.getAttribute('height')) || image.naturalHeight || 320;
    const scale = Math.min(1, 600 / originalWidth);
    return [new ImageRun({
      data: bytes,
      type,
      transformation: { width: Math.round(originalWidth * scale), height: Math.round(originalHeight * scale) },
      altText: { name: element.alt || 'Document image', title: element.alt || 'Document image', description: element.alt || '' },
    })];
  } catch {
    return [new TextRun(element.alt ? `[Image: ${element.alt}]` : '[Image]')];
  }
}

async function inline(nodes: Iterable<Node>, inherited: IRunOptions = {}): Promise<ParagraphChild[]> {
  const runs: ParagraphChild[] = [];
  for (const node of nodes) {
    if (node.nodeType === Node.TEXT_NODE) {
      if (node.textContent) runs.push(new TextRun({ ...inherited, text: node.textContent }));
      continue;
    }
    if (!(node instanceof HTMLElement)) continue;
    const tag = node.tagName.toLowerCase();
    if (tag === 'br') runs.push(new TextRun({ break: 1 }));
    else if (tag === 'img') runs.push(...await imageRun(node as HTMLImageElement));
    else {
      const children = await inline(node.childNodes, runStyle(node, inherited));
      const href = node.getAttribute('href');
      if (tag === 'a' && href && /^(https?:|mailto:)/i.test(href)) {
        runs.push(new ExternalHyperlink({ link: href, children }));
      } else runs.push(...children);
    }
  }
  return runs;
}

function paragraphLineSpacing(element: HTMLElement): number {
  const source = element.style.lineHeight ? element : Array.from(element.querySelectorAll<HTMLElement>('[style]')).find(node => node.style.lineHeight);
  const value = source?.style.lineHeight || '';
  const number = parseFloat(value);
  if (!Number.isFinite(number) || number <= 0) return 396;
  if (value.endsWith('%')) return Math.round(number * 2.4);
  if (/^[\d.]+$/.test(value) || value.endsWith('em')) return Math.round(number * 240);
  const font = source?.style.fontSize || element.style.fontSize;
  const fontPoints = font ? parseFloat(font) * (font.endsWith('pt') ? 1 : 0.75) : 11;
  const linePoints = number * (value.endsWith('pt') ? 1 : 0.75);
  return Number.isFinite(fontPoints) && fontPoints > 0 ? Math.round(linePoints / fontPoints * 240) : 396;
}

function paragraphOptions(element: HTMLElement): IParagraphOptions {
  const alignment = {
    left: AlignmentType.LEFT, center: AlignmentType.CENTER,
    right: AlignmentType.RIGHT, justify: AlignmentType.JUSTIFIED,
  }[element.style.textAlign];
  const heading = {
    h1: HeadingLevel.HEADING_1, h2: HeadingLevel.HEADING_2, h3: HeadingLevel.HEADING_3,
    h4: HeadingLevel.HEADING_4, h5: HeadingLevel.HEADING_5, h6: HeadingLevel.HEADING_6,
  }[element.tagName.toLowerCase()];
  return { alignment, heading, spacing: { after: 160, line: paragraphLineSpacing(element) } };
}

function plainText(root: HTMLElement): string {
  const walk = (node: Node): string => {
    if (node.nodeType === Node.TEXT_NODE) return node.textContent || '';
    if (!(node instanceof HTMLElement)) return '';
    const tag = node.tagName.toLowerCase();
    if (tag === 'br') return '\n';
    if (tag === 'img') return node.getAttribute('alt') || '';
    const text = Array.from(node.childNodes).map(walk).join('');
    if (tag === 'li') {
      const parent = node.parentElement;
      const index = parent ? Array.from(parent.children).indexOf(node) + Number(parent.getAttribute('start') || 1) : 1;
      return `${parent?.tagName === 'OL' ? `${index}.` : '•'} ${text.trim()}\n`;
    }
    if (tag === 'td' || tag === 'th') return `${text.trim()}\t`;
    if (/^(p|h[1-6]|div|blockquote|pre|tr|ul|ol|table)$/.test(tag)) return `${text.trimEnd()}\n`;
    return text;
  };
  return walk(root).replace(/\t\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
}

export async function exportDocument(html: string, name: string, format: 'docx' | 'html' | 'txt', options: { landscape?: boolean; margin?: 'normal' | 'narrow' | 'wide' } = {}): Promise<void> {
  const safeHtml = DOMPurify.sanitize(html, { USE_PROFILES: { html: true } });
  const root = document.createElement('div');
  root.innerHTML = safeHtml;
  const filename = (name.trim() || 'Untitled document').replace(/[\\/:*?"<>|]/g, '-').replace(/\.(docx|html?|txt)$/i, '');
  if (format === 'txt') {
    download(new Blob([plainText(root)], { type: 'text/plain;charset=utf-8' }), `${filename}.txt`);
    return;
  }
  if (format === 'html') {
    const title = document.createElement('span');
    title.textContent = name || 'Untitled document';
    download(new Blob([`<!doctype html><html><head><meta charset="utf-8"><title>${title.innerHTML}</title><style>body{font:11pt Arial,sans-serif;max-width:760px;margin:48px auto;line-height:1.65;padding:0 24px}h1{font:normal 30pt Georgia,serif;color:#243c2d}h2{font:normal 21pt Georgia,serif;color:#34573d}h3{font:normal 14pt Arial,sans-serif;color:#456b4a}table{border-collapse:collapse;width:100%}td,th{border:1px solid #cbd5e1;padding:8px}img{max-width:100%;height:auto}blockquote{border-left:3px solid #94a3b8;margin-left:0;padding-left:20px}</style></head><body>${safeHtml}</body></html>`], { type: 'text/html;charset=utf-8' }), `${filename}.html`);
    return;
  }

  const numbering: INumberingOptions['config'][number][] = [];
  async function blocks(container: HTMLElement, depth = 0): Promise<Block[]> {
    const result: Block[] = [];
    let loose: Node[] = [];
    async function flush() {
      if (loose.some(n => n.nodeType !== Node.TEXT_NODE || n.textContent?.trim())) {
        result.push(new Paragraph({ children: await inline(loose), spacing: { after: 160 } }));
      }
      loose = [];
    }
    for (const node of Array.from(container.childNodes)) {
      if (!(node instanceof HTMLElement)) { loose.push(node); continue; }
      const tag = node.tagName.toLowerCase();
      if (!/^(p|h[1-6]|div|blockquote|pre|ul|ol|table|hr)$/.test(tag)) { loose.push(node); continue; }
      await flush();
      if (tag === 'ul' || tag === 'ol') {
        const reference = `list-${numbering.length}`;
        const level = Math.min(depth, 8);
        numbering.push({ reference, levels: Array.from({ length: 9 }, (_, i) => ({
          level: i, format: tag === 'ol' ? LevelFormat.DECIMAL : LevelFormat.BULLET,
          text: tag === 'ol' ? `%${i + 1}.` : '•', start: Number(node.getAttribute('start')) || 1,
          alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: (i + 1) * 720, hanging: 360 } } },
        })) });
        const context: ListContext = { reference, level };
        for (const item of Array.from(node.children)) {
          if (!(item instanceof HTMLElement) || item.tagName !== 'LI') continue;
          let first = true;
          let direct: Node[] = [];
          const addParagraph = async (element?: HTMLElement) => {
            const children = await inline(element ? element.childNodes : direct, element ? runStyle(element, {}) : {});
            result.push(new Paragraph({ ...(element ? paragraphOptions(element) : {}), children,
              numbering: first ? context : undefined, indent: first ? undefined : { left: (level + 1) * 720 }, spacing: { after: 100, line: paragraphLineSpacing(element || item) },
            }));
            first = false;
            direct = [];
          };
          for (const child of Array.from(item.childNodes)) {
            if (child instanceof HTMLElement && /^(UL|OL)$/.test(child.tagName)) {
              if (direct.length || first) await addParagraph();
              const wrapper = document.createElement('div');
              wrapper.appendChild(child.cloneNode(true));
              result.push(...await blocks(wrapper, depth + 1));
            } else if (child instanceof HTMLElement && /^(P|H[1-6])$/.test(child.tagName)) {
              if (direct.some(n => n.textContent?.trim())) await addParagraph();
              await addParagraph(child);
            } else direct.push(child);
          }
          if (direct.some(n => n.textContent?.trim()) || first) await addParagraph();
        }
      } else if (tag === 'table') {
        const rows: TableRow[] = [];
        for (const row of Array.from(node.querySelectorAll('tr')).filter(row => row.closest('table') === node)) {
          const cells: TableCell[] = [];
          for (const cell of Array.from(row.children)) {
            if (!(cell instanceof HTMLElement)) continue;
            const children = await blocks(cell, depth);
            if (!children.length || children[children.length - 1] instanceof Table) children.push(new Paragraph(''));
            cells.push(new TableCell({ children, columnSpan: Number(cell.getAttribute('colspan')) || 1,
              shading: cell.tagName === 'TH' ? { fill: 'F1F5F9' } : undefined,
              margins: { top: 100, bottom: 100, left: 120, right: 120 },
            }));
          }
          if (cells.length) rows.push(new TableRow({ children: cells }));
        }
        if (rows.length) result.push(new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } }));
      } else if (tag === 'div' || tag === 'blockquote') result.push(...await blocks(node, depth));
      else if (tag === 'hr') result.push(new Paragraph({ border: { bottom: { color: 'CBD5E1', style: BorderStyle.SINGLE, size: 6 } } }));
      else result.push(new Paragraph({ ...paragraphOptions(node), children: await inline(node.childNodes, runStyle(node, {})) }));
    }
    await flush();
    return result;
  }

  const children = await blocks(root);
  const margin = { normal: 1440, narrow: 720, wide: 2160 }[options.margin || 'normal'];
  const doc = new Document({
    title: name, creator: 'Folio', numbering: { config: numbering },
    styles: { default: {
      document: { run: { font: 'Arial', size: 22 }, paragraph: { spacing: { after: 160, line: 396 } } },
      heading1: { run: { font: 'Georgia', size: 60, bold: false, color: '243C2D' } },
      heading2: { run: { font: 'Georgia', size: 42, bold: false, color: '34573D' } },
      heading3: { run: { font: 'Arial', size: 28, bold: false, color: '456B4A' } },
    } },
    sections: [{ properties: { page: {
      size: { width: 11906, height: 16838, orientation: options.landscape ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT },
      margin: { top: margin, right: margin, bottom: margin, left: margin },
    } }, children: children.length ? children : [new Paragraph('')] }],
  });
  download(await Packer.toBlob(doc), `${filename}.docx`);
}
