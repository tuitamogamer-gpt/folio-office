import DOMPurify from 'dompurify';
import {
  AlignmentType, BorderStyle, commentIdToParaId, CommentRangeEnd, CommentRangeStart, CommentReference,
  Document, ExternalHyperlink, Footer, Header, HeadingLevel,
  ImageRun, LevelFormat, Packer, PageBreak, PageNumber, PageOrientation, Paragraph, ShadingType, Table, TableCell,
  TableRow, TextRun, UnderlineType, VerticalMergeType, WidthType,
  type IParagraphOptions, type IRunOptions, type INumberingOptions,
  type ParagraphChild,
} from 'docx';
import type { DocumentComment, DocumentPageSetup } from '../types';

type Block = Paragraph | Table;
type ListContext = { reference: string; level: number };
export type ExportOptions = Partial<DocumentPageSetup> & { comments?: DocumentComment[] };
type InlineContext = {
  maxImageWidth: number;
  commentStarts: Map<HTMLElement, number>;
  commentEnds: Map<HTMLElement, number>;
  anchoredComments: Set<number>;
};
const PAPER_SIZES = {
  a4: { width: 11906, height: 16838, css: 'A4' },
  letter: { width: 12240, height: 15840, css: 'letter' },
  legal: { width: 12240, height: 20160, css: 'legal' },
};

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

async function imageRun(element: HTMLImageElement, maxWidth: number): Promise<ParagraphChild[]> {
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
    const requestedWidth = Number(element.getAttribute('width')) || cssLengthTwips(element.style.width) / 15;
    const originalWidth = requestedWidth > 0 ? requestedWidth : image.naturalWidth || 480;
    const requestedHeight = Number(element.getAttribute('height')) || cssLengthTwips(element.style.height) / 15;
    const originalHeight = requestedHeight > 0 ? requestedHeight : originalWidth * (image.naturalHeight || 320) / (image.naturalWidth || 480);
    const scale = Math.min(1, maxWidth / originalWidth);
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

async function inline(nodes: Iterable<Node>, inherited: IRunOptions, context: InlineContext): Promise<ParagraphChild[]> {
  const runs: ParagraphChild[] = [];
  for (const node of nodes) {
    if (node.nodeType === Node.TEXT_NODE) {
      if (node.textContent) runs.push(new TextRun({ ...inherited, text: node.textContent }));
      continue;
    }
    if (!(node instanceof HTMLElement)) continue;
    const commentStart = context.commentStarts.get(node);
    if (commentStart !== undefined) {
      runs.push(new CommentRangeStart(commentStart));
      context.anchoredComments.add(commentStart);
    }
    const tag = node.tagName.toLowerCase();
    if (tag === 'br') runs.push(new TextRun({ break: 1 }));
    else if (tag === 'img') runs.push(...await imageRun(node as HTMLImageElement, context.maxImageWidth));
    else {
      const children = await inline(node.childNodes, runStyle(node, inherited), context);
      const href = node.getAttribute('href');
      if (tag === 'a' && href && /^(https?:|mailto:)/i.test(href)) {
        runs.push(new ExternalHyperlink({ link: href, children }));
      } else runs.push(...children);
    }
    const commentEnd = context.commentEnds.get(node);
    if (commentEnd !== undefined) runs.push(new CommentRangeEnd(commentEnd), new TextRun({ children: [new CommentReference(commentEnd)] }));
  }
  return runs;
}

function cssLengthTwips(value: string): number {
  const match = value.trim().match(/^([\d.]+)(px|pt|in|cm|mm)?$/i);
  if (!match) return 0;
  const units: Record<string, number> = { px: 15, pt: 20, in: 1440, cm: 1440 / 2.54, mm: 1440 / 25.4 };
  const length = Number(match[1]) * units[match[2]?.toLowerCase() || 'px'];
  return Number.isFinite(length) && length > 0 ? Math.round(length) : 0;
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
  const image = element.tagName === 'IMG' ? element : !element.textContent?.trim() ? element.querySelector('img') : null;
  const alignmentValue = element.style.textAlign || image?.getAttribute('data-align') || '';
  const alignment = {
    left: AlignmentType.LEFT, center: AlignmentType.CENTER,
    right: AlignmentType.RIGHT, justify: AlignmentType.JUSTIFIED,
  }[alignmentValue];
  const heading = {
    h1: HeadingLevel.HEADING_1, h2: HeadingLevel.HEADING_2, h3: HeadingLevel.HEADING_3,
    h4: HeadingLevel.HEADING_4, h5: HeadingLevel.HEADING_5, h6: HeadingLevel.HEADING_6,
  }[element.tagName.toLowerCase()];
  const level = Number(element.getAttribute('data-indent'));
  const indent = element.hasAttribute('data-indent') && Number.isFinite(level)
    ? Math.max(0, Math.min(8, level)) * 720 : cssLengthTwips(element.style.marginLeft);
  return {
    alignment, heading, indent: indent ? { left: indent } : undefined,
    pageBreakBefore: ['page', 'always'].includes(element.style.breakBefore || element.style.pageBreakBefore) || undefined,
    spacing: { after: 160, line: paragraphLineSpacing(element) },
  };
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

/** Build the document without opening a browser download. */
export async function buildDocumentBlob(html: string, name: string, format: 'docx' | 'html' | 'txt', options: ExportOptions = {}): Promise<Blob> {
  const safeHtml = DOMPurify.sanitize(html, { USE_PROFILES: { html: true } });
  const root = document.createElement('div');
  root.innerHTML = safeHtml;
  const paper = PAPER_SIZES[options.size || 'a4'] || PAPER_SIZES.a4;
  const margin = { normal: 1440, narrow: 720, wide: 2160 }[options.margin || 'normal'] || 1440;
  const comments = Array.from(new Map((options.comments || []).map(comment => [comment.id, comment])).values());
  if (format === 'txt') {
    return new Blob([plainText(root)], { type: 'text/plain;charset=utf-8' });
  }
  if (format === 'html') {
    const escape = (text: string) => {
      const span = document.createElement('span');
      span.textContent = text;
      return span.innerHTML;
    };
    const cssText = (text: string) => `"${text.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/[\r\n]/g, ' ').replace(/</g, '\\3c ').replace(/>/g, '\\3e ')}"`;
    const pageWidth = ((options.landscape ? paper.height : paper.width) - margin * 2) / 15;
    const footerText = options.footer || '';
    const pageFooter = `${cssText(footerText + (footerText && options.pageNumbers ? ' · ' : ''))}${options.pageNumbers ? ' "Page " counter(page)' : ''}`;
    const notes = comments.length ? `<aside class="document-comments"><h2>Comments</h2><ol>${comments.map(comment => `<li${comment.resolved ? ' data-resolved="true"' : ''}><p>${escape(comment.text)}${comment.resolved ? ' (Resolved)' : ''}</p>${comment.quote ? `<blockquote>${escape(comment.quote)}</blockquote>` : ''}</li>`).join('')}</ol></aside>` : '';
    return new Blob([`<!doctype html><html><head><meta charset="utf-8"><title>${escape(name || 'Untitled document')}</title><style>@page{size:${paper.css} ${options.landscape ? 'landscape' : 'portrait'};margin:${margin / 1440}in;@top-center{content:${cssText(options.header || '')};font:9pt Arial,sans-serif;color:#64748b}@bottom-center{content:${pageFooter};font:9pt Arial,sans-serif;color:#64748b}}body{font:11pt Arial,sans-serif;max-width:${pageWidth}px;margin:48px auto;line-height:1.65;padding:0 24px}h1{font:normal 30pt Georgia,serif;color:#243c2d}h2{font:normal 21pt Georgia,serif;color:#34573d}h3{font:normal 14pt Arial,sans-serif;color:#456b4a}table{border-collapse:collapse;width:100%}td,th{border:1px solid #cbd5e1;padding:8px}img{max-width:100%;height:auto}blockquote{border-left:3px solid #94a3b8;margin-left:0;padding-left:20px}[data-page-break]{break-after:page;page-break-after:always}.page-header,.page-footer{white-space:pre-line;text-align:center;color:#64748b;font-size:9pt;margin:24px 0}.document-comments{border-top:1px solid #cbd5e1;margin-top:36px}[data-comment-id]{background:#fef3c7}@media print{body{max-width:none;margin:0;padding:0}.page-header,.page-footer{display:none}}</style></head><body>${options.header ? `<header class="page-header">${escape(options.header)}</header>` : ''}<main>${safeHtml}</main>${footerText ? `<footer class="page-footer">${escape(footerText)}</footer>` : ''}${notes}</body></html>`], { type: 'text/html;charset=utf-8' });
  }

  const numbering: INumberingOptions['config'][number][] = [];
  const context: InlineContext = {
    maxImageWidth: ((options.landscape ? paper.height : paper.width) - margin * 2) / 15,
    commentStarts: new Map(), commentEnds: new Map(), anchoredComments: new Set(),
  };
  const commentIds = new Map(comments.map((comment, index) => [comment.id, index]));
  const commentNodes = new Map<number, HTMLElement[]>();
  for (const element of root.querySelectorAll<HTMLElement>('[data-comment-id]')) {
    const id = commentIds.get(element.getAttribute('data-comment-id') || '');
    if (id !== undefined) commentNodes.set(id, [...(commentNodes.get(id) || []), element]);
  }
  for (const [id, nodes] of commentNodes) {
    context.commentStarts.set(nodes[0], id);
    context.commentEnds.set(nodes[nodes.length - 1], id);
  }
  async function blocks(container: HTMLElement, depth = 0): Promise<Block[]> {
    const result: Block[] = [];
    let loose: Node[] = [];
    async function flush() {
      if (loose.some(n => n.nodeType !== Node.TEXT_NODE || n.textContent?.trim())) {
        result.push(new Paragraph({ children: await inline(loose, {}, context), spacing: { after: 160 } }));
      }
      loose = [];
    }
    for (const node of Array.from(container.childNodes)) {
      if (!(node instanceof HTMLElement)) { loose.push(node); continue; }
      const tag = node.tagName.toLowerCase();
      if (!/^(p|h[1-6]|div|blockquote|pre|ul|ol|table|hr|img)$/.test(tag)) { loose.push(node); continue; }
      await flush();
      if (node.getAttribute('data-page-break') === 'true') {
        result.push(new Paragraph({ children: [new PageBreak()], spacing: { after: 0, before: 0 } }));
      } else if (tag === 'img') {
        result.push(new Paragraph({ ...paragraphOptions(node), children: await imageRun(node as HTMLImageElement, context.maxImageWidth) }));
      } else if (tag === 'ul' || tag === 'ol') {
        const reference = `list-${numbering.length}`;
        const level = Math.min(depth, 8);
        numbering.push({ reference, levels: Array.from({ length: 9 }, (_, i) => ({
          level: i, format: tag === 'ol' ? LevelFormat.DECIMAL : LevelFormat.BULLET,
          text: tag === 'ol' ? `%${i + 1}.` : '•', start: Number(node.getAttribute('start')) || 1,
          alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: (i + 1) * 720, hanging: 360 } } },
        })) });
        const listContext: ListContext = { reference, level };
        for (const item of Array.from(node.children)) {
          if (!(item instanceof HTMLElement) || item.tagName !== 'LI') continue;
          let first = true;
          let direct: Node[] = [];
          const addParagraph = async (element?: HTMLElement) => {
            const children = await inline(element ? element.childNodes : direct, element ? runStyle(element, {}) : {}, context);
            result.push(new Paragraph({ ...(element ? paragraphOptions(element) : {}), children,
              numbering: first ? listContext : undefined, indent: first ? undefined : { left: (level + 1) * 720 }, spacing: { after: 100, line: paragraphLineSpacing(element || item) },
            }));
            first = false;
            direct = [];
          };
          for (const child of Array.from(item.childNodes)) {
            if (child instanceof HTMLElement && /^(UL|OL)$/.test(child.tagName)) {
              if (direct.length || first) await addParagraph();
              const wrapper = document.createElement('div');
              // Keep node identity: comment anchors refer to the sanitized
              // elements, including marks inside nested list items.
              wrapper.appendChild(child);
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
        const htmlRows = Array.from(node.querySelectorAll('tr')).filter(row => row.closest('table') === node);
        type Span = { left: number; columns: number; fill?: string };
        let spans = new Map<number, Span>();
        for (const [rowIndex, row] of htmlRows.entries()) {
          const cells: TableCell[] = [];
          const nextSpans = new Map<number, Span>();
          let column = 0;
          const continueSpans = () => {
            while (spans.has(column)) {
              const span = spans.get(column)!;
              cells.push(new TableCell({ children: [new Paragraph('')], columnSpan: span.columns,
                verticalMerge: VerticalMergeType.CONTINUE,
                shading: span.fill ? { type: ShadingType.CLEAR, fill: span.fill } : undefined,
              }));
              if (span.left > 1) nextSpans.set(column, { ...span, left: span.left - 1 });
              spans.delete(column);
              column += span.columns;
            }
          };
          for (const cell of Array.from(row.children)) {
            if (!(cell instanceof HTMLElement) || !['TD', 'TH'].includes(cell.tagName)) continue;
            continueSpans();
            const children = await blocks(cell, depth);
            if (!children.length || children[children.length - 1] instanceof Table) children.push(new Paragraph(''));
            const columnSpan = Math.min(1000, Math.max(1, Math.floor(Number(cell.getAttribute('colspan'))) || 1));
            const rowSpan = Math.min(htmlRows.length - rowIndex, Math.max(1, Math.floor(Number(cell.getAttribute('rowspan'))) || 1));
            const fill = color(cell.style.backgroundColor || cell.getAttribute('bgcolor') || '') || (cell.tagName === 'TH' ? 'F1F5F9' : undefined);
            cells.push(new TableCell({ children, columnSpan,
              verticalMerge: rowSpan > 1 ? VerticalMergeType.RESTART : undefined,
              shading: fill ? { type: ShadingType.CLEAR, fill } : undefined,
              margins: { top: 100, bottom: 100, left: 120, right: 120 },
            }));
            if (rowSpan > 1) nextSpans.set(column, { left: rowSpan - 1, columns: columnSpan, fill });
            column += columnSpan;
          }
          while (spans.size) {
            if (spans.has(column)) continueSpans();
            else if (Math.min(...spans.keys()) < column) break;
            else { cells.push(new TableCell({ children: [new Paragraph('')] })); column++; }
          }
          spans = nextSpans;
          if (cells.length) rows.push(new TableRow({ children: cells, tableHeader: row.children.length > 0 && Array.from(row.children).every(cell => cell.tagName === 'TH') }));
        }
        if (rows.length) result.push(new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } }));
      } else if (tag === 'div' || tag === 'blockquote') result.push(...await blocks(node, depth));
      else if (tag === 'hr') result.push(new Paragraph({ border: { bottom: { color: 'CBD5E1', style: BorderStyle.SINGLE, size: 6 } } }));
      else result.push(new Paragraph({ ...paragraphOptions(node), children: await inline(node.childNodes, runStyle(node, {}), context) }));
    }
    await flush();
    return result;
  }

  const children = await blocks(root);
  // Comments whose selected text was deleted remain as point annotations at
  // the end, with their original quote included in the comment itself.
  const unanchoredComments = comments.map((_, id) => id).filter(id => !context.anchoredComments.has(id));
  if (unanchoredComments.length) children.push(new Paragraph({ children: unanchoredComments.flatMap(id => [
    new CommentRangeStart(id), new CommentRangeEnd(id), new TextRun({ children: [new CommentReference(id)] }),
  ]) }));
  const footerChildren: ParagraphChild[] = [];
  if (options.footer) footerChildren.push(new TextRun({ text: options.footer, color: '64748B', size: 18 }));
  if (options.pageNumbers) footerChildren.push(new TextRun({ text: options.footer ? ' · Page ' : 'Page ', size: 18 }), new TextRun({ children: [PageNumber.CURRENT], size: 18 }));
  const doc = new Document({
    title: name, creator: 'Folio', numbering: { config: numbering },
    comments: comments.length ? { children: comments.map((comment, id) => ({
      id, author: 'Folio', initials: 'F', date: new Date(Number.isFinite(comment.createdAt) ? comment.createdAt : Date.now()),
      durableId: commentIdToParaId(id),
      children: [
        ...(!context.anchoredComments.has(id) && comment.quote ? [new Paragraph({ children: [new TextRun({ text: `Original selection: ${comment.quote}`, italics: true })] })] : []),
        ...comment.text.split(/\r?\n/).map(text => new Paragraph(text)),
      ],
    })) } : undefined,
    styles: { default: {
      document: { run: { font: 'Arial', size: 22 }, paragraph: { spacing: { after: 160, line: 396 } } },
      heading1: { run: { font: 'Georgia', size: 60, bold: false, color: '243C2D' } },
      heading2: { run: { font: 'Georgia', size: 42, bold: false, color: '34573D' } },
      heading3: { run: { font: 'Arial', size: 28, bold: false, color: '456B4A' } },
    } },
    sections: [{ properties: { page: {
      size: { width: paper.width, height: paper.height, orientation: options.landscape ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT },
      margin: { top: margin, right: margin, bottom: margin, left: margin, header: Math.min(720, margin / 2), footer: Math.min(720, margin / 2) },
    } },
    headers: options.header ? { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: options.header, color: '64748B', size: 18 })] })] }) } : undefined,
    footers: footerChildren.length ? { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: footerChildren })] }) } : undefined,
    children: children.length ? children : [new Paragraph('')] }],
  });
  // docx emits resolved state only for reply threads. Register the standard
  // extended-comments part explicitly so standalone comments keep that state.
  const extraParts: { path: string; data: string }[] = [];
  if (comments.length) {
    doc.ContentTypes.addCommentsExtended();
    doc.Document.Relationships.addRelationship('FolioCommentState', 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended', 'commentsExtended.xml');
    extraParts.push({ path: 'word/commentsExtended.xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">${comments.map((comment, id) => `<w15:commentEx w15:paraId="${commentIdToParaId(id)}" w15:done="${comment.resolved ? '1' : '0'}"/>`).join('')}</w15:commentsEx>` });
  }
  return Packer.toBlob(doc, false, extraParts);
}

export async function exportDocument(html: string, name: string, format: 'docx' | 'html' | 'txt', options: ExportOptions = {}): Promise<void> {
  const filename = (name.trim() || 'Untitled document').replace(/[\\/:*?"<>|]/g, '-').replace(/\.(docx|html?|txt)$/i, '') || 'Untitled document';
  download(await buildDocumentBlob(html, name, format, options), `${filename}.${format}`);
}
