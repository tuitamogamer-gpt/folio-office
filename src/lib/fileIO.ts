import DOMPurify from 'dompurify';
import type { DocumentComment, DocumentPageSetup, FileKind, OfficeFile, SheetContent, SlideData, SlideElement } from '../types';

export interface ImportedOfficeFile {
  name: string;
  kind: FileKind;
  content: string | SheetContent | SlideData[];
  warning?: string;
  pageSetup?: DocumentPageSetup;
  comments?: DocumentComment[];
}

const MAX_FILE_SIZE = 30 * 1024 * 1024;
const SUPPORTED_EXTENSIONS = new Set(['docx', 'txt', 'html', 'htm', 'md', 'xlsx', 'xls', 'csv', 'pptx']);

function baseName(name: string) {
  return name.replace(/\.[^.]+$/, '').trim() || 'Untitled';
}

function safeFilename(name: string, extension: string) {
  return `${name.replace(/\.(docx|xlsx|pptx|csv|txt|html|htm|md)$/i, '').replace(/[<>:"/\\|?*\x00-\x1f]/g, '-').trim() || 'Untitled'}.${extension}`;
}

export function downloadBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  // Keep the URL alive until the browser has started receiving the download.
  window.setTimeout(() => URL.revokeObjectURL(url), 30_000);
}

function escapeHTML(value: string) {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function textAsHTML(value: string) {
  return value.replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n').split('\n').map(line => `<p>${escapeHTML(line) || '<br>'}</p>`).join('');
}

function cleanHTML(value: string) {
  return DOMPurify.sanitize(value, {
    USE_PROFILES: { html: true },
    FORBID_TAGS: ['script', 'style', 'iframe', 'object', 'embed', 'form', 'input', 'button', 'textarea', 'select'],
  }) || '<p></p>';
}

function markdownAsHTML(value: string) {
  // A conservative Markdown subset. Escape first so raw HTML never becomes active content.
  const inline = (text: string) => escapeHTML(text)
    .replace(/`([^`]+)`/g, '<code>$1</code>')
    .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
    .replace(/__([^_]+)__/g, '<strong>$1</strong>')
    .replace(/\*([^*]+)\*/g, '<em>$1</em>')
    .replace(/\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g, '<a href="$2">$1</a>');
  const lines = value.replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n').split('\n');
  const parts: string[] = [];
  let list: 'ul' | 'ol' | null = null;
  let inCode = false;
  const closeList = () => { if (list) { parts.push(`</${list}>`); list = null; } };
  for (const line of lines) {
    if (/^\s*```/.test(line)) {
      closeList();
      parts.push(inCode ? '</code></pre>' : '<pre><code>');
      inCode = !inCode;
      continue;
    }
    if (inCode) { parts.push(`${escapeHTML(line)}\n`); continue; }
    const bullet = line.match(/^\s*[-+*]\s+(.+)$/);
    const ordered = line.match(/^\s*\d+[.)]\s+(.+)$/);
    if (bullet || ordered) {
      const nextList = bullet ? 'ul' : 'ol';
      if (list !== nextList) { closeList(); list = nextList; parts.push(`<${list}>`); }
      parts.push(`<li><p>${inline((bullet || ordered)![1])}</p></li>`);
      continue;
    }
    closeList();
    const heading = line.match(/^(#{1,6})\s+(.+)$/);
    if (heading) parts.push(`<h${heading[1].length}>${inline(heading[2])}</h${heading[1].length}>`);
    else if (/^>\s?/.test(line)) parts.push(`<blockquote><p>${inline(line.replace(/^>\s?/, ''))}</p></blockquote>`);
    else if (/^\s*(?:---+|\*\*\*+)\s*$/.test(line)) parts.push('<hr>');
    else parts.push(`<p>${inline(line) || '<br>'}</p>`);
  }
  closeList();
  if (inCode) parts.push('</code></pre>');
  return cleanHTML(parts.join(''));
}

function xmlDocument(source: string) {
  const xml = new DOMParser().parseFromString(source, 'application/xml');
  if (xml.getElementsByTagName('parsererror').length) throw new Error('This Office file contains invalid XML. Try saving a new copy.');
  return xml;
}

function elements(node: Document | Element, localName: string) {
  return Array.from(node.getElementsByTagNameNS('*', localName));
}

function resolvePackagePath(base: string, target: string) {
  if (/^[a-z]+:/i.test(target)) return '';
  const path = target.startsWith('/') ? target.slice(1) : `${base.slice(0, base.lastIndexOf('/') + 1)}${target}`;
  const segments: string[] = [];
  for (const segment of path.split('/')) {
    if (segment === '..') segments.pop();
    else if (segment && segment !== '.') segments.push(segment);
  }
  return segments.join('/');
}

function wordAttribute(element: Element | undefined, name: string) {
  return element?.getAttributeNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', name) || element?.getAttribute(`w:${name}`) || '';
}

async function documentMetadata(buffer: ArrayBuffer): Promise<{ pageSetup: DocumentPageSetup; comments: DocumentComment[]; warning?: string }> {
  const { default: JSZip } = await import('jszip');
  const zip = await JSZip.loadAsync(buffer);
  const entry = zip.file('word/document.xml');
  if (!entry) throw new Error('This file is not a valid Word document.');
  const document = xmlDocument(await entry.async('string'));
  const sections = elements(document, 'sectPr');
  const section = sections[sections.length - 1];
  const size = section ? elements(section, 'pgSz')[0] : undefined;
  const margins = section ? elements(section, 'pgMar')[0] : undefined;
  const width = Number(wordAttribute(size, 'w'));
  const height = Number(wordAttribute(size, 'h'));
  const shorter = Math.min(width, height);
  const longer = Math.max(width, height);
  const margin = Number(wordAttribute(margins, 'left'));
  const pageSetup: DocumentPageSetup = { landscape: wordAttribute(size, 'orient') === 'landscape' || width > height, margin: margin > 0 && margin < 1080 ? 'narrow' : margin >= 2160 ? 'wide' : 'normal', size: longer > 18000 ? 'legal' : Math.abs(shorter - 12240) < 100 ? 'letter' : 'a4', header: '', footer: '', pageNumbers: false };
  const relationshipEntry = zip.file('word/_rels/document.xml.rels');
  const relationships = relationshipEntry ? xmlDocument(await relationshipEntry.async('string')) : null;
  const targets = new Map<string, string>();
  if (relationships) for (const rel of elements(relationships, 'Relationship')) {
    if (rel.getAttribute('TargetMode') !== 'External') targets.set(rel.getAttribute('Id') || '', resolvePackagePath('word/document.xml', rel.getAttribute('Target') || ''));
  }
  for (const kind of ['header', 'footer'] as const) {
    const refs = section ? elements(section, `${kind}Reference`) : [];
    const reference = refs.find(ref => wordAttribute(ref, 'type') === 'default') || refs[0];
    const target = targets.get(reference?.getAttribute('r:id') || '');
    const part = target ? zip.file(target) : null;
    if (!part) continue;
    const xml = xmlDocument(await part.async('string'));
    const fieldText = elements(xml, 'instrText').map(field => field.textContent || '').join(' ');
    const simpleFields = elements(xml, 'fldSimple').map(field => wordAttribute(field, 'instr')).join(' ');
    if (/\bPAGE\b/.test(`${fieldText} ${simpleFields}`)) pageSetup.pageNumbers = true;
    // Cached field results are omitted so a page-number field does not become static footer text.
    const parts: string[] = [];
    for (const paragraph of elements(xml, 'p')) {
      let field = false;
      let text = '';
      for (const node of Array.from(paragraph.getElementsByTagName('*'))) {
        if (node.localName === 'fldChar') field = wordAttribute(node, 'fldCharType') !== 'end';
        if (node.localName === 't' && !field && !node.closest('fldSimple')) text += node.textContent || '';
      }
      if (text.trim()) parts.push(text.trim());
    }
    pageSetup[kind] = parts.join('\n');
  }
  const quotes = new Map<string, string>();
  const active = new Set<string>();
  for (const node of Array.from(document.getElementsByTagName('*'))) {
    if (node.localName === 'commentRangeStart') { const id = wordAttribute(node, 'id'); active.add(id); quotes.set(id, quotes.get(id) || ''); }
    else if (node.localName === 'commentRangeEnd') active.delete(wordAttribute(node, 'id'));
    else if (node.localName === 't') for (const id of active) quotes.set(id, `${quotes.get(id) || ''}${node.textContent || ''}`);
  }
  const extendedEntry = zip.file('word/commentsExtended.xml');
  const resolvedParagraphs = new Set<string>();
  if (extendedEntry) for (const comment of elements(xmlDocument(await extendedEntry.async('string')), 'commentEx')) {
    if (['1', 'true'].includes(comment.getAttribute('w15:done') || '')) resolvedParagraphs.add(comment.getAttribute('w15:paraId') || '');
  }
  const comments: DocumentComment[] = [];
  const commentsEntry = zip.file('word/comments.xml');
  if (commentsEntry) for (const comment of elements(xmlDocument(await commentsEntry.async('string')), 'comment')) {
    const id = wordAttribute(comment, 'id');
    const paragraphs = elements(comment, 'p');
    comments.push({ id: crypto.randomUUID(), text: paragraphs.map(paragraph => elements(paragraph, 't').map(text => text.textContent || '').join('')).join('\n'), quote: quotes.get(id) || '', createdAt: Date.parse(wordAttribute(comment, 'date')) || Date.now(), resolved: paragraphs.some(paragraph => resolvedParagraphs.has(paragraph.getAttribute('w14:paraId') || '')) });
  }
  const warnings: string[] = [];
  if (sections.length > 1) warnings.push('The last section’s page setup is applied throughout this document.');
  if (elements(document, 'ins').length || elements(document, 'del').length) warnings.push('Tracked changes are imported as current document text.');
  if (Object.keys(zip.files).some(path => /^word\/charts\//.test(path))) warnings.push('Embedded charts are not editable.');
  return { pageSetup, comments, warning: warnings.length ? warnings.join(' ') : undefined };
}

async function importPresentation(buffer: ArrayBuffer): Promise<{ slides: SlideData[]; warning?: string }> {
  const { default: JSZip } = await import('jszip');
  const zip = await JSZip.loadAsync(buffer);
  const presentationEntry = zip.file('ppt/presentation.xml');
  if (!presentationEntry) throw new Error('This file is not a valid PowerPoint presentation.');

  const presentation = xmlDocument(await presentationEntry.async('string'));
  const relationshipEntry = zip.file('ppt/_rels/presentation.xml.rels');
  const relationships = relationshipEntry ? xmlDocument(await relationshipEntry.async('string')) : null;
  const targets = new Map<string, string>();
  if (relationships) for (const rel of elements(relationships, 'Relationship')) {
    if (rel.getAttribute('TargetMode') === 'External') continue;
    targets.set(rel.getAttribute('Id') || '', resolvePackagePath('ppt/presentation.xml', rel.getAttribute('Target') || ''));
  }
  let slidePaths = elements(presentation, 'sldId').map(slide => {
    const relationId = slide.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || slide.getAttribute('r:id') || '';
    return targets.get(relationId) || '';
  }).filter(Boolean);
  if (!slidePaths.length) slidePaths = Object.keys(zip.files).filter(path => /^ppt\/slides\/slide\d+\.xml$/.test(path)).sort((a, b) => Number(a.match(/slide(\d+)\.xml$/)?.[1]) - Number(b.match(/slide(\d+)\.xml$/)?.[1]));
  if (!slidePaths.length) throw new Error('This presentation does not contain any slides.');
  if (slidePaths.length > 300) throw new Error('Please import a presentation with 300 slides or fewer.');

  const themeColors: Record<string, string> = { lt1: 'FFFFFF', dk1: '172B4D', lt2: 'F3F4F6', dk2: '27364B', accent1: '526D4E' };
  const themeFile = zip.file('ppt/theme/theme1.xml');
  if (themeFile) {
    const theme = xmlDocument(await themeFile.async('string'));
    const scheme = elements(theme, 'clrScheme')[0];
    if (scheme) for (const entry of Array.from(scheme.children)) {
      const color = elements(entry, 'srgbClr')[0]?.getAttribute('val') || elements(entry, 'sysClr')[0]?.getAttribute('lastClr');
      if (color && /^[a-f0-9]{6}$/i.test(color)) themeColors[entry.localName] = color;
    }
  }
  themeColors.bg1 = themeColors.lt1;
  themeColors.bg2 = themeColors.lt2;
  themeColors.tx1 = themeColors.dk1;
  themeColors.tx2 = themeColors.dk2;
  const pageSize = elements(presentation, 'sldSz')[0];
  const pageWidth = Number(pageSize?.getAttribute('cx')) || 12192000;
  const pageHeight = Number(pageSize?.getAttribute('cy')) || 6858000;
  const warnings = new Set<string>();
  const readColor = (node: Element | undefined): string | undefined => {
    if (!node) return undefined;
    const direct = elements(node, 'srgbClr')[0]?.getAttribute('val');
    if (direct && /^[a-f0-9]{6}$/i.test(direct)) return `#${direct}`;
    const theme = elements(node, 'schemeClr')[0]?.getAttribute('val');
    return theme && themeColors[theme] ? `#${themeColors[theme]}` : undefined;
  };

  const backgroundFor = async (path: string, xml: Document, depth = 0): Promise<string> => {
    const background = elements(xml, 'bg')[0];
    if (background) {
      const direct = elements(background, 'srgbClr')[0]?.getAttribute('val');
      if (direct && /^[a-f0-9]{6}$/i.test(direct)) return `#${direct}`;
      const themed = elements(background, 'schemeClr')[0]?.getAttribute('val');
      if (themed && themeColors[themed]) return `#${themeColors[themed]}`;
    }
    if (depth >= 2) return '#ffffff';
    const relationshipPath = `${path.slice(0, path.lastIndexOf('/') + 1)}_rels/${path.slice(path.lastIndexOf('/') + 1)}.rels`;
    const relationshipFile = zip.file(relationshipPath);
    if (relationshipFile) {
      const relDoc = xmlDocument(await relationshipFile.async('string'));
      const parent = elements(relDoc, 'Relationship').find(rel => /\/(slideLayout|slideMaster)$/.test(rel.getAttribute('Type') || '') && rel.getAttribute('TargetMode') !== 'External');
      if (parent) {
        const parentPath = resolvePackagePath(path, parent.getAttribute('Target') || '');
        const parentFile = zip.file(parentPath);
        if (parentFile) return backgroundFor(parentPath, xmlDocument(await parentFile.async('string')), depth + 1);
      }
    }
    return '#ffffff';
  };

  const slides: SlideData[] = [];
  for (const [index, path] of slidePaths.entries()) {
    const file = zip.file(path);
    if (!file) throw new Error(`Slide ${index + 1} is missing from this presentation.`);
    const xml = xmlDocument(await file.async('string'));
    const relationshipPath = `${path.slice(0, path.lastIndexOf('/') + 1)}_rels/${path.slice(path.lastIndexOf('/') + 1)}.rels`;
    const relationshipFile = zip.file(relationshipPath);
    const slideRels = relationshipFile ? elements(xmlDocument(await relationshipFile.async('string')), 'Relationship') : [];
    const relationships = new Map(slideRels.filter(rel => rel.getAttribute('TargetMode') !== 'External').map(rel => [rel.getAttribute('Id') || '', resolvePackagePath(path, rel.getAttribute('Target') || '')]));
    let metadata: Partial<SlideData> = {};
    const metadataName = elements(xml, 'cNvPr').map(el => el.getAttribute('name') || '').find(name => name.startsWith('folio-meta-'));
    if (metadataName) {
      try { metadata = JSON.parse(decodeURIComponent(metadataName.slice('folio-meta-'.length))); } catch { /* Import native content if an optional Folio marker is invalid. */ }
    }
    const folio = !!metadataName || elements(xml, 'cNvPr').some(el => /^folio-title-/.test(el.getAttribute('name') || ''));
    const textBlocks = [...elements(xml, 'sp'), ...elements(xml, 'graphicFrame')].map(shape => {
      const text = elements(shape, 'p').map(paragraph => {
        return Array.from(paragraph.childNodes).map(node => {
          if (node.nodeType !== Node.ELEMENT_NODE) return '';
          const element = node as Element;
          return element.localName === 'br' ? '\n' : elements(element, 't').map(run => run.textContent || '').join('');
        }).join('');
      }).join('\n').trim();
      return { text, type: elements(shape, 'ph')[0]?.getAttribute('type') || '', name: elements(shape, 'cNvPr')[0]?.getAttribute('name') || '' };
    }).filter(block => (block.text || /^folio-(title-|body|subtitle)/.test(block.name)) && !block.name.startsWith('folio-element-') && !block.name.startsWith('folio-meta-') && block.name !== 'folio-page-number' && !['sldNum', 'ftr', 'dt'].includes(block.type));
    const explicitTitle = textBlocks.find(block => ['title', 'ctrTitle'].includes(block.type) || /^folio-title-(title|content|split)$/.test(block.name));
    const title = explicitTitle || textBlocks[0];
    const subtitle = textBlocks.find(block => (block.type === 'subTitle' || block.name === 'folio-subtitle') && block !== title);
    const body = textBlocks.filter(block => block !== title && block !== subtitle).map(block => block.text).join('\n\n');
    const originalLayout = title?.name.match(/^folio-title-(title|content|split)$/)?.[1] as SlideData['layout'] | undefined;
    const slideElements: SlideElement[] = [];
    const shapeTree = elements(xml, 'spTree')[0];
    for (const shape of shapeTree ? Array.from(shapeTree.children) : []) {
      if (!['sp', 'pic', 'cxnSp'].includes(shape.localName)) {
        if (shape.localName === 'grpSp') warnings.add('Grouped objects are not imported.');
        if (shape.localName === 'graphicFrame') warnings.add('Charts, tables, and SmartArt are not imported as editable objects.');
        continue;
      }
      const name = elements(shape, 'cNvPr')[0]?.getAttribute('name') || '';
      if (name.startsWith('folio-meta-') || name.startsWith('folio-decoration-') || name === 'folio-page-number' || (folio && /^folio-(title-|subtitle|body)/.test(name))) continue;
      const placeholder = elements(shape, 'ph')[0]?.getAttribute('type');
      if (placeholder && ['sldNum', 'ftr', 'dt'].includes(placeholder)) continue;
      const transform = elements(shape, 'xfrm')[0];
      const offset = transform ? elements(transform, 'off')[0] : undefined;
      const extent = transform ? elements(transform, 'ext')[0] : undefined;
      const properties = elements(shape, 'spPr')[0];
      const geometry = properties ? elements(properties, 'prstGeom')[0]?.getAttribute('prst') : '';
      const paragraphs = elements(shape, 'p');
      const text = paragraphs.map(paragraph => elements(paragraph, 't').map(run => run.textContent || '').join('')).join('\n');
      const run = elements(shape, 'rPr')[0] || elements(shape, 'defRPr')[0];
      const paragraph = paragraphs[0] ? elements(paragraphs[0], 'pPr')[0] : undefined;
      const solidFill = properties ? Array.from(properties.children).find(child => child.localName === 'solidFill') : undefined;
      const line = properties ? Array.from(properties.children).find(child => child.localName === 'ln') : undefined;
      const element: SlideElement = {
        id: name.startsWith('folio-element-') ? name.slice('folio-element-'.length) || crypto.randomUUID() : crypto.randomUUID(),
        type: shape.localName === 'pic' ? 'image' : text || elements(shape, 'cNvSpPr')[0]?.getAttribute('txBox') === '1' ? 'text' : 'shape',
        x: offset ? Number(offset.getAttribute('x')) / pageWidth * 100 : 6,
        y: offset ? Number(offset.getAttribute('y')) / pageHeight * 100 : 10 + slideElements.length * 15,
        width: extent ? Math.max(0.1, Number(extent.getAttribute('cx')) / pageWidth * 100) : 88,
        height: extent ? Math.max(0.1, Number(extent.getAttribute('cy')) / pageHeight * 100) : 15,
        rotation: transform ? Number(transform.getAttribute('rot') || 0) / 60000 : 0,
      };
      if (element.type === 'image') {
        const blip = elements(shape, 'blip')[0];
        const imagePath = relationships.get(blip?.getAttribute('r:embed') || '');
        const image = imagePath ? zip.file(imagePath) : null;
        if (!image || !imagePath) { warnings.add('Linked images could not be imported.'); continue; }
        const extension = imagePath.split('.').pop()?.toLowerCase();
        const mime = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp' }[extension || ''];
        if (!mime) { warnings.add('Some images use a format the browser cannot display.'); continue; }
        element.src = `data:${mime};base64,${await image.async('base64')}`;
      } else if (element.type === 'text') {
        element.text = text;
        element.fontSize = Number(run?.getAttribute('sz') || 1800) / 100;
        element.bold = run?.getAttribute('b') === '1';
        element.color = readColor(run) || '#172B4D';
        element.align = paragraph?.getAttribute('algn') === 'ctr' ? 'center' : paragraph?.getAttribute('algn') === 'r' ? 'right' : 'left';
        if (solidFill && !elements(solidFill, 'alpha').some(alpha => alpha.getAttribute('val') === '0')) element.fill = readColor(solidFill);
      } else {
        element.shape = geometry === 'ellipse' ? 'ellipse' : geometry === 'line' || shape.localName === 'cxnSp' ? 'line' : 'rectangle';
        element.fill = solidFill && elements(solidFill, 'alpha').some(alpha => alpha.getAttribute('val') === '0') ? 'transparent' : readColor(solidFill) || (properties && Array.from(properties.children).some(child => child.localName === 'noFill') ? 'transparent' : '#526D4E');
        element.color = readColor(line) || '#526D4E';
        if (geometry && !['rect', 'ellipse', 'line', 'roundRect'].includes(geometry)) warnings.add('Some shapes were approximated with rectangles.');
      }
      slideElements.push(element);
    }
    let notes = '';
    const notesRel = slideRels.find(rel => /\/notesSlide$/.test(rel.getAttribute('Type') || ''));
    const notesPath = notesRel ? resolvePackagePath(path, notesRel.getAttribute('Target') || '') : '';
    const notesEntry = notesPath ? zip.file(notesPath) : null;
    if (notesEntry) {
      const notesXml = xmlDocument(await notesEntry.async('string'));
      const blocks = elements(notesXml, 'sp').filter(shape => elements(shape, 'ph')[0]?.getAttribute('type') === 'body');
      notes = blocks.map(shape => elements(shape, 'p').map(paragraph => elements(paragraph, 't').map(text => text.textContent || '').join('')).join('\n')).join('\n');
    }
    if (elements(xml, 'timing').length) warnings.add('Object animations are not imported.');
    const nativeTransition = elements(xml, 'transition')[0];
    const transition = metadata.transition || (nativeTransition && elements(nativeTransition, 'fade').length ? 'fade' : nativeTransition && elements(nativeTransition, 'push').length ? 'slide' : 'none');
    slides.push({
      id: crypto.randomUUID(),
      title: title ? title.text : metadata.layout === 'blank' && typeof metadata.title === 'string' ? metadata.title : `Slide ${index + 1}`,
      body: folio ? metadata.layout === 'blank' && typeof metadata.body === 'string' ? metadata.body : body : '',
      subtitle: folio ? metadata.layout === 'blank' && typeof metadata.subtitle === 'string' ? metadata.subtitle : subtitle?.text || '' : '',
      background: await backgroundFor(path, xml),
      layout: ['title', 'content', 'split', 'blank'].includes(metadata.layout || '') ? metadata.layout! : folio ? originalLayout || (subtitle && !body ? 'title' : 'content') : 'blank',
      notes,
      hidden: xml.documentElement.getAttribute('show') === '0' || !!metadata.hidden,
      transition,
      titleSize: typeof metadata.titleSize === 'number' ? metadata.titleSize : undefined,
      bodySize: typeof metadata.bodySize === 'number' ? metadata.bodySize : undefined,
      textColor: typeof metadata.textColor === 'string' ? metadata.textColor : undefined,
      fontFamily: typeof metadata.fontFamily === 'string' ? metadata.fontFamily : undefined,
      elements: slideElements,
    });
  }
  return { slides, warning: warnings.size ? [...warnings].join(' ') : undefined };
}

export async function importOfficeFile(file: File): Promise<ImportedOfficeFile> {
  const extension = file.name.split('.').pop()?.toLowerCase() || '';
  if (!SUPPORTED_EXTENSIONS.has(extension)) throw new Error('Supported files: DOCX, XLSX, XLS, PPTX, CSV, TXT, HTML, and Markdown. Save legacy DOC or PPT files in a modern format first.');
  if (file.size > MAX_FILE_SIZE) throw new Error('Please choose a file smaller than 30 MB.');
  if (!file.size) throw new Error('This file is empty. Choose a file with some content.');
  const name = baseName(file.name);
  try {
    if (extension === 'docx') {
      const mammoth = await import('mammoth');
      const buffer = await file.arrayBuffer();
      const [result, metadata] = await Promise.all([mammoth.convertToHtml({ arrayBuffer: buffer }, { externalFileAccess: false }), documentMetadata(buffer)]);
      let content = cleanHTML(result.value);
      // Anchor a comment only when its exact quote occurs once; repeated text must not get a misleading anchor.
      const parsed = new DOMParser().parseFromString(content, 'text/html');
      for (const comment of metadata.comments) {
        if (!comment.quote || comment.resolved) continue;
        const walker = parsed.createTreeWalker(parsed.body, NodeFilter.SHOW_TEXT);
        const nodes: Text[] = [];
        while (walker.nextNode()) nodes.push(walker.currentNode as Text);
        const allText = nodes.map(node => node.data).join('');
        const start = allText.indexOf(comment.quote);
        if (start < 0 || allText.indexOf(comment.quote, start + 1) >= 0) continue;
        let cursor = 0;
        for (const node of nodes) {
          const end = cursor + node.length;
          const from = Math.max(0, start - cursor);
          const to = Math.min(node.length, start + comment.quote.length - cursor);
          if (to > from && !node.parentElement?.closest('[data-comment-id]')) {
            const range = parsed.createRange();
            range.setStart(node, from);
            range.setEnd(node, to);
            const span = parsed.createElement('span');
            span.setAttribute('data-comment-id', comment.id);
            range.surroundContents(span);
          }
          cursor = end;
        }
      }
      content = parsed.body.innerHTML;
      return {
        name, kind: 'document', content, ...metadata,
      };
    }
    if (['txt', 'html', 'htm', 'md'].includes(extension)) {
      const text = await file.text();
      return {
        name, kind: 'document',
        content: extension === 'txt' ? textAsHTML(text) : extension === 'md' ? markdownAsHTML(text) : cleanHTML(text),
        ...(extension === 'md' ? { warning: 'Imported basic Markdown formatting. Advanced Markdown syntax may appear as text.' } : {}),
      };
    }
    if (['xlsx', 'xls', 'csv'].includes(extension)) {
      const { importWorkbook } = await import('./workbookIO');
      return { name, kind: 'spreadsheet', ...await importWorkbook(file) };
    }
    const presentation = await importPresentation(await file.arrayBuffer());
    return { name, kind: 'presentation', content: presentation.slides, warning: presentation.warning };
  } catch (error) {
    const message = error instanceof Error ? error.message : '';
    if (/^(This |Please |Slide )/.test(message)) throw error;
    throw new Error(`Could not open “${file.name}”. The file may be damaged, password-protected, or saved in an unsupported format.`);
  }
}

export async function downloadOfficeFile(file: OfficeFile): Promise<void> {
  if (file.kind === 'document') {
    const { exportDocument } = await import('../editors/documentExport');
    await exportDocument(typeof file.content === 'string' ? file.content : '<p></p>', file.name, 'docx', { ...file.pageSetup, comments: file.comments });
  } else if (file.kind === 'spreadsheet') {
    const { exportWorkbook } = await import('./workbookIO');
    await exportWorkbook(file.content as SheetContent, file.name, 'xlsx');
  }
  else if (file.kind === 'presentation') {
    if (!Array.isArray(file.content) || !file.content.length) throw new Error('Add a slide before exporting this presentation.');
    const { exportPresentationPptx } = await import('../editors/PresentationEditor');
    await exportPresentationPptx(file.content as SlideData[], safeFilename(file.name, 'pptx').replace(/\.pptx$/i, ''));
  }
  else throw new Error('This file type cannot be exported.');
}
