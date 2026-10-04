import DOMPurify from 'dompurify';
import type { FileKind, OfficeFile, SheetContent, SlideData } from '../types';

export interface ImportedOfficeFile {
  name: string;
  kind: FileKind;
  content: string | SheetContent | SlideData[];
  warning?: string;
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
  if (xml.getElementsByTagName('parsererror').length) throw new Error('This presentation contains invalid XML. Try saving it as a new PPTX file.');
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

async function importPresentation(buffer: ArrayBuffer): Promise<SlideData[]> {
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
    const textBlocks = [...elements(xml, 'sp'), ...elements(xml, 'graphicFrame')].map(shape => {
      const text = elements(shape, 'p').map(paragraph => {
        return Array.from(paragraph.childNodes).map(node => {
          if (node.nodeType !== Node.ELEMENT_NODE) return '';
          const element = node as Element;
          return element.localName === 'br' ? '\n' : elements(element, 't').map(run => run.textContent || '').join('');
        }).join('');
      }).join('\n').trim();
      return { text, type: elements(shape, 'ph')[0]?.getAttribute('type') || '', name: elements(shape, 'cNvPr')[0]?.getAttribute('name') || '' };
    }).filter(block => block.text && block.name !== 'folio-page-number' && !['sldNum', 'ftr', 'dt'].includes(block.type));
    const explicitTitle = textBlocks.find(block => ['title', 'ctrTitle'].includes(block.type) || /^folio-title-(title|content|split)$/.test(block.name));
    const title = explicitTitle || textBlocks[0];
    const subtitle = textBlocks.find(block => (block.type === 'subTitle' || block.name === 'folio-subtitle') && block !== title);
    const body = textBlocks.filter(block => block !== title && block !== subtitle).map(block => block.text).join('\n\n');
    const originalLayout = title?.name.match(/^folio-title-(title|content|split)$/)?.[1] as SlideData['layout'] | undefined;
    slides.push({
      id: crypto.randomUUID(),
      title: title?.text || `Slide ${index + 1}`,
      body,
      subtitle: subtitle?.text || '',
      background: await backgroundFor(path, xml),
      layout: originalLayout || (subtitle && !body ? 'title' : 'content'),
    });
  }
  return slides;
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
      const result = await mammoth.convertToHtml({ arrayBuffer: await file.arrayBuffer() }, { externalFileAccess: false });
      return {
        name, kind: 'document', content: cleanHTML(result.value),
        warning: 'Document text and basic formatting imported. Page layout, comments, tracked changes, and some embedded objects may differ.',
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
      const XLSX = await import('xlsx');
      const workbook = extension === 'csv'
        ? XLSX.read(await file.text(), { type: 'string', cellFormula: true, raw: true })
        : XLSX.read(await file.arrayBuffer(), { type: 'array', cellFormula: true, cellDates: false, sheetStubs: true });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      if (!sheetName || !sheet) throw new Error('This workbook does not contain a worksheet.');
      const cells: Record<string, string> = {};
      let populatedCells = 0;
      let hasCellsOutsideGrid = false;
      for (const [address, cell] of Object.entries(sheet)) {
        if (!/^[A-Z]+[1-9]\d*$/.test(address) || !cell) continue;
        if (cell.v == null && !cell.f) continue;
        if (++populatedCells > 100_000) throw new Error('Please import a sheet with 100,000 populated cells or fewer.');
        cells[address] = cell.f ? `=${cell.f}` : cell.v == null ? '' : String(cell.v);
        const position = XLSX.utils.decode_cell(address);
        if (position.r >= 2_000 || position.c >= 100) hasCellsOutsideGrid = true;
      }
      const warnings: string[] = [];
      if (workbook.SheetNames.length > 1) warnings.push(`Imported values and formulas from the first sheet, “${sheetName}”. Other sheets and workbook formatting are not included.`);
      else if (extension !== 'csv') warnings.push('Imported cell values and formulas. Workbook formatting, charts, and macros are not retained.');
      if (hasCellsOutsideGrid) warnings.push('Cells beyond 2,000 rows or 100 columns are retained for export but are outside the editable grid.');
      return {
        name, kind: 'spreadsheet', content: { name: sheetName, cells },
        warning: warnings.length ? warnings.join(' ') : undefined,
      };
    }
    return {
      name, kind: 'presentation', content: await importPresentation(await file.arrayBuffer()),
      warning: 'Imported slide text and solid backgrounds into editable layouts. Images, animations, charts, and original positioning are not retained.',
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : '';
    if (/^(This |Please |Slide )/.test(message)) throw error;
    throw new Error(`Could not open “${file.name}”. The file may be damaged, password-protected, or saved in an unsupported format.`);
  }
}

async function exportSpreadsheet(file: OfficeFile) {
  const XLSX = await import('xlsx');
  const { calculateCells } = await import('../editors/SpreadsheetEditor');
  const content = file.content as SheetContent;
  if (!content || !content.cells || typeof content.cells !== 'object') throw new Error('This spreadsheet has no valid cell data.');
  const evaluate = calculateCells(content.cells);
  const sheet: import('xlsx').WorkSheet = {};
  let maxRow = 0;
  let maxColumn = 0;
  for (const [address, rawValue] of Object.entries(content.cells)) {
    if (!/^[A-Z]+[1-9]\d*$/.test(address)) continue;
    const position = XLSX.utils.decode_cell(address);
    if (position.r >= 1_048_576 || position.c >= 16_384) continue;
    const value = String(rawValue ?? '');
    if (!value) continue;
    maxRow = Math.max(maxRow, position.r);
    maxColumn = Math.max(maxColumn, position.c);
    if (value.startsWith('=') && value.length > 1) {
      const computed = evaluate(address);
      sheet[address] = { t: typeof computed === 'number' ? 'n' : 's', f: value.slice(1), v: computed };
    }
    else if (/^(true|false)$/i.test(value)) sheet[address] = { t: 'b', v: value.toLowerCase() === 'true' };
    else if (/^[+-]?(?:0|[1-9]\d*)(?:\.\d*)?(?:e[+-]?\d+)?$/i.test(value) && Number.isFinite(Number(value)) && value.replace(/\D/g, '').length <= 15) sheet[address] = { t: 'n', v: Number(value) };
    else sheet[address] = { t: 's', v: value };
  }
  sheet['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxRow, c: maxColumn } });
  sheet['!cols'] = Array.from({ length: maxColumn + 1 }, () => ({ wch: 16 }));
  const workbook = XLSX.utils.book_new();
  const sheetName = (content.name || 'Sheet1').replace(/[\\/*?:\[\]]/g, '').replace(/^'+|'+$/g, '').slice(0, 31) || 'Sheet1';
  XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
  const output = XLSX.write(workbook, { bookType: 'xlsx', type: 'array', compression: true });
  downloadBlob(new Blob([output], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), safeFilename(file.name, 'xlsx'));
}

export async function downloadOfficeFile(file: OfficeFile): Promise<void> {
  if (file.kind === 'document') {
    const { exportDocument } = await import('../editors/documentExport');
    await exportDocument(typeof file.content === 'string' ? file.content : '<p></p>', file.name, 'docx', file.pageSetup);
  } else if (file.kind === 'spreadsheet') await exportSpreadsheet(file);
  else if (file.kind === 'presentation') {
    if (!Array.isArray(file.content) || !file.content.length) throw new Error('Add a slide before exporting this presentation.');
    const { exportPresentationPptx } = await import('../editors/PresentationEditor');
    await exportPresentationPptx(file.content as SlideData[], safeFilename(file.name, 'pptx').replace(/\.pptx$/i, ''));
  }
  else throw new Error('This file type cannot be exported.');
}
