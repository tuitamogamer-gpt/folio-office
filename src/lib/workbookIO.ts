import type { CellStyle, SheetContent, SheetTab } from '../types';
import { normalizeWorkbook, packWorkbook } from './spreadsheetModel';
import { calculateCells } from '../editors/spreadsheetFormula';

const ADDRESS = /^[A-Z]+[1-9]\d*$/;
const MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const tags = (node: Document | Element, name: string) => Array.from(node.getElementsByTagNameNS('*', name));
const xml = (value: string) => new DOMParser().parseFromString(value, 'application/xml');
const color = (value?: string | null) => value && /^(?:[a-f\d]{2})?[a-f\d]{6}$/i.test(value) ? `#${value.slice(-6)}` : undefined;

// These supported functions postdate the original OOXML function set. Excel
// expects the future-function prefix on disk, while users can type normal names.
const FUTURE_FUNCTIONS = new Set(['XLOOKUP', 'CONCAT']);
function officeFormula(formula: string): string {
  let output = '';
  for (let position = 0; position < formula.length;) {
    const start = position;
    const quote = formula[position];
    if (quote === '"' || quote === "'") {
      position++;
      while (position < formula.length) {
        if (formula[position++] === quote) {
          if (formula[position] === quote) position++;
          else break;
        }
      }
      output += formula.slice(start, position);
      continue;
    }
    const name = formula.slice(position).match(/^[\p{L}_$][\p{L}\p{N}_.$]*/u)?.[0];
    if (name) {
      position += name.length;
      // Dotted names already include their namespace; do not qualify them again.
      output += FUTURE_FUNCTIONS.has(name.toUpperCase()) && /^\s*\(/.test(formula.slice(position)) ? `_xlfn.${name}` : name;
    } else output += formula[position++];
  }
  return output;
}

function numberFormat(style: CellStyle): string {
  const places = Math.max(0, Math.min(10, style.decimals ?? (style.numberFormat === 'percentage' ? 0 : 2)));
  const decimal = places ? `.${'0'.repeat(places)}` : '';
  switch (style.numberFormat) {
    case 'number': return `#,##0${decimal}`;
    case 'currency': return `€#,##0${decimal}`;
    case 'percentage': return `0${decimal}%`;
    case 'date': return 'yyyy-mm-dd';
    default: return 'General';
  }
}

function readNumberFormat(format: string | undefined): Partial<CellStyle> {
  if (!format || format === 'General' || format === '@') return {};
  const decimal = format.match(/\.([0#]+)/)?.[1].length ?? 0;
  if (/[ymd]/i.test(format.replace(/"[^"]*"/g, ''))) return { numberFormat: 'date' };
  if (format.includes('%')) return { numberFormat: 'percentage', decimals: decimal };
  if (/[$€£¥]|\[\$/.test(format)) return { numberFormat: 'currency', decimals: decimal };
  if (/[0#]/.test(format)) return { numberFormat: 'number', decimals: decimal };
  return {};
}

async function xlsxMetadata(buffer: ArrayBuffer) {
  const { default: JSZip } = await import('jszip');
  const zip = await JSZip.loadAsync(buffer);
  const styleEntry = zip.file('xl/styles.xml');
  const styles: CellStyle[] = [];
  const styleDoc = styleEntry ? xml(await styleEntry.async('string')) : null;
  const themeEntry = zip.file('xl/theme/theme1.xml');
  const themeColors: string[] = ['FFFFFF', '000000', 'EEECE1', '1F497D', '4F81BD', 'C0504D', '9BBB59', '8064A2', '4BACC6', 'F79646'];
  if (themeEntry) {
    const scheme = tags(xml(await themeEntry.async('string')), 'clrScheme')[0];
    ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'].forEach((name, index) => {
      const entry = scheme ? Array.from(scheme.children).find(child => child.localName === name) : undefined;
      const value = entry ? tags(entry, 'srgbClr')[0]?.getAttribute('val') || tags(entry, 'sysClr')[0]?.getAttribute('lastClr') : '';
      if (value) themeColors[index] = value;
    });
  }
  const indexedColors = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080'];
  const readStyleColor = (element: Element | undefined) => {
    if (!element) return undefined;
    const raw = element.getAttribute('rgb') || (element.hasAttribute('theme') ? themeColors[Number(element.getAttribute('theme'))] : element.hasAttribute('indexed') ? indexedColors[Number(element.getAttribute('indexed'))] : undefined);
    const rgb = color(raw);
    const tint = Number(element.getAttribute('tint') || 0);
    if (!rgb || !tint) return rgb;
    return `#${[1, 3, 5].map(start => {
      const channel = parseInt(rgb.slice(start, start + 2), 16);
      return Math.round(Math.max(0, Math.min(255, tint < 0 ? channel * (1 + tint) : channel * (1 - tint) + 255 * tint))).toString(16).padStart(2, '0');
    }).join('')}`;
  };
  const builtin: Record<number, string> = { 0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 5: '$#,##0', 6: '$#,##0', 7: '$#,##0.00', 8: '$#,##0.00', 9: '0%', 10: '0.00%', 14: 'mm-dd-yy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 22: 'm/d/yy h:mm' };
  if (styleDoc) {
    const fonts = tags(styleDoc, 'fonts')[0]?.children;
    const fills = tags(styleDoc, 'fills')[0]?.children;
    for (const format of tags(styleDoc, 'numFmt')) builtin[Number(format.getAttribute('numFmtId'))] = format.getAttribute('formatCode') || '';
    for (const xf of Array.from(tags(styleDoc, 'cellXfs')[0]?.children || [])) {
      const font = fonts?.[Number(xf.getAttribute('fontId') || 0)];
      const fill = fills?.[Number(xf.getAttribute('fillId') || 0)];
      const alignment = tags(xf, 'alignment')[0]?.getAttribute('horizontal');
      const fontColor = font ? readStyleColor(tags(font, 'color')[0]) : undefined;
      const background = fill ? readStyleColor(tags(fill, 'fgColor')[0]) : undefined;
      const style: CellStyle = { ...readNumberFormat(builtin[Number(xf.getAttribute('numFmtId') || 0)]) };
      if (font) for (const [tag, property] of [['b', 'bold'], ['i', 'italic'], ['u', 'underline']] as const) {
        const el = tags(font, tag)[0];
        if (el && !['0', 'false', 'none'].includes(el.getAttribute('val') || '')) style[property] = true;
      }
      if (fontColor) style.color = fontColor;
      if (background) style.background = background;
      if (alignment === 'left' || alignment === 'center' || alignment === 'right') style.align = alignment;
      styles.push(style);
    }
  }
  const workbookEntry = zip.file('xl/workbook.xml');
  const workbook = workbookEntry ? xml(await workbookEntry.async('string')) : null;
  const relEntry = zip.file('xl/_rels/workbook.xml.rels');
  const rels = relEntry ? xml(await relEntry.async('string')) : null;
  const relationPaths = new Map(rels ? tags(rels, 'Relationship').map(rel => {
    const target = rel.getAttribute('Target') || '';
    return [rel.getAttribute('Id') || '', target.startsWith('/') ? target.slice(1) : `xl/${target}`];
  }) : []);
  const sheets = [];
  for (const sheet of workbook ? tags(workbook, 'sheet') : []) {
    const path = relationPaths.get(sheet.getAttribute('r:id') || '') || '';
    const entry = zip.file(path);
    const data = entry ? xml(await entry.async('string')) : null;
    const cellStyles: Record<string, CellStyle> = {};
    if (data) for (const cell of tags(data, 'c')) {
      const style = styles[Number(cell.getAttribute('s') || 0)];
      const address = cell.getAttribute('r') || '';
      if (ADDRESS.test(address) && style && Object.keys(style).length) cellStyles[address] = { ...style };
    }
    const pane = data ? tags(data, 'pane')[0] : null;
    sheets.push({ name: sheet.getAttribute('name') || '', styles: cellStyles, freezeRows: pane?.getAttribute('state')?.startsWith('frozen') ? Number(pane.getAttribute('ySplit') || 0) : 0, hasMerges: !!data && tags(data, 'mergeCell').length > 0, conditional: !!data && tags(data, 'conditionalFormatting').length > 0 });
  }
  return { sheets, activeTab: workbook ? Number(tags(workbook, 'workbookView')[0]?.getAttribute('activeTab') || 0) : 0, charts: Object.keys(zip.files).some(path => /^xl\/charts\/chart\d+\.xml$/.test(path)), macros: !!zip.file('xl/vbaProject.bin') };
}

export async function importWorkbook(file: File): Promise<{ content: SheetContent; warning?: string }> {
  const XLSX = await import('xlsx-js-style');
  const extension = file.name.split('.').pop()?.toLowerCase();
  const buffer = extension === 'csv' ? null : await file.arrayBuffer();
  const workbook = extension === 'csv'
    ? XLSX.read(await file.text(), { type: 'string', raw: true, cellFormula: true })
    : XLSX.read(buffer, { type: 'array', cellFormula: true, cellDates: false, sheetStubs: true, cellStyles: true, cellNF: true, xlfn: true });
  if (!workbook.SheetNames.length) throw new Error('This workbook does not contain a worksheet.');
  if (workbook.SheetNames.length > 100) throw new Error('Please import a workbook with 100 worksheets or fewer.');
  const metadata = extension === 'xlsx' && buffer ? await xlsxMetadata(buffer) : null;
  let populatedCells = 0;
  let outsideGrid = false;
  const sheets: SheetTab[] = workbook.SheetNames.map(name => {
    const source = workbook.Sheets[name];
    const meta = metadata?.sheets.find(sheet => sheet.name === name);
    const cells: Record<string, string> = {};
    const styles = { ...meta?.styles };
    const columnWidths: Record<string, number> = {};
    for (const [address, cell] of Object.entries(source)) {
      if (!ADDRESS.test(address) || !cell || (cell.v == null && !cell.f && !cell.s)) continue;
      if (++populatedCells > 100_000) throw new Error('Please import a workbook with 100,000 populated cells or fewer.');
      if (cell.f) cells[address] = `=${cell.f}`;
      else if (cell.v != null) cells[address] = cell.t === 'b' ? (cell.v ? 'TRUE' : 'FALSE') : cell.t === 'e' ? cell.w || '#VALUE!' : String(cell.v);
      if (!styles[address] && cell.z) {
        const style = readNumberFormat(String(cell.z));
        if (Object.keys(style).length) styles[address] = style;
      }
      const position = XLSX.utils.decode_cell(address);
      if (position.r >= 2_000 || position.c >= 100) outsideGrid = true;
    }
    source['!cols']?.forEach((column, index) => {
      if (column) columnWidths[XLSX.utils.encode_col(index)] = Math.max(24, Math.min(600, column.wpx ?? (column.wch != null ? column.wch * 7 + 5 : 112)));
    });
    return { id: crypto.randomUUID(), name, cells, styles, columnWidths, freezeRows: meta?.freezeRows || 0 };
  });
  const warnings: string[] = [];
  if (metadata?.charts) warnings.push('Embedded charts are not imported.');
  if (metadata?.macros || workbook.vbaraw) warnings.push('Macros are not imported.');
  if (metadata?.sheets.some(sheet => sheet.hasMerges)) warnings.push('Merged cells are imported as individual cells.');
  if (metadata?.sheets.some(sheet => sheet.conditional)) warnings.push('Conditional formatting rules are not imported.');
  if (outsideGrid) warnings.push('Cells beyond 2,000 rows or 100 columns are retained for export but are outside the editable grid.');
  return { content: packWorkbook(sheets, sheets[Math.min(metadata?.activeTab || 0, sheets.length - 1)].id), warning: warnings.length ? warnings.join(' ') : undefined };
}

/** Build a real Office archive. Kept separate from download for compatibility tests. */
export async function createWorkbookBlob(content: SheetContent, format: 'xlsx' | 'csv' = 'xlsx'): Promise<Blob> {
  const XLSX = await import('xlsx-js-style');
  const normalized = normalizeWorkbook(content);
  const { sheets, activeSheetId } = normalized;
  const workbook = XLSX.utils.book_new();
  const usedNames = new Set<string>();
  for (const tab of sheets) {
    if (format === 'csv' && tab.id !== activeSheetId) continue;
    const evaluate = calculateCells(tab.cells, { sheets, currentSheetId: tab.id });
    const sheet: import('xlsx-js-style').WorkSheet = {};
    let maxRow = 0;
    let maxColumn = 0;
    for (const address of new Set([...Object.keys(tab.cells), ...Object.keys(tab.styles || {})])) {
      if (!ADDRESS.test(address)) continue;
      const position = XLSX.utils.decode_cell(address);
      if (position.r >= 1_048_576 || position.c >= 16_384) continue;
      maxRow = Math.max(maxRow, position.r);
      maxColumn = Math.max(maxColumn, position.c);
      const value = String(tab.cells[address] ?? '');
      let cell: import('xlsx-js-style').CellObject;
      if (value.startsWith('=') && value.length > 1) {
        const computed = evaluate(address);
        const errors: Record<string, number> = { '#NULL!': 0, '#DIV/0!': 7, '#VALUE!': 15, '#REF!': 23, '#NAME?': 29, '#NUM!': 36, '#N/A': 42 };
        cell = typeof computed === 'number' ? { t: 'n', v: computed } : computed in errors ? { t: 'e', v: errors[computed], w: computed } : ['TRUE', 'FALSE'].includes(computed) ? { t: 'b', v: computed === 'TRUE' } : { t: 's', v: computed };
        cell.f = officeFormula(value.slice(1));
      } else if (/^(true|false)$/i.test(value)) cell = { t: 'b', v: value.toLowerCase() === 'true' };
      else if (/^[+-]?(?:0|[1-9]\d*)(?:\.\d*)?(?:e[+-]?\d+)?$/i.test(value) && Number.isFinite(Number(value)) && value.replace(/\D/g, '').length <= 15) cell = { t: 'n', v: Number(value) };
      else cell = { t: 's', v: value };
      const style = tab.styles?.[address];
      if (style) {
        cell.s = { font: { bold: !!style.bold, italic: !!style.italic, underline: !!style.underline, ...(style.color ? { color: { rgb: style.color.replace('#', '') } } : {}) }, ...(style.background ? { fill: { patternType: 'solid', fgColor: { rgb: style.background.replace('#', '') } } } : {}), ...(style.align ? { alignment: { horizontal: style.align } } : {}), numFmt: numberFormat(style) };
        cell.z = numberFormat(style);
      }
      sheet[address] = cell;
    }
    sheet['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxRow, c: maxColumn } });
    for (const column of Object.keys(tab.columnWidths || {})) if (/^[A-Z]+$/.test(column)) maxColumn = Math.max(maxColumn, XLSX.utils.decode_col(column));
    sheet['!cols'] = Array.from({ length: Math.min(maxColumn + 1, 16_384) }, (_, index) => ({ wpx: tab.columnWidths?.[XLSX.utils.encode_col(index)] || 112 }));
    const base = (tab.name || 'Sheet').replace(/[\\/*?:\[\]]/g, '').replace(/^'+|'+$/g, '').slice(0, 31) || 'Sheet';
    let name = base;
    let suffix = 1;
    while (usedNames.has(name.toLowerCase())) { const ending = ` (${++suffix})`; name = `${base.slice(0, 31 - ending.length)}${ending}`; }
    usedNames.add(name.toLowerCase());
    XLSX.utils.book_append_sheet(workbook, sheet, name);
  }
  if (format === 'csv') return new Blob(['\uFEFF', XLSX.utils.sheet_to_csv(workbook.Sheets[workbook.SheetNames[0]])], { type: 'text/csv;charset=utf-8' });
  const output = XLSX.write(workbook, { bookType: 'xlsx', type: 'array', compression: true, cellStyles: true });
  const { default: JSZip } = await import('jszip');
  const zip = await JSZip.loadAsync(output);
  for (const [index, tab] of sheets.entries()) {
    if (!tab.freezeRows) continue;
    const path = `xl/worksheets/sheet${index + 1}.xml`;
    const entry = zip.file(path);
    if (!entry) continue;
    const count = Math.max(0, Math.min(1_048_575, Math.floor(tab.freezeRows)));
    const views = `<sheetViews><sheetView workbookViewId="0"><pane ySplit="${count}" topLeftCell="A${count + 1}" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A${count + 1}" sqref="A${count + 1}"/></sheetView></sheetViews>`;
    let source = await entry.async('string');
    source = /<sheetViews\b/.test(source) ? source.replace(/<sheetViews\b[^>]*>[\s\S]*?<\/sheetViews>/, views) : source.replace(/(<dimension\b[^>]*\/>)/, `$1${views}`);
    zip.file(path, source);
  }
  const wbEntry = zip.file('xl/workbook.xml');
  if (wbEntry) {
    const activeTab = Math.max(0, sheets.findIndex(sheet => sheet.id === activeSheetId));
    let source = await wbEntry.async('string');
    if (/<workbookView\b/.test(source)) source = source.replace(/<workbookView\b([^>]*?)(\/?)>/, (_, attrs: string, slash: string) => `<workbookView${attrs.replace(/\sactiveTab="[^"]*"/, '')} activeTab="${activeTab}"${slash}>`);
    else source = source.replace('<sheets>', `<bookViews><workbookView activeTab="${activeTab}"/></bookViews><sheets>`);
    zip.file('xl/workbook.xml', source);
  }
  return new Blob([await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' })], { type: MIME });
}

export async function exportWorkbook(content: SheetContent, name: string, format: 'xlsx' | 'csv' = 'xlsx'): Promise<void> {
  const { downloadBlob } = await import('./fileIO');
  const filename = `${name.replace(/\.(xlsx|xls|csv)$/i, '').replace(/[<>:"/\\|?*\x00-\x1f]/g, '-').trim() || 'Untitled'}.${format}`;
  downloadBlob(await createWorkbookBlob(content, format), filename);
}
