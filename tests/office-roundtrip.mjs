import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import JSZip from 'jszip';
import XLSX from 'xlsx-js-style';

const browser = await chromium.launch({ headless: true, executablePath: '/usr/bin/chromium', args: ['--no-sandbox'] });
const page = await browser.newPage();
const errors = [];
page.on('pageerror', error => errors.push(error.message));
await page.goto(process.env.BASE_URL || 'http://localhost:5173');
const importBytes = async (bytes, name) => page.evaluate(async ({ bytes, name }) => {
  const { importOfficeFile } = await import('/src/lib/fileIO.ts');
  return importOfficeFile(new File([new Uint8Array(bytes)], name));
}, { bytes: Array.from(bytes), name });

try {
  const workbook = { name: 'Totals', cells: {}, activeSheetId: 'totals', sheets: [
    { id: 'source', name: 'Source data', cells: { A1: 'Item', B1: 'Amount', A2: 'North', B2: '14', A3: 'South', B3: '21', C3: '=SUM(B2:B3)' }, styles: { A1: { bold: true, italic: true, underline: true, background: '#DAE8FF', color: '#AB1234', align: 'center' }, B2: { numberFormat: 'currency', decimals: 2 }, B3: { numberFormat: 'percentage', decimals: 1 }, D4: { background: '#FFEEDD' } }, columnWidths: { A: 190, B: 130 }, freezeRows: 1 },
    { id: 'totals', name: 'Totals', cells: { A1: 'Summary', B2: "='Source data'!B2+'Source data'!B3", C2: "=SUM('Source data'!B2:B3)", D2: '=IF(B2=35,"ok","bad")', A3: 'Active only' }, styles: { B2: { bold: true, numberFormat: 'number', decimals: 0 } }, freezeRows: 2 },
  ] };
  const bytes = await page.evaluate(async content => {
    const { createWorkbookBlob } = await import('/src/lib/workbookIO.ts');
    return Array.from(new Uint8Array(await (await createWorkbookBlob(content)).arrayBuffer()));
  }, workbook);
  const zip = await JSZip.loadAsync(new Uint8Array(bytes), { checkCRC32: true });
  assert.match(await zip.file('xl/worksheets/sheet1.xml').async('string'), /ySplit="1"/);
  assert.match(await zip.file('xl/worksheets/sheet2.xml').async('string'), /ySplit="2"/);
  assert.match(await zip.file('xl/workbook.xml').async('string'), /activeTab="1"/);
  const native = XLSX.read(new Uint8Array(bytes), { type: 'array', cellStyles: true, cellNF: true });
  assert.deepEqual(native.SheetNames, ['Source data', 'Totals']);
  assert.equal(native.Sheets.Totals.B2.f, "'Source data'!B2+'Source data'!B3");
  assert.equal(native.Sheets.Totals.B2.v, 35);
  assert.equal(native.Sheets.Totals.C2.v, 35);
  assert.equal(native.Sheets.Totals.D2.v, 'ok');
  const reimport = await importBytes(bytes, 'Roundtrip.xlsx');
  assert.equal(reimport.warning, undefined);
  assert.equal(reimport.content.sheets.length, 2);
  assert.equal(reimport.content.name, 'Totals');
  assert.equal(reimport.content.sheets[0].freezeRows, 1);
  assert.equal(reimport.content.sheets[1].freezeRows, 2);
  assert.deepEqual(reimport.content.sheets[0].styles.A1, workbook.sheets[0].styles.A1);
  assert.equal(reimport.content.sheets[0].styles.B2.numberFormat, 'currency');
  assert.equal(reimport.content.sheets[0].styles.D4.background, '#FFEEDD');
  assert.ok(Math.abs(reimport.content.sheets[0].columnWidths.A - 190) < 2);
  const csv = await page.evaluate(async content => {
    const { createWorkbookBlob } = await import('/src/lib/workbookIO.ts');
    return (await createWorkbookBlob(content, 'csv')).text();
  }, workbook);
  assert.match(csv, /Active only/);
  assert.match(csv, /35,35,ok/);
  assert.doesNotMatch(csv, /North|South/);
  const xls = XLSX.write(native, { type: 'array', bookType: 'biff8' });
  const legacy = await importBytes(new Uint8Array(xls), 'Legacy.xls');
  assert.equal(legacy.content.sheets.length, 2);
  assert.equal(legacy.content.sheets[0].cells.B3, '21');

  const pptx = await page.evaluate(async () => {
    const { buildPresentationBlob } = await import('/src/editors/slidesPptx.ts');
    const canvas = document.createElement('canvas'); canvas.width = 8; canvas.height = 4; canvas.getContext('2d').fillRect(0, 0, 8, 4);
    const slides = [{ id: 'slide-1', title: 'Editable title', body: 'Editable body', subtitle: 'Subtitle', background: '#244e40', layout: 'title', notes: 'Remember the agenda.\nSecond line.', hidden: true, transition: 'fade', titleSize: 38, bodySize: 20, textColor: '#FFFFFF', fontFamily: 'Arial', elements: [
      { id: 'copy', type: 'text', x: 8, y: 70, width: 32, height: 10, text: 'Overlay copy', color: '#123456', fontSize: 19, bold: true, align: 'right', rotation: 15 },
      { id: 'ellipse', type: 'shape', shape: 'ellipse', x: 70, y: 65, width: 12, height: 20, fill: '#ABCDEF', color: '#345678' },
      { id: 'picture', type: 'image', x: 45, y: 70, width: 20, height: 15, src: canvas.toDataURL('image/png') },
    ] }];
    return Array.from(new Uint8Array(await (await buildPresentationBlob(slides, 'Roundtrip')).arrayBuffer()));
  });
  const deck = await importBytes(pptx, 'Roundtrip.pptx');
  assert.equal(deck.warning, undefined);
  const slide = deck.content[0];
  assert.equal(slide.title, 'Editable title');
  assert.equal(slide.body, 'Editable body');
  assert.equal(slide.subtitle, 'Subtitle');
  assert.equal(slide.hidden, true);
  assert.equal(slide.transition, 'fade');
  assert.equal(slide.notes, 'Remember the agenda.\nSecond line.');
  assert.equal(slide.elements.length, 3);
  assert.equal(slide.elements[0].text, 'Overlay copy');
  assert.equal(slide.elements[0].fontSize, 19);
  assert.equal(slide.elements[0].rotation, 15);
  assert.ok(Math.abs(slide.elements[0].x - 8) < .01);
  assert.equal(slide.elements[1].shape, 'ellipse');
  assert.equal(slide.elements[1].fill.toUpperCase(), '#ABCDEF');
  assert.match(slide.elements[2].src, /^data:image\/png;base64,/);
  const edited = await JSZip.loadAsync(new Uint8Array(pptx));
  const slideSource = await edited.file('ppt/slides/slide1.xml').async('string');
  edited.file('ppt/slides/slide1.xml', slideSource.replace('>Editable title<', '>Changed in PowerPoint<').replace('>Editable body<', '>New native body<'));
  const nativeEdited = await importBytes(await edited.generateAsync({ type: 'uint8array' }), 'Edited.pptx');
  assert.equal(nativeEdited.content[0].title, 'Changed in PowerPoint');
  assert.equal(nativeEdited.content[0].body, 'New native body');

  const download = page.waitForEvent('download');
  await page.evaluate(async () => {
    const { exportDocument } = await import('/src/editors/documentExport.ts');
    await exportDocument('<h1>Review</h1><p>Read <span data-comment-id="comment-1">these words</span>.</p><p><span data-comment-id="comment-2">Resolved quote</span></p>', 'Review', 'docx', { size: 'legal', landscape: true, margin: 'narrow', header: 'Annual report', footer: 'Internal', pageNumbers: true, comments: [{ id: 'comment-1', text: 'Verify totals', quote: 'these words', createdAt: 1791115200000, resolved: false }, { id: 'comment-2', text: 'Checked', quote: 'Resolved quote', createdAt: 1791115200000, resolved: true }] });
  });
  const documentBytes = await readFile(await (await download).path());
  const doc = await importBytes(documentBytes, 'Review.docx');
  assert.equal(doc.pageSetup.landscape, true);
  assert.equal(doc.pageSetup.size, 'legal');
  assert.equal(doc.pageSetup.margin, 'narrow');
  assert.equal(doc.pageSetup.header, 'Annual report');
  assert.match(doc.pageSetup.footer, /^Internal/);
  assert.equal(doc.pageSetup.pageNumbers, true);
  assert.equal(doc.comments.length, 2);
  assert.equal(doc.comments[0].text, 'Verify totals');
  assert.equal(doc.comments[0].quote, 'these words');
  assert.equal(doc.comments[0].resolved, false);
  assert.equal(doc.comments[1].resolved, true);
  assert.match(doc.content, /data-comment-id=/);
  assert.deepEqual(errors, []);
  console.log('PASS: real XLSX all-sheet/style/freeze/width/formula caches and XLS/active CSV import; native PPTX editable objects/images/notes/hidden/transition and PowerPoint edits; DOCX layout/headers/footer/page fields/comments roundtrip.');
} finally { await browser.close(); }
