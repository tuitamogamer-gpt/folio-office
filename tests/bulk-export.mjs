import assert from 'node:assert/strict';
import { chromium } from '@playwright/test';
import JSZip from 'jszip';
import XLSX from 'xlsx-js-style';

const browser = await chromium.launch({ headless: true, executablePath: '/usr/bin/chromium', args: ['--no-sandbox'] });
const page = await browser.newPage();
const errors = [];
const downloads = [];
page.on('pageerror', error => errors.push(error.message));
page.on('download', download => downloads.push(download.suggestedFilename()));
await page.goto(process.env.BASE_URL || 'http://localhost:5173');

try {
  const result = await page.evaluate(async () => {
    const { exportOfficeFilesZip } = await import('/src/lib/bulkExport.ts');
    const { buildDocumentBlob } = await import('/src/editors/documentExport.ts');
    const base = { createdAt: 1791115200000, updatedAt: 1791115200000, starred: false, trashed: false };
    const document = (name, index) => ({ ...base, id: `doc-${index}`, name, kind: 'document', content: `<p>Document ${index}</p>` });
    const docs = ['Report', 'report.docx', 'REPORT (2)', '../../Outside', 'Sales/a', 'Sales\\a', 'CON', 'Résumé', 'Re\u0301sume\u0301', 'é'.repeat(200)].map(document);
    docs[0].content = '<h1>Annual report</h1><p>Check <span data-comment-id="c1">these numbers</span>.</p>';
    docs[0].pageSetup = { landscape: true, margin: 'narrow', size: 'legal', header: 'Annual review', footer: 'Internal', pageNumbers: true };
    docs[0].comments = [{ id: 'c1', text: 'Verified total', quote: 'these numbers', createdAt: 1791115200000, resolved: true }];
    const workbook = { ...base, id: 'budget', name: 'Budget.xlsx', kind: 'spreadsheet', content: { name: 'Totals', cells: { A1: "='Source'!A1*2" }, activeSheetId: 'totals', sheets: [
      { id: 'source', name: 'Source', cells: { A1: '21' }, styles: { A1: { bold: true, background: '#AABBCC', numberFormat: 'currency' } }, freezeRows: 1 },
      { id: 'totals', name: 'Totals', cells: { A1: "='Source'!A1*2" } },
    ] } };
    const canvas = window.document.createElement('canvas'); canvas.width = 2; canvas.height = 1; canvas.getContext('2d').fillRect(0, 0, 2, 1);
    const deck = { ...base, id: 'deck', name: 'Deck.pptx', kind: 'presentation', content: [{ id: 's1', title: 'Native slide', body: 'Native content', background: '#FFFFFF', layout: 'content', notes: 'Present the chart next.', hidden: true, transition: 'fade', elements: [
      { id: 'shape', type: 'shape', shape: 'ellipse', x: 10, y: 20, width: 20, height: 15, fill: '#ABCDEF' },
      { id: 'picture', type: 'image', x: 65, y: 65, width: 20, height: 10, src: canvas.toDataURL('image/png') },
    ] }] };
    const before = JSON.stringify([...docs, workbook, deck]);
    const blob = await exportOfficeFilesZip([...docs, workbook, deck]);
    const plain = await buildDocumentBlob('<p>Plain <strong>text</strong></p>', 'Pure', 'txt');
    const html = await buildDocumentBlob('<p>HTML <script>alert(1)</script>export</p>', 'Pure', 'html');
    const failures = [];
    for (const files of [[], [docs[0], { ...base, id: 'broken', name: 'Broken deck', kind: 'presentation', content: [] }], [{ ...docs[0], name: 'Broken document', content: null }], [{ ...workbook, name: 'Broken workbook', content: { sheets: [{ name: 'Missing cells' }] } }]]) {
      try { await exportOfficeFilesZip(files); failures.push('unexpected success'); } catch (error) { failures.push(error.message); }
    }
    return { bytes: Array.from(new Uint8Array(await blob.arrayBuffer())), type: blob.type, before, after: JSON.stringify([...docs, workbook, deck]), plain: await plain.text(), html: await html.text(), failures };
  });
  assert.equal(result.type, 'application/zip');
  assert.equal(result.before, result.after, 'native exports do not mutate source files');
  assert.equal(result.plain, 'Plain text');
  assert.match(result.html, /HTML export/);
  assert.doesNotMatch(result.html, /<script>/);
  assert.match(result.failures[0], /Select at least one file/);
  for (const [index, name] of ['Broken deck', 'Broken document', 'Broken workbook'].entries()) {
    assert.ok(result.failures[index + 1].includes(name));
    assert.match(result.failures[index + 1], /No ZIP archive was created/);
  }
  const archive = await JSZip.loadAsync(new Uint8Array(result.bytes), { checkCRC32: true });
  const entries = Object.keys(archive.files);
  assert.equal(entries.length, 12, 'every requested document, workbook, and presentation exists');
  assert.equal(new Set(entries.map(name => name.toLowerCase())).size, entries.length, 'case-insensitive names are unique');
  assert.ok(entries.includes('Report.docx'));
  assert.ok(entries.includes('report (2).docx'));
  assert.ok(entries.includes('REPORT (2) (2).docx'));
  assert.ok(entries.includes('Sales-a.docx'));
  assert.ok(entries.includes('Sales-a (2).docx'));
  assert.ok(entries.includes('_CON.docx'));
  assert.ok(entries.includes('Résumé.docx'));
  assert.ok(entries.includes('Résumé (2).docx'));
  for (const name of entries) {
    assert.equal(archive.files[name].dir, false);
    assert.doesNotMatch(name, /[\\/\x00-\x1f]/);
    assert.ok(!name.startsWith('.'));
    assert.ok(new TextEncoder().encode(name).length < 240);
  }
  const doc = await JSZip.loadAsync(await archive.file('Report.docx').async('uint8array'), { checkCRC32: true });
  assert.match(await doc.file('word/document.xml').async('string'), /Annual report/);
  assert.match(await doc.file('word/document.xml').async('string'), /w:orient="landscape"/);
  assert.match(await doc.file('word/comments.xml').async('string'), /Verified total/);
  assert.match(await doc.file('word/commentsExtended.xml').async('string'), /w15:done="1"/);
  const headerPath = Object.keys(doc.files).find(path => /^word\/header\d+\.xml$/.test(path));
  const footerPath = Object.keys(doc.files).find(path => /^word\/footer\d+\.xml$/.test(path));
  assert.match(await doc.file(headerPath).async('string'), /Annual review/);
  assert.match(await doc.file(footerPath).async('string'), /PAGE/);
  const duplicate = await JSZip.loadAsync(await archive.file('report (2).docx').async('uint8array'));
  assert.match(await duplicate.file('word/document.xml').async('string'), /Document 1/);
  const workbookBytes = await archive.file('Budget.xlsx').async('uint8array');
  const workbook = XLSX.read(workbookBytes, { type: 'array', cellStyles: true });
  assert.deepEqual(workbook.SheetNames, ['Source', 'Totals']);
  assert.equal(workbook.Sheets.Totals.A1.f, "'Source'!A1*2");
  assert.equal(workbook.Sheets.Totals.A1.v, 42);
  const workbookZip = await JSZip.loadAsync(workbookBytes);
  assert.match(await workbookZip.file('xl/worksheets/sheet1.xml').async('string'), /state="frozen"/);
  assert.match(await workbookZip.file('xl/styles.xml').async('string'), /AABBCC/);
  const pptx = await JSZip.loadAsync(await archive.file('Deck.pptx').async('uint8array'), { checkCRC32: true });
  const slide = await pptx.file('ppt/slides/slide1.xml').async('string');
  assert.match(slide, /Native slide/);
  assert.match(slide, /prst="ellipse"/);
  assert.match(slide, /<p:pic>/);
  assert.match(slide, /<p:fade\/>/);
  assert.match(slide, /show="0"/);
  assert.match(await pptx.file('ppt/notesSlides/notesSlide1.xml').async('string'), /Present the chart next\./);
  assert.deepEqual(downloads, [], 'pure blob APIs never trigger downloads');
  assert.deepEqual(errors, []);
  console.log('PASS: native DOCX/XLSX/PPTX ZIP exports preserve content, comments/layout/formulas/elements; safe unique filenames; all-or-nothing errors; pure blob APIs produce no downloads or mutations.');
} finally { await browser.close(); }
