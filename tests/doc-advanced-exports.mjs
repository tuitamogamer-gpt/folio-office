import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import JSZip from 'jszip';

const browser = await chromium.launch({ headless: true, executablePath: '/usr/bin/chromium', args: ['--no-sandbox'] });
const page = await browser.newPage();
const runtimeErrors = [];
page.on('pageerror', error => runtimeErrors.push(error.message));
await page.goto(process.env.BASE_URL || 'http://localhost:5174');

async function exported(html, format, options = {}, name = 'Advanced document') {
  const pending = page.waitForEvent('download');
  await page.evaluate(async ({ html, format, options, name }) => {
    const { exportDocument } = await import('/src/editors/documentExport.ts');
    await exportDocument(html, name, format, options);
  }, { html, format, options, name });
  const download = await pending;
  assert.equal(download.suggestedFilename(), `${name.replace(/[\\/:*?"<>|]/g, '-')}.${format}`);
  return readFile(await download.path());
}

async function xml(zip, name) {
  const file = zip.file(name);
  assert.ok(file, `${name} exists`);
  const content = await file.async('string');
  assert.equal(await page.evaluate(source => new DOMParser().parseFromString(source, 'application/xml').querySelector('parsererror')?.textContent || '', content), '');
  return content;
}

try {
  const image = await page.evaluate(() => {
    const canvas = document.createElement('canvas');
    canvas.width = 2; canvas.height = 1;
    canvas.getContext('2d').fillRect(0, 0, 2, 1);
    return canvas.toDataURL('image/png');
  });
  const markup = `<h1>Review draft</h1>
    <p data-indent="2" style="margin-left:72pt;text-align:justify;line-height:1.5"><span data-comment-id="review"><strong>Selected</strong> words</span> H<sub>2</sub>O x<sup>2</sup></p>
    <p><span data-comment-id="review">across a paragraph</span> <a href="https://example.com">source link</a></p>
    <div data-page-break="true" style="break-after:page"></div><p>Next page</p>
    <img src="${image}" width="240" data-align="right" alt="Ratio test">
    <table><tbody><tr><td colspan="2" rowspan="2" style="background-color:rgb(219,234,254)"><p>Merged blue cell</p></td><th><p>Header</p></th></tr><tr><td><p>Third column</p></td></tr></tbody></table>
    <ol start="3"><li><p>Numbered item</p><ul><li><p><span data-comment-id="nested">Nested bullet</span></p></li></ul></li></ol>`;
  const comments = [
    { id: 'review', text: 'Check these numbers.\nThen approve.', quote: 'Selected words across a paragraph', createdAt: Date.UTC(2026, 9, 4, 12), resolved: true },
    { id: 'deleted', text: 'Keep this note.', quote: 'Text removed later', createdAt: Date.UTC(2026, 9, 4, 13), resolved: false },
    { id: 'nested', text: 'Nested list note.', quote: 'Nested bullet', createdAt: Date.UTC(2026, 9, 4, 14), resolved: false },
  ];
  const bytes = await exported(markup, 'docx', { size: 'legal', landscape: true, margin: 'narrow', header: 'Folio review', footer: 'Internal draft', pageNumbers: true, comments });
  const zip = await JSZip.loadAsync(bytes, { checkCRC32: true });
  const document = await xml(zip, 'word/document.xml');
  const commentXML = await xml(zip, 'word/comments.xml');
  const commentsExtended = await xml(zip, 'word/commentsExtended.xml');
  const relationships = await xml(zip, 'word/_rels/document.xml.rels');
  const header = await xml(zip, Object.keys(zip.files).find(path => /^word\/header\d+\.xml$/.test(path)));
  const footer = await xml(zip, Object.keys(zip.files).find(path => /^word\/footer\d+\.xml$/.test(path)));
  const numbering = await xml(zip, 'word/numbering.xml');
  assert.match(document, /<w:pgSz[^>]*w:w="20160"[^>]*w:h="12240"[^>]*w:orient="landscape"/);
  assert.match(document, /<w:pgMar[^>]*w:top="720"[^>]*w:right="720"[^>]*w:bottom="720"[^>]*w:left="720"/);
  assert.match(document, /<w:ind[^>]*w:left="1440"/);
  assert.match(document, /<w:spacing[^>]*w:line="360"/);
  assert.match(document, /<w:vertAlign w:val="subscript"/);
  assert.match(document, /<w:vertAlign w:val="superscript"/);
  assert.match(document, /<w:br w:type="page"/);
  assert.match(document, /<wp:extent cx="2286000" cy="1143000"/);
  assert.match(document, /<w:jc w:val="right"/);
  assert.equal((document.match(/<w:gridSpan w:val="2"/g) || []).length, 2, 'horizontal span preserved on both rows');
  assert.match(document, /<w:vMerge w:val="restart"/);
  assert.match(document, /<w:vMerge w:val="continue"/);
  assert.equal((document.match(/w:fill="DBEAFE"/g) || []).length, 2, 'merged shading covers both rows');
  assert.equal((document.match(/<w:commentRangeStart w:id="0"/g) || []).length, 1, 'one range across multiple comment spans');
  assert.equal((document.match(/<w:commentRangeEnd w:id="0"/g) || []).length, 1);
  assert.match(document, /<w:r><w:commentReference w:id="0"\/><\/w:r>/, 'comment reference is a run child');
  assert.match(document, /<w:commentRangeStart w:id="1"/);
  assert.match(document, /<w:commentRangeStart w:id="2"\/><w:r><w:t[^>]*>Nested bullet<\/w:t><\/w:r><w:commentRangeEnd w:id="2"/);
  assert.match(commentXML, /Check these numbers\./);
  assert.match(commentXML, /Then approve\./);
  assert.match(commentXML, /Original selection: Text removed later/);
  assert.match(commentXML, /w:date="2026-10-04T12:00:00(?:\.000)?Z"/);
  assert.match(commentsExtended, /w15:done="1"/);
  assert.match(commentsExtended, /w15:done="0"/);
  assert.match(commentsExtended, /w15:paraId="00000001"/);
  assert.match(commentXML, /w14:paraId="00000001"/);
  assert.match(relationships, /commentsExtended[^>]*Target="commentsExtended.xml"/);
  assert.match(await xml(zip, '[Content_Types].xml'), /PartName="\/word\/commentsExtended.xml"/);
  assert.match(relationships, /Target="https:\/\/example.com"[^>]*TargetMode="External"/);
  assert.match(header, /Folio review/);
  assert.match(footer, /Internal draft/);
  assert.match(footer, /<w:instrText[^>]*>PAGE<\/w:instrText>/);
  assert.match(numbering, /<w:start w:val="3"/);
  assert.match(numbering, /w:val="bullet"/);
  assert.ok(Object.keys(zip.files).some(path => /^word\/media\/[^/]+\.png$/.test(path)));

  for (const [size, dimensions] of Object.entries({ a4: [11906, 16838], letter: [12240, 15840], legal: [12240, 20160] })) {
    for (const landscape of [false, true]) {
      const simple = await JSZip.loadAsync(await exported('<p>Paper size check</p>', 'docx', { size, landscape }));
      const output = await xml(simple, 'word/document.xml');
      const [width, height] = landscape ? [...dimensions].reverse() : dimensions;
      assert.match(output, new RegExp(`<w:pgSz[^>]*w:w="${width}"[^>]*w:h="${height}"`));
      assert.match(output, /<w:pgMar[^>]*w:top="1440"/);
      assert.equal(Object.keys(simple.files).some(path => /^word\/(header|footer)\d+\.xml$/.test(path)), false, 'no unsolicited header/footer');
    }
  }

  const html = (await exported(markup + '<script>throw new Error("unsafe")</script>', 'html', { size: 'letter', landscape: true, margin: 'wide', header: '</style><script>unsafe()</script>', footer: 'Review & approve', pageNumbers: true, comments })).toString();
  assert.match(html, /@page\{size:letter landscape;margin:1.5in/);
  assert.match(html, /counter\(page\)/);
  assert.match(html, /data-page-break="true"/);
  assert.match(html, /Check these numbers\./);
  assert.match(html, /\(Resolved\)/);
  assert.doesNotMatch(html, /<script>/);
  assert.match(html, /Review &amp; approve/);

  const text = (await exported('<h1>Title</h1><p>H<sub>2</sub>O</p><ol start="3"><li>Third</li></ol>', 'txt')).toString();
  assert.match(text, /Title\nH2O\n3\. Third/);
  assert.deepEqual(runtimeErrors, []);
  console.log('PASS: DOCX OOXML for page sizes/orientation/margins, header/footer/page field, native resolved comments, breaks, indentation, sub/sup, merged shaded cells, image aspect/alignment, hyperlinks and nested numbering; safe HTML layout/comments and TXT.');
} finally {
  await browser.close();
}
