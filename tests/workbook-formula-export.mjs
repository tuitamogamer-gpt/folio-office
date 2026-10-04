import assert from 'node:assert/strict';
import { chromium } from '@playwright/test';
import JSZip from 'jszip';
import XLSX from 'xlsx-js-style';

const browser = await chromium.launch({ headless: true, executablePath: '/usr/bin/chromium', args: ['--no-sandbox'] });
const page = await browser.newPage();
const errors = [];
page.on('pageerror', error => errors.push(error.message));
await page.goto(process.env.BASE_URL || 'http://localhost:5173');

try {
  const result = await page.evaluate(async () => {
    const { createWorkbookBlob, importWorkbook } = await import('/src/lib/workbookIO.ts');
    const { calculateCells, FORMULA_CATALOG } = await import('/src/editors/spreadsheetFormula.ts');
    const { packWorkbook } = await import('/src/lib/spreadsheetModel.ts');
    const formulas = {
      A1: '=XLOOKUP("North",Data!A2:A4,Data!C2:C4)',
      A2: '=VLOOKUP("South",Data!A2:C4,3,FALSE)',
      A3: '=HLOOKUP("West",Data!E1:G2,2,FALSE)',
      A4: '=INDEX(Data!C2:C4,MATCH("South",Data!A2:A4,0))',
      A5: '=SUMIFS(Data!C2:C4,Data!A2:A4,"North",Data!B2:B4,"Paid")',
      A6: '=COUNTIFS(Data!A2:A4,"North",Data!B2:B4,"Paid")',
      A7: '=SUM(XLOOKUP("North",Data!A2:A4,Data!C2:C4),1)',
      A8: '=_xlfn.XLOOKUP("South",Data!A2:A4,Data!C2:C4)',
      A9: '=xlookup ("North",Data!A2:A4,Data!C2:C4)',
      A10: '=XLOOKUP("Unknown",Data!A2:A4,Data!C2:C4,"Quoted ""XLOOKUP("" and _xlfn.XLOOKUP(")',
      A11: '=CONCAT("XLOOKUP("," | _xlfn.XLOOKUP("," | ",XLOOKUP("North",Data!A2:A4,Data!C2:C4))',
      A12: "='XLOOKUP(''26)'!A1+XLOOKUP(\"North\",Data!A2:A4,Data!C2:C4)",
      A13: '=_xlfn._xlws.XLOOKUP("North",Data!A2:A4,Data!C2:C4)',
      A14: '=IFERROR(XLOOKUP("Unknown",Data!A2:A4,Data!C2:C4),"missing")',
      A15: '=XLOOKUP("Unknown",Data!A2:A4,Data!C2:C4)',
      A16: '=CONCATENATE("literal CONCAT( and XLOOKUP("," preserved")',
      A17: '=_xlfn.CONCAT("Qualified ","CONCAT(")',
      A18: '=XLOOKUP("Unknown",Data!A2:A4,Data!C2:C4,"")',
    };
    const catalogCells = { A2: '20', A3: '40', A4: '60', B2: '100', B3: '200', B4: '300', C2: 'Paid', C3: 'Open', C4: 'Paid', F1: '20', G1: '40', H1: '60', F2: '20', G2: 'Tier one', H2: '4', F3: '40', G3: 'Tier two', H3: '8' };
    FORMULA_CATALOG.forEach((entry, index) => { catalogCells[`Z${index + 1}`] = entry.example; });
    const sheets = [
      { id: 'data', name: 'Data', cells: { A2: 'North', B2: 'Paid', C2: '20', A3: 'South', B3: 'Open', C3: '30', A4: 'North', B4: 'Paid', C4: '40', E1: 'North', F1: 'South', G1: 'West', E2: '2', F2: '3', G2: '4' } },
      { id: 'report', name: 'Report', cells: formulas },
      { id: 'quoted', name: "XLOOKUP('26)", cells: { A1: '7' } },
      { id: 'catalog', name: 'Catalog', cells: catalogCells },
    ];
    const content = packWorkbook(sheets, 'report');
    const expected = {};
    for (const sheet of sheets) {
      const evaluate = calculateCells(sheet.cells, { sheets, currentSheetId: sheet.id });
      expected[sheet.name] = Object.fromEntries(Object.entries(sheet.cells).filter(([, value]) => value.startsWith('=')).map(([address]) => [address, evaluate(address)]));
    }
    const snapshot = JSON.stringify(content);
    const blob = await createWorkbookBlob(content);
    const imported = await importWorkbook(new File([blob], 'Functions.xlsx'));
    const recalculated = {};
    for (const sheet of imported.content.sheets) {
      const evaluate = calculateCells(sheet.cells, { sheets: imported.content.sheets, currentSheetId: sheet.id });
      recalculated[sheet.name] = Object.fromEntries(Object.entries(sheet.cells).filter(([, value]) => value.startsWith('=')).map(([address]) => [address, evaluate(address)]));
    }
    const secondBlob = await createWorkbookBlob(imported.content);
    return { bytes: Array.from(new Uint8Array(await blob.arrayBuffer())), secondBytes: Array.from(new Uint8Array(await secondBlob.arrayBuffer())), formulas, expected, recalculated, imported: imported.content.sheets.find(sheet => sheet.name === 'Report').cells, unchanged: snapshot === JSON.stringify(content), names: FORMULA_CATALOG.map(entry => entry.name) };
  });
  assert.ok(result.unchanged, 'export never adds prefixes to UI formulas');
  assert.deepEqual(result.recalculated, result.expected, 'every catalog function recalculates identically after import');
  const zip = await JSZip.loadAsync(new Uint8Array(result.bytes), { checkCRC32: true });
  const xml = await zip.file('xl/worksheets/sheet2.xml').async('string');
  assert.match(xml, /<f>_xlfn\.XLOOKUP\(/);
  assert.match(xml, /<f>_xlfn\.CONCAT\(/);
  assert.doesNotMatch(xml, /_xlfn\._xlfn\./);
  const workbook = XLSX.read(new Uint8Array(result.bytes), { type: 'array', xlfn: true });
  const report = workbook.Sheets.Report;
  assert.equal(report.A1.f, result.formulas.A1.slice(1).replace(/^XLOOKUP/, '_xlfn.XLOOKUP'));
  for (const address of ['A2', 'A3', 'A4', 'A5', 'A6', 'A8', 'A13', 'A16', 'A17']) assert.equal(report[address].f, result.formulas[address].slice(1), `${address}: older and qualified functions unchanged`);
  assert.equal(report.A9.f, '_xlfn.xlookup ("North",Data!A2:A4,Data!C2:C4)');
  assert.equal(report.A10.f, '_xlfn.XLOOKUP("Unknown",Data!A2:A4,Data!C2:C4,"Quoted ""XLOOKUP("" and _xlfn.XLOOKUP(")');
  assert.equal(report.A11.f, '_xlfn.CONCAT("XLOOKUP("," | _xlfn.XLOOKUP("," | ",_xlfn.XLOOKUP("North",Data!A2:A4,Data!C2:C4))');
  assert.equal(report.A12.f, "'XLOOKUP(''26)'!A1+_xlfn.XLOOKUP(\"North\",Data!A2:A4,Data!C2:C4)");
  assert.equal(report.A10.v, 'Quoted "XLOOKUP(" and _xlfn.XLOOKUP(');
  assert.equal(report.A11.v, 'XLOOKUP( | _xlfn.XLOOKUP( | 20');
  assert.equal(report.A12.v, 27);
  assert.equal(report.A15.t, 'e');
  assert.equal(report.A15.v, 42, 'missing XLOOKUP cached as the native #N/A error');
  assert.equal(report.A18.v, '', 'empty fallback is retained');
  for (const [sheet, cells] of Object.entries(result.expected)) for (const [address, expected] of Object.entries(cells)) {
    const native = workbook.Sheets[sheet][address];
    const value = native.t === 'b' ? native.v ? 'TRUE' : 'FALSE' : native.t === 'e' ? native.w : native.v;
    assert.equal(value, expected, `${sheet}!${address}: cached result agrees with the calculator`);
  }
  for (const address of Object.keys(result.formulas)) assert.equal(result.imported[address], `=${report[address].f}`, `${address}: import preserves prefixes and quoted content`);
  const second = XLSX.read(new Uint8Array(result.secondBytes), { type: 'array', xlfn: true });
  for (const address of Object.keys(result.formulas)) assert.equal(second.Sheets.Report[address].f, report[address].f, `${address}: repeated export is idempotent`);
  for (const name of ['VLOOKUP', 'HLOOKUP', 'XLOOKUP', 'INDEX', 'MATCH', 'SUMIFS', 'COUNTIFS', 'CONCAT']) assert.ok(result.names.includes(name));
  assert.deepEqual(errors, []);
  console.log(`PASS: ${result.names.length} catalog functions retain cached values and reload results; XLOOKUP/CONCAT have native _xlfn prefixes; legacy/qualified functions, quoted strings and sheet names stay intact; repeated exports do not mutate formulas.`);
} finally { await browser.close(); }
