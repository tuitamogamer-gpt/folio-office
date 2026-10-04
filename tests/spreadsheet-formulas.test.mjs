import test from 'node:test';
import assert from 'node:assert/strict';
import { calculateCells, shiftFormulaReferences } from '../src/editors/spreadsheetFormula.ts';
import { normalizeWorkbook, packWorkbook, renameSheetReferences } from '../src/lib/spreadsheetModel.ts';

const evaluate = (formula, cells = {}) => calculateCells({ ...cells, XFD1048576: formula })('XFD1048576');

test('arithmetic uses spreadsheet precedence, percentages, absolute references and comparisons', () => {
  const cases = {
    '=2+3*4': 14, '=(2+3)*4': 20, '=-2^2': 4, '=2^3^2': 64, '=2^-2': 0.25,
    '=200*10%': 20, '=50%%': 0.005, '=0.1+0.2': 0.3, '=$A$1+B$1+$C1': 9,
    '=A1>=2': 'TRUE', '=A1<>2': 'FALSE', '=A1<3': 'TRUE', '=A1<=1': 'FALSE',
    '="Hello"="hello"': 'TRUE', '="1"=1': 'FALSE', '=Z1=0': 'TRUE', '=Z1=""': 'TRUE',
    '="Total: "&A1+1': 'Total: 3', '="He said ""hello"""': 'He said "hello"',
  };
  for (const [formula, expected] of Object.entries(cases)) assert.equal(evaluate(formula, { A1: '2', B1: '3', C1: '4' }), expected, formula);
});

test('literal identifiers retain leading zeros and long digit strings', () => {
  const f = calculateCells({ A1: '00123', A2: '12345678901234567890', A3: '1.25', A4: "'0009", A5: '1e2', A6: 'TRUE' });
  assert.equal(f('A1'), '00123'); assert.equal(f('A2'), '12345678901234567890');
  assert.equal(f('A3'), 1.25); assert.equal(f('A4'), '0009'); assert.equal(f('A5'), 100); assert.equal(f('A6'), 'TRUE');
  assert.equal(f('Z1'), ''); assert.equal(evaluate('=Z1'), 0);
});

test('aggregate functions ignore referenced text, blanks and booleans, while counting numbers correctly', () => {
  const cells = { A1: '2', A2: '6', A3: 'hello', A4: 'TRUE', A6: '=""', A7: '001' };
  for (const [formula, expected] of [
    ['=SUM(A1:A7)', 8], ['=AVERAGE(A1:A7)', 4], ['=MIN(A1:A7)', 2], ['=MAX(A1:A7)', 6],
    ['=COUNT(A1:A7)', 2], ['=COUNTA(A1:A7)', 6], ['=MEDIAN(A1:A7)', 4], ['=MEDIAN(1,9,3)', 3],
    ['=SUM(A1,A3,A4)', 2], ['=SUM("3",TRUE,A1)', 6], ['=SUM()', 0], ['=COUNT("3",TRUE,"hello")', 2],
    ['=SUM(A2:A1)', 8], ['=AVERAGE(Z1:Z3)', '#DIV/0!'], ['=MIN(Z1:Z3)', 0], ['=SUM("text")', '#VALUE!'],
  ]) assert.equal(evaluate(formula, cells), expected, formula);
  assert.equal(evaluate('=COUNT(A1:A3)', { A1: '1', A2: '=1/0', A3: '#N/A' }), 1);
  assert.equal(evaluate('=COUNTA(A1:A3)', { A1: '1', A2: '=1/0', A3: '#N/A' }), 3);
  assert.equal(evaluate('=SUM(A1:A3)', { A1: '1', A2: '=1/0' }), '#DIV/0!');
});

test('lazy IF and IFERROR never evaluate unused branches, including cyclic and invalid references', () => {
  for (const [formula, expected] of [
    ['=IF(TRUE,42,1/0)', 42], ['=IF(FALSE,1/0,"safe")', 'safe'], ['=IF(FALSE,9)', 'FALSE'],
    ['=IF(TRUE,IF(FALSE,1/0,9),1/0)', 9], ['=IFERROR(1/0,"divide")', 'divide'],
    ['=IFERROR(3,1/0)', 3], ['=IFERROR(1e999,2)', 2], ['=IFERROR(UnknownFunction(),"unsupported")', 'unsupported'],
    ['=IFERROR(A1,"recover")', 'recover'], ['=IFERROR(1/0,1/0)', '#DIV/0!'],
    ['=IF(FALSE,A1,8)', 8], ['=IF(TRUE,,9)', 0], ['=IF(FALSE,1,)', 0],
  ]) assert.equal(evaluate(formula, { A1: '=A1' }), expected, formula);
});

test('cross-sheet references support quoted names, apostrophes, ranges, mixed case and cross-sheet cycles', () => {
  const sheets = [
    { id: 'summary', name: 'Summary', cells: { A1: "='Budget 2026'!$A$1+Sheet2!A1", A2: "=SUM('Budget 2026'!A1:B2)", A3: "='O''Brien'!A1", A4: '=sheet2!A1', A5: '=Missing!A1', A6: '=Sheet2!A2', A7: '=SUM(Sheet2!A1:Summary!A1)' } },
    { id: 'budget', name: 'Budget 2026', cells: { A1: '10', A2: '20', B1: '3', B2: '7' } },
    { id: 'other', name: 'Sheet2', cells: { A1: '2', A2: '=Summary!A6' } },
    { id: 'quote', name: "O'Brien", cells: { A1: '19' } },
  ];
  const f = calculateCells(sheets[0].cells, { sheets, currentSheetId: 'summary' });
  for (const [key, expected] of [['A1', 12], ['A2', 40], ['A3', 19], ['A4', 2], ['A5', '#REF!'], ['A6', '#CYCLE!'], ['A7', '#REF!']]) assert.equal(f(key), expected, key);
  const updated = calculateCells({ ...sheets[0].cells, B1: '5', B2: '=B1+1' }, { sheets, currentSheetId: 'summary' });
  assert.equal(updated('B2'), 6, 'the live current sheet takes priority over an older workbook copy');
});

test('conditional sums and counts support comparisons, wildcards, escaped wildcards and empty cells', () => {
  const cells = { A1: 'Apple', A2: 'apricot', A3: 'Berry', A4: 'a*', B1: '10', B2: '25', B3: '5', B4: '7' };
  for (const [formula, expected] of [
    ['=COUNTIF(A1:A4,"a*")', 3], ['=COUNTIF(A1:A4,"a~*")', 1], ['=COUNTIF(A1:A4,"?????")', 2],
    ['=COUNTIF(A1:A4,"<>Berry")', 3], ['=COUNTIF(B1:B4,">=10")', 2], ['=COUNTIF(A1:A5,"")', 1],
    ['=SUMIF(A1:A4,"a*",B1:B4)', 42], ['=SUMIF(B1:B4,">10")', 25], ['=SUMIF(B1:B4,"<="&10)', 22],
    ['=SUMIF(A1:A4,"Berry",B1:B4)', 5],
  ]) assert.equal(evaluate(formula, cells), expected, formula);
});

test('logical, numeric and text functions enforce argument bounds and spreadsheet rounding', () => {
  for (const [formula, expected] of [
    ['=AND(TRUE,1,2>1)', 'TRUE'], ['=OR(FALSE,0,2<1)', 'FALSE'], ['=NOT(FALSE)', 'TRUE'],
    ['=AND(A1:A3)', 'TRUE'], ['=ROUND(-1.5,0)', -2], ['=ROUND(1.005,2)', 1.01], ['=ROUND(149,-1)', 150],
    ['=ROUNDUP(-1.21,1)', -1.3], ['=ROUNDDOWN(-1.29,1)', -1.2], ['=ROUNDDOWN(1.15,2)', 1.15], ['=ROUNDUP(1.11,2)', 1.11], ['=ABS(-7)', 7], ['=SQRT(9)', 3],
    ['=SQRT(-1)', '#NUM!'], ['=POWER(2,4)', 16], ['=MOD(-3,2)', 1], ['=MOD(3,-2)', -1],
    ['=LEN("Folio")', 5], ['=LEFT("Folio",2)', 'Fo'], ['=RIGHT("Folio",2)', 'io'], ['=RIGHT("Folio",0)', ''],
    ['=MID("Folio",2,3)', 'oli'], ['=MID("Folio",0,1)', '#VALUE!'], ['=LEFT("Folio",-1)', '#VALUE!'],
    ['=TRIM("  a   b  ")', 'a b'], ['=UPPER("hello")', 'HELLO'], ['=LOWER("HELLO")', 'hello'],
    ['=CONCAT("Value ",A1,TRUE)', 'Value 1TRUE'], ['=CONCAT(A1:A3)', '1textTRUE'], ['=CONCATENATE("a","b")', 'ab'],
    ['=ROUND(1)', '#VALUE!'], ['=NOT()', '#VALUE!'], ['=SUMIF(A1:A3)', '#VALUE!'],
  ]) assert.equal(evaluate(formula, { A1: '1', A2: 'text', A3: 'TRUE' }), expected, formula);
});

test('DATE and TODAY return Excel serial dates including the 1900 leap-year convention', () => {
  assert.equal(evaluate('=DATE(1900,1,1)'), 1);
  assert.equal(evaluate('=DATE(1900,2,28)'), 59);
  assert.equal(evaluate('=DATE(1900,2,29)'), 60);
  assert.equal(evaluate('=DATE(1900,3,1)'), 61);
  assert.equal(evaluate('=DATE(1900,3,0)'), 60);
  assert.equal(evaluate('=DATE(1900,2,30)'), 61);
  assert.equal(evaluate('=DATE(2024,1,1)'), 45292);
  assert.equal(evaluate('=DATE(2024,13,1)'), 45658);
  assert.equal(evaluate('=DATE(24,1,1)'), 8767);
  assert.equal(evaluate('=DATE(-1,1,1)'), '#NUM!');
  const now = new Date();
  assert.equal(evaluate('=TODAY()'), evaluate(`=DATE(${now.getFullYear()},${now.getMonth() + 1},${now.getDate()})`));
});

test('malformed, excessive and executable-looking formulas are safely rejected', () => {
  for (const formula of ['=globalThis.process.exit()', '=1;process.exit()', '=SUM(A1:A2', '="unclosed', '=1 2', '=A0', '=XFE1', '=A1048577']) {
    assert.match(String(evaluate(formula)), /^#/, formula);
  }
  assert.equal(evaluate('=1/0'), '#DIV/0!');
  assert.equal(evaluate('=SUM(A1:XFD1048576)'), '#NUM!');
  assert.equal(evaluate('=10^999'), '#NUM!');
  assert.equal(evaluate('='.concat('('.repeat(300), '1', ')'.repeat(300))), '#DEPTH!');
  assert.equal(evaluate('=A1', { A1: '=A2', A2: '=A1' }), '#CYCLE!');
});

test('fill shifts mixed absolute references across ranges and sheets without rewriting text or names', () => {
  assert.equal(shiftFormulaReferences('=A1+$B1+C$1+$D$1', 2, 3), '=D3+$B3+F$1+$D$1');
  assert.equal(shiftFormulaReferences('=SUM(A1:B3)+Sheet2!A1', 1, 1), '=SUM(B2:C4)+Sheet2!B2');
  assert.equal(shiftFormulaReferences('="A1"&\'Budget A1\'!A1+\'O\'\'Brien\'!B$2', 3, 2), '="A1"&\'Budget A1\'!C4+\'O\'\'Brien\'!D$2');
  assert.equal(shiftFormulaReferences('=LOG10(A1)+A1_LABEL+S1 ! B2', 1, 1), '=LOG10(B2)+A1_LABEL+S1 ! C3');
  assert.equal(shiftFormulaReferences('=A1+$B$2', -1, 0), '=#REF!+$B$2');
  assert.equal(shiftFormulaReferences('=XFD1048576', 1, 1), '=#REF!');
  assert.equal(shiftFormulaReferences('A1', 2, 2), 'A1');
});

test('workbook migration keeps every tab, style, column width and active tab with a legacy mirror', () => {
  const legacy = { name: 'Budget', cells: { A1: '001' } };
  const normalized = normalizeWorkbook(legacy);
  assert.equal(normalized.sheets.length, 1); assert.equal(normalized.sheets[0].name, 'Budget'); assert.equal(normalized.sheets[0].cells.A1, '001');
  const sheets = [{ id: 'one', name: 'One', cells: { A1: '1' }, styles: { A1: { bold: true, numberFormat: 'currency' } }, columnWidths: { A: 150 }, freezeRows: 1 }, { id: 'two', name: 'Two', cells: { B2: '=One!A1' } }];
  const packed = packWorkbook(sheets, 'two');
  assert.equal(packed.name, 'Two'); assert.deepEqual(packed.cells, sheets[1].cells); assert.deepEqual(packed.sheets, sheets);
  const read = normalizeWorkbook(packed); assert.deepEqual(read, { sheets, activeSheetId: 'two' });
  read.sheets[0].styles.A1.bold = false; assert.equal(sheets[0].styles.A1.bold, true, 'normalizing does not mutate undo/history snapshots');
  const dirty = normalizeWorkbook({ sheets: [{ id: 'same', name: 'A/B', cells: {} }, { id: 'same', name: 'ab', cells: {} }, { name: 'x'.repeat(40), cells: {} }], activeSheetId: 'gone' });
  assert.deepEqual(dirty.sheets.map(sheet => sheet.name), ['AB', 'ab (2)', 'x'.repeat(31)]);
  assert.equal(new Set(dirty.sheets.map(sheet => sheet.id)).size, 3); assert.equal(dirty.activeSheetId, dirty.sheets[0].id);
  assert.equal(normalizeWorkbook(null).sheets.length, 1); assert.equal(packWorkbook([], 'missing').name, 'Sheet 1');
});

test('sheet rename and invalid-name migration rewrite references without rewriting quoted text or chaining names', () => {
  assert.equal(renameSheetReferences('=Sheet1!A1+"Sheet1!A1"+Foo.Sheet1!A2', 'sheet1', "Owner's budget"), '=\'Owner\'\'s budget\'!A1+"Sheet1!A1"+Foo.Sheet1!A2');
  assert.equal(renameSheetReferences("='O''Brien' ! $A$1", "O'Brien", 'Plan'), "='Plan'! $A$1");
  const workbook = normalizeWorkbook({ sheets: [
    { id: 'a', name: 'A/B', cells: { A1: '2' } },
    { id: 'b', name: 'AB', cells: { A1: '5', A2: "='A/B'!A1+'AB'!A1", A3: '=\"A/B!A1\"' } },
  ], activeSheetId: 'b' });
  assert.deepEqual(workbook.sheets.map(sheet => sheet.name), ['AB', 'AB (2)']);
  assert.equal(workbook.sheets[1].cells.A2, "='AB'!A1+'AB (2)'!A1");
  assert.equal(workbook.sheets[1].cells.A3, '=\"A/B!A1\"');
  assert.equal(calculateCells(workbook.sheets[1].cells, { sheets: workbook.sheets, currentSheetId: 'b' })('A2'), 7);
});
