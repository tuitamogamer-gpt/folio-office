import test from 'node:test';
import assert from 'node:assert/strict';
import { calculateCells, FORMULA_CATALOG, shiftFormulaReferences } from '../src/editors/spreadsheetFormula.ts';
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

test('VLOOKUP and HLOOKUP return exact, wildcard and approximate matches with typed keys', () => {
  const cells = {
    A1: '10', B1: 'Starter', C1: '4', A2: '20', B2: 'Team', C2: '8', A3: '20', B3: 'Team plus', C3: '12', A4: '30', B4: 'Business', C4: '16',
    E1: 'Code', F1: 'apple', G1: 'apricot', H1: 'a*', E2: 'Price', F2: '2', G2: '3', H2: '4',
    J1: "'10", K1: 'Text identifier', J2: '10', K2: 'Numeric identifier',
  };
  for (const [formula, expected] of [
    ['=VLOOKUP(20,A1:C4,2,FALSE)', 'Team'], ['=VLOOKUP(20,A1:C4,2,TRUE)', 'Team plus'],
    ['=VLOOKUP(25,A1:C4,3)', 12], ['=VLOOKUP(5,A1:C4,2)', '#N/A'], ['=VLOOKUP(31,A1:C4,2)', 'Business'],
    ['=VLOOKUP(25,A1:C4,2,FALSE)', '#N/A'], ['=VLOOKUP(10,A1:C4,0,FALSE)', '#VALUE!'], ['=VLOOKUP(10,A1:C4,4,FALSE)', '#REF!'],
    ['=VLOOKUP(10,J1:K2,2,FALSE)', 'Numeric identifier'], ['=VLOOKUP("10",J1:K2,2,FALSE)', 'Text identifier'],
    ['=HLOOKUP("APPLE",F1:H2,2,FALSE)', 2], ['=HLOOKUP("ap*",F1:H2,2,FALSE)', 2], ['=HLOOKUP("a~*",F1:H2,2,FALSE)', 4],
    ['=HLOOKUP("missing",F1:H2,2,FALSE)', '#N/A'], ['=HLOOKUP("apple",F1:H2,3,FALSE)', '#REF!'],
  ]) assert.equal(evaluate(formula, cells), expected, formula);
  assert.equal(evaluate('=HLOOKUP(15,A1:C2,2)', { A1: '10', B1: '20', C1: '30', A2: 'low', B2: 'medium', C2: 'high' }), 'low');
  assert.equal(evaluate('=VLOOKUP("yes",A1:B2,2,FALSE)', { A1: 'yes', A2: 'no', B2: '=1/0' }), 0, 'an empty selected result is zero; unrelated return-cell errors do not leak');
  assert.equal(evaluate('=VLOOKUP("yes",A1:B2,2,FALSE)', { A1: 'yes', B1: '=1/0', A2: 'no' }), '#DIV/0!');
});

test('INDEX and MATCH compose for exact keys, positions, row/column selection and bounds', () => {
  const cells = { A1: 'apple', A2: 'pear', A3: 'plum', B1: '3', B2: '6', B3: '9', C1: '2', C2: '4', C3: '8', E1: '30', E2: '20', E3: '10' };
  for (const [formula, expected] of [
    ['=MATCH("PEAR",A1:A3,0)', 2], ['=MATCH("p*",A1:A3,0)', 2], ['=MATCH("missing",A1:A3,0)', '#N/A'],
    ['=MATCH(7,B1:B3)', 2], ['=MATCH(2,B1:B3,1)', '#N/A'], ['=MATCH(25,E1:E3,-1)', 1], ['=MATCH(5,E1:E3,-1)', 3],
    ['=MATCH(40,E1:E3,-1)', '#N/A'], ['=MATCH(6,B1:C3,0)', '#N/A'],
    ['=INDEX(B1:C3,MATCH("pear",A1:A3,0),2)', 4], ['=INDEX(B1:B3,2)', 6], ['=INDEX(A1:C1,2)', 3],
    ['=SUM(INDEX(B1:C3,0,2))', 14], ['=SUM(INDEX(B1:C3,2,0))', 10], ['=SUM(INDEX(B1:C3,2))', 10],
    ['=SUM(INDEX(B1:C3,0,0))', 32], ['=INDEX(B1:B3,2,0)', 6], ['=INDEX(B1:C3,4,1)', '#REF!'],
    ['=INDEX(B1:C3,1,3)', '#REF!'], ['=INDEX(B1:C3,-1,1)', '#VALUE!'],
  ]) assert.equal(evaluate(formula, cells), expected, formula);
  assert.equal(evaluate('=MATCH("a~*",A1:A2,0)', { A1: 'apple', A2: 'a*' }), 2);
  assert.equal(evaluate('=MATCH(20,A1:A4)', { A1: '10', A2: '20', A3: '20', A4: '30' }), 3);
});

test('XLOOKUP defaults to exact matching, preserves type and supports first/last search plus lazy fallback', () => {
  const cells = { A1: 'one', A2: 'two', A3: 'one', B1: '10', B2: '20', B3: '30', C1: "'10", C2: '10', D1: 'text', D2: 'number' };
  for (const [formula, expected] of [
    ['=XLOOKUP("ONE",A1:A3,B1:B3)', 10], ['=XLOOKUP("one",A1:A3,B1:B3,,0,-1)', 30],
    ['=XLOOKUP("none",A1:A3,B1:B3)', '#N/A'], ['=XLOOKUP("none",A1:A3,B1:B3,,0)', '#N/A'],
    ['=XLOOKUP("none",A1:A3,B1:B3,"Not listed")', 'Not listed'], ['=XLOOKUP("one",A1:A3,B1:B3,1/0)', 10],
    ['=XLOOKUP("none",A1:A3,B1:B3,1/0)', '#DIV/0!'], ['=XLOOKUP("none",A1:A3,B1:B3,"")', ''],
    ['=XLOOKUP("one",A1:A3,B1:B3,0,,)', 10], ['=XLOOKUP(10,C1:C2,D1:D2)', 'number'], ['=XLOOKUP("10",C1:C2,D1:D2)', 'text'],
    ['=XLOOKUP("one",A1:A3,B1:B2)', '#VALUE!'], ['=XLOOKUP("one",A1:B3,C1:C3)', '#VALUE!'],
    ['=XLOOKUP("one",A1:A3,B1:B3,,9)', '#VALUE!'], ['=XLOOKUP("one",A1:A3,B1:B3,,0,9)', '#VALUE!'],
    ['=_xlfn.XLOOKUP("two",A1:A3,B1:B3)', 20],
  ]) assert.equal(evaluate(formula, cells), expected, formula);
  assert.equal(evaluate('=XLOOKUP("yes",A1:A2,B1:B2,"missing")', { A1: 'yes', B1: '#N/A', A2: 'no', B2: '2' }), '#N/A', 'if_not_found does not hide errors from a matched result');
  assert.equal(evaluate('=XLOOKUP("yes",A1:A2,B1:B2,1/0)', { A1: 'yes', A2: 'no', B2: '=1/0' }), 0);
});

test('XLOOKUP supports approximate unsorted scans, horizontal results, wildcards and sorted binary searches', () => {
  const cells = { A1: '30', A2: '10', A3: '20', B1: 'large', B2: 'small', B3: 'medium', C1: '10', C2: '20', C3: '30', D1: 'small', D2: 'medium', D3: 'large', E1: '30', E2: '20', E3: '10', F1: 'large', F2: 'medium', F3: 'small' };
  for (const [formula, expected] of [
    ['=XLOOKUP(25,A1:A3,B1:B3,,-1)', 'medium'], ['=XLOOKUP(25,A1:A3,B1:B3,,1)', 'large'],
    ['=XLOOKUP(5,A1:A3,B1:B3,"none",-1)', 'none'], ['=XLOOKUP(35,A1:A3,B1:B3,"none",1)', 'none'],
    ['=XLOOKUP(20,C1:C3,D1:D3,,0,2)', 'medium'], ['=XLOOKUP(25,C1:C3,D1:D3,,-1,2)', 'medium'], ['=XLOOKUP(25,C1:C3,D1:D3,,1,2)', 'large'],
    ['=XLOOKUP(20,E1:E3,F1:F3,,0,-2)', 'medium'], ['=XLOOKUP(25,E1:E3,F1:F3,,-1,-2)', 'medium'], ['=XLOOKUP(25,E1:E3,F1:F3,,1,-2)', 'large'],
    ['=XLOOKUP(1,C1:C3,D1:D3,"none",-1,2)', 'none'], ['=XLOOKUP(99,E1:E3,F1:F3,"none",1,-2)', 'none'],
  ]) assert.equal(evaluate(formula, cells), expected, formula);
  const horizontal = { A1: 'Apple', B1: 'Apricot', C1: 'a*', A2: '2', B2: '3', C2: '4', A3: '20', B3: '30', C3: '40' };
  for (const [formula, expected] of [
    ['=XLOOKUP("Apricot",A1:C1,A2:C2)', 3], ['=XLOOKUP("a*",A1:C1,A2:C2)', 4], ['=XLOOKUP("ap*",A1:C1,A2:C2,,2)', 2],
    ['=XLOOKUP("ap*",A1:C1,A2:C2,,2,-1)', 3], ['=XLOOKUP("a~*",A1:C1,A2:C2,,2)', 4],
    ['=SUM(XLOOKUP("Apricot",A1:C1,A2:C3))', 33], ['=SUM(XLOOKUP("Apple",A1:A3,B1:C3))', 0],
  ]) assert.equal(evaluate(formula, horizontal), expected, formula);
});

test('SUMIFS and COUNTIFS combine criteria, validate full range shape and only propagate selected sum errors', () => {
  const cells = { A1: 'North', A2: 'South', A3: 'North', A4: 'North', A5: 'North', B1: 'Paid', B2: 'Paid', B3: 'Open', B4: 'Paid', B5: 'Open', C1: '10', C2: '20', C3: '30', C4: '40', C5: '=1/0' };
  for (const [formula, expected] of [
    ['=SUMIFS(C1:C5,A1:A5,"north",B1:B5,"Paid")', 50], ['=COUNTIFS(A1:A5,"N*",B1:B5,"Paid")', 2],
    ['=COUNTIFS(A1:A5,"North",C1:C5,">=20")', 2], ['=SUMIFS(C1:C5,A1:A5,"South")', 20],
    ['=SUMIFS(C1:C5,A1:A5,"Missing")', 0], ['=SUMIFS(C1:C5,A1:A5,"North",B1:B5,"Open")', '#DIV/0!'],
    ['=COUNTIFS(A1:A5,"Missing")', 0], ['=COUNTIFS(A1:A5,"North",B1:B4,"Paid")', '#VALUE!'],
    ['=SUMIFS(C1:C5,A1:A4,"North")', '#VALUE!'], ['=COUNTIFS(A1:A4,"North",A1:B2,"Paid")', '#VALUE!'],
    ['=COUNTIFS(A1:A5,"North",B1:B5)', '#VALUE!'], ['=SUMIFS(C1:C5,A1:A5)', '#VALUE!'],
    ['=COUNTIFS(1,1)', '#VALUE!'], ['=SUMIFS(C1,C1,">0")', 10], ['=COUNTIFS(A6:A8,"")', 3],
  ]) assert.equal(evaluate(formula, cells), expected, formula);
  assert.equal(evaluate('=COUNTIFS(A1:A3,"a~*",B1:B3,TRUE)', { A1: 'a*', A2: 'apple', A3: 'a*', B1: 'TRUE', B2: 'TRUE', B3: 'FALSE' }), 1);
});

test('lookup and multiple-criteria functions work across quoted sheets and remain fill-compatible', () => {
  const sheets = [
    { id: 'sales', name: "Sales '26", cells: { A1: 'P01', A2: 'P02', A3: 'P03', B1: 'North', B2: 'South', B3: 'North', C1: '10', C2: '20', C3: '30' } },
    { id: 'report', name: 'Report', cells: { A1: 'P02', A2: "=XLOOKUP(A1,'Sales ''26'!$A$1:$A$3,'Sales ''26'!$C$1:$C$3)", A3: "=SUMIFS('Sales ''26'!C1:C3,'Sales ''26'!B1:B3,\"North\")", A4: "=INDEX('Sales ''26'!C1:C3,MATCH(A1,'Sales ''26'!A1:A3,0))" } },
  ];
  const f = calculateCells(sheets[1].cells, { sheets, currentSheetId: 'report' });
  assert.equal(f('A2'), 20); assert.equal(f('A3'), 40); assert.equal(f('A4'), 20);
  assert.equal(shiftFormulaReferences(sheets[1].cells.A2, 1, 0), "=XLOOKUP(A2,'Sales ''26'!$A$1:$A$3,'Sales ''26'!$C$1:$C$3)");
});

test('function help offers unique supported examples instead of suggesting unknown formula names', () => {
  assert.equal(new Set(FORMULA_CATALOG.map(entry => entry.name)).size, FORMULA_CATALOG.length);
  for (const entry of FORMULA_CATALOG) {
    assert.ok(entry.description && entry.category && entry.signature.startsWith(`${entry.name}(`), entry.name);
    assert.notEqual(evaluate(entry.example), '#NAME?', `${entry.name} advertises an unsupported example`);
  }
  for (const name of ['VLOOKUP', 'HLOOKUP', 'XLOOKUP', 'INDEX', 'MATCH', 'COUNTIFS', 'SUMIFS']) assert.ok(FORMULA_CATALOG.some(entry => entry.name === name));
});
