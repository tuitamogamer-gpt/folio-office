import type { SheetTab } from '../types';

type Scalar = string | number | boolean | null | FormulaError;
type Value = Scalar | { kind: 'reference'; value: Scalar } | { kind: 'range'; rows: Scalar[][] };
type Token = { kind: 'number' | 'string' | 'name' | 'quoted' | 'symbol' | 'error'; text: string };
type Node = { kind: 'literal'; value: Scalar } | { kind: 'reference'; sheet?: string; ref: string } | { kind: 'range'; start: Extract<Node, { kind: 'reference' }>; end: Extract<Node, { kind: 'reference' }> } | { kind: 'unary'; op: string; node: Node } | { kind: 'binary'; op: string; left: Node; right: Node } | { kind: 'call'; name: string; args: Node[] };
export interface FormulaContext { sheets: SheetTab[]; currentSheetId: string; }
export interface FormulaDefinition { name: string; signature: string; description: string; category: string; example: string; }
export const FORMULA_CATALOG: readonly FormulaDefinition[] = [
  { name: 'SUM', signature: 'SUM(number1, [number2], …)', description: 'Add numbers and ranges.', category: 'Math', example: '=SUM(B2:B10)' },
  { name: 'AVERAGE', signature: 'AVERAGE(number1, [number2], …)', description: 'Find the arithmetic mean of numeric values.', category: 'Statistics', example: '=AVERAGE(B2:B10)' },
  { name: 'MIN', signature: 'MIN(number1, [number2], …)', description: 'Find the smallest numeric value.', category: 'Statistics', example: '=MIN(B2:B10)' },
  { name: 'MAX', signature: 'MAX(number1, [number2], …)', description: 'Find the largest numeric value.', category: 'Statistics', example: '=MAX(B2:B10)' },
  { name: 'COUNT', signature: 'COUNT(value1, [value2], …)', description: 'Count cells containing numbers.', category: 'Statistics', example: '=COUNT(B2:B10)' },
  { name: 'COUNTA', signature: 'COUNTA(value1, [value2], …)', description: 'Count nonempty cells, including text and errors.', category: 'Statistics', example: '=COUNTA(A2:A10)' },
  { name: 'COUNTIF', signature: 'COUNTIF(range, criteria)', description: 'Count matching cells; criteria support comparisons and wildcards.', category: 'Statistics', example: '=COUNTIF(B2:B10,">=100")' },
  { name: 'COUNTIFS', signature: 'COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], …)', description: 'Count rows matching every condition across ranges of the same size.', category: 'Statistics', example: '=COUNTIFS(A2:A10,"Paid",B2:B10,">100")' },
  { name: 'SUMIF', signature: 'SUMIF(range, criteria, [sum_range])', description: 'Add numbers whose corresponding cells match a condition.', category: 'Math', example: '=SUMIF(A2:A10,"Paid",B2:B10)' },
  { name: 'SUMIFS', signature: 'SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], …)', description: 'Add numbers matching every condition across ranges of the same size.', category: 'Math', example: '=SUMIFS(C2:C10,A2:A10,"Paid",B2:B10,">100")' },
  { name: 'IF', signature: 'IF(logical_test, value_if_true, [value_if_false])', description: 'Choose a result; only the selected branch is calculated.', category: 'Logical', example: '=IF(B2>=100,"On target","Below target")' },
  { name: 'IFERROR', signature: 'IFERROR(value, value_if_error)', description: 'Use a fallback when a calculation returns an error.', category: 'Logical', example: '=IFERROR(B2/C2,0)' },
  { name: 'AND', signature: 'AND(logical1, [logical2], …)', description: 'Return TRUE when every condition is true.', category: 'Logical', example: '=AND(B2>0,C2="Paid")' },
  { name: 'OR', signature: 'OR(logical1, [logical2], …)', description: 'Return TRUE when at least one condition is true.', category: 'Logical', example: '=OR(B2>100,C2="Priority")' },
  { name: 'NOT', signature: 'NOT(logical)', description: 'Reverse TRUE and FALSE.', category: 'Logical', example: '=NOT(B2=0)' },
  { name: 'ROUND', signature: 'ROUND(number, num_digits)', description: 'Round to a number of decimal places.', category: 'Math', example: '=ROUND(B2,2)' },
  { name: 'ROUNDUP', signature: 'ROUNDUP(number, num_digits)', description: 'Round away from zero.', category: 'Math', example: '=ROUNDUP(B2,2)' },
  { name: 'ROUNDDOWN', signature: 'ROUNDDOWN(number, num_digits)', description: 'Round toward zero.', category: 'Math', example: '=ROUNDDOWN(B2,2)' },
  { name: 'ABS', signature: 'ABS(number)', description: 'Return the absolute value.', category: 'Math', example: '=ABS(B2)' },
  { name: 'SQRT', signature: 'SQRT(number)', description: 'Find the square root of a nonnegative number.', category: 'Math', example: '=SQRT(B2)' },
  { name: 'POWER', signature: 'POWER(number, power)', description: 'Raise a number to a power.', category: 'Math', example: '=POWER(B2,2)' },
  { name: 'MOD', signature: 'MOD(number, divisor)', description: 'Return the remainder with the divisor’s sign.', category: 'Math', example: '=MOD(B2,7)' },
  { name: 'MEDIAN', signature: 'MEDIAN(number1, [number2], …)', description: 'Find the middle numeric value.', category: 'Statistics', example: '=MEDIAN(B2:B10)' },
  { name: 'LEN', signature: 'LEN(text)', description: 'Count characters in text.', category: 'Text', example: '=LEN(A2)' },
  { name: 'LEFT', signature: 'LEFT(text, [num_chars])', description: 'Take characters from the beginning of text.', category: 'Text', example: '=LEFT(A2,3)' },
  { name: 'RIGHT', signature: 'RIGHT(text, [num_chars])', description: 'Take characters from the end of text.', category: 'Text', example: '=RIGHT(A2,3)' },
  { name: 'MID', signature: 'MID(text, start_num, num_chars)', description: 'Extract text starting at a position, counting from 1.', category: 'Text', example: '=MID(A2,2,4)' },
  { name: 'CONCAT', signature: 'CONCAT(text1, [text2], …)', description: 'Join text and ranges.', category: 'Text', example: '=CONCAT(A2," ",B2)' },
  { name: 'CONCATENATE', signature: 'CONCATENATE(text1, [text2], …)', description: 'Join text values.', category: 'Text', example: '=CONCATENATE(A2," ",B2)' },
  { name: 'TRIM', signature: 'TRIM(text)', description: 'Remove extra spaces between and around words.', category: 'Text', example: '=TRIM(A2)' },
  { name: 'UPPER', signature: 'UPPER(text)', description: 'Convert text to uppercase.', category: 'Text', example: '=UPPER(A2)' },
  { name: 'LOWER', signature: 'LOWER(text)', description: 'Convert text to lowercase.', category: 'Text', example: '=LOWER(A2)' },
  { name: 'TODAY', signature: 'TODAY()', description: 'Return today’s Excel date number; apply Date formatting to display it.', category: 'Date', example: '=TODAY()' },
  { name: 'DATE', signature: 'DATE(year, month, day)', description: 'Build an Excel date number from a year, month and day.', category: 'Date', example: '=DATE(2026,10,1)' },
  { name: 'VLOOKUP', signature: 'VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])', description: 'Search the first column. FALSE gives an exact match; TRUE or omitted requires ascending keys for an approximate match.', category: 'Lookup', example: '=VLOOKUP(A2,F2:H10,3,FALSE)' },
  { name: 'HLOOKUP', signature: 'HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])', description: 'Search the first row. FALSE gives an exact match; TRUE or omitted requires ascending keys for an approximate match.', category: 'Lookup', example: '=HLOOKUP(A2,F1:J3,3,FALSE)' },
  { name: 'INDEX', signature: 'INDEX(array, row_num, [column_num])', description: 'Return a value by row and column. Zero selects a whole row or column for another function.', category: 'Lookup', example: '=INDEX(B2:D10,3,2)' },
  { name: 'MATCH', signature: 'MATCH(lookup_value, lookup_array, [match_type])', description: 'Find a position. Use 0 for exact, 1 for ascending approximate, or -1 for descending approximate.', category: 'Lookup', example: '=MATCH(A2,F2:F10,0)' },
  { name: 'XLOOKUP', signature: 'XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])', description: 'Exact match by default; match modes -1/1 choose the next smaller/larger value and 2 enables wildcards. Search 1/-1 scans first/last; 2/-2 requires sorted ascending/descending keys and excludes wildcards.', category: 'Lookup', example: '=XLOOKUP(A2,F2:F10,G2:G10,"Not found")' },
];
class FormulaError extends Error {}
const fail = (code = '#VALUE!'): never => { throw new FormulaError(code); };
const MAX_ROW = 1048576;
const MAX_COL = 16384;
const ERROR_PATTERN = /^#(?:REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|NULL!|CYCLE!|DEPTH!)/i;
const isStoredNumber = (value: string) => /^[+-]?(?:0|[1-9]\d*)(?:\.\d*)?(?:e[+-]?\d+)?$/i.test(value) && Number.isFinite(Number(value)) && value.replace(/\D/g, '').length <= 15;
const columnName = (n: number): string => { let name = ''; for (n++; n; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + (n - 1) % 26) + name; return name; };
function pointFor(ref: string) {
  const match = ref.match(/^\$?([A-Z]+)\$?([1-9]\d*)$/i);
  if (!match) return fail('#REF!');
  const col = [...match[1].toUpperCase()].reduce((sum, c) => sum * 26 + c.charCodeAt(0) - 64, 0) - 1;
  const row = Number(match[2]) - 1;
  if (row >= MAX_ROW || col >= MAX_COL || !Number.isSafeInteger(row)) return fail('#REF!');
  return { row, col };
}
const canonicalRef = (ref: string) => { const point = pointFor(ref); return `${columnName(point.col)}${point.row + 1}`; };

function tokenize(source: string): Token[] {
  const tokens: Token[] = [];
  for (let position = 0; position < source.length;) {
    const rest = source.slice(position);
    if (/^\s/.test(rest)) { position++; continue; }
    const first = rest[0];
    if (first === '"' || first === "'") {
      let value = ''; let closed = false; position++;
      while (position < source.length) {
        const char = source[position++];
        if (char === first) { if (source[position] === first) { value += first; position++; } else { closed = true; break; } } else value += char;
      }
      if (!closed) fail();
      tokens.push({ kind: first === '"' ? 'string' : 'quoted', text: value });
    } else {
      const error = rest.match(ERROR_PATTERN);
      const number = rest.match(/^(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?/i);
      const name = rest.match(/^\$?[\p{L}_][\p{L}\p{N}_.$]*/u);
      const symbol = rest.match(/^(?:<>|<=|>=|[+\-*/^%(),:!&=<>])/);
      const token = error ? { kind: 'error' as const, text: error[0] } : number ? { kind: 'number' as const, text: number[0] } : name ? { kind: 'name' as const, text: name[0] } : symbol ? { kind: 'symbol' as const, text: symbol[0] } : null;
      if (!token) return fail();
      tokens.push(token); position += token.text.length;
    }
    if (tokens.length > 4000) fail('#NUM!');
  }
  return tokens;
}

function parse(source: string): Node {
  const tokens = tokenize(source); let position = 0; let depth = 0;
  const peek = (text: string) => tokens[position]?.text === text;
  const accept = (text: string) => peek(text) ? (position++, true) : false;
  const expect = (text: string) => { if (!accept(text)) fail(); };
  const reference = (): Extract<Node, { kind: 'reference' }> => {
    let token = tokens[position++]; if (!token) return fail();
    let sheet: string | undefined;
    if (accept('!')) { sheet = token.text; token = tokens[position++]; }
    if (!token || token.kind !== 'name' || !/^\$?[A-Z]+\$?[1-9]\d*$/i.test(token.text)) return fail('#REF!');
    pointFor(token.text);
    return { kind: 'reference', sheet, ref: token.text };
  };
  const primary = (): Node => {
    if (++depth > 200) fail('#DEPTH!');
    try {
      const token = tokens[position]; if (!token) return fail();
      if (accept('(')) { const result = comparison(); expect(')'); return result; }
      if (token.kind === 'number') { position++; return { kind: 'literal', value: Number(token.text) }; }
      if (token.kind === 'string') { position++; return { kind: 'literal', value: token.text }; }
      if (token.kind === 'error') { position++; return { kind: 'literal', value: new FormulaError(token.text.toUpperCase()) }; }
      if (token.kind !== 'name' && token.kind !== 'quoted') return fail();
      if (tokens[position + 1]?.text === '(') {
        position += 2; const args: Node[] = [];
        if (!peek(')')) {
          while (true) {
            args.push(peek(',') || peek(')') ? { kind: 'literal', value: null } : comparison());
            if (!accept(',')) break;
          }
        }
        expect(')'); return { kind: 'call', name: token.text.toUpperCase(), args };
      }
      if (tokens[position + 1]?.text !== '!' && /^(TRUE|FALSE)$/i.test(token.text)) { position++; return { kind: 'literal', value: token.text.toUpperCase() === 'TRUE' }; }
      if (tokens[position + 1]?.text === '!' || /^\$?[A-Z]+\$?[1-9]\d*$/i.test(token.text)) {
        const start = reference();
        if (accept(':')) { const end = reference(); return { kind: 'range', start, end: { ...end, sheet: end.sheet ?? start.sheet } }; }
        return start;
      }
      position++; return { kind: 'literal', value: new FormulaError('#NAME?') };
    } finally { depth--; }
  };
  const unary = (): Node => {
    if (peek('+') || peek('-')) { const op = tokens[position++].text; if (++depth > 200) fail('#DEPTH!'); try { return { kind: 'unary', op, node: unary() }; } finally { depth--; } }
    return primary();
  };
  const percent = (): Node => { let result = unary(); while (accept('%')) result = { kind: 'unary', op: '%', node: result }; return result; };
  const binary = (next: () => Node, operators: string[]): Node => { let result = next(); while (operators.includes(tokens[position]?.text)) { const op = tokens[position++].text; result = { kind: 'binary', op, left: result, right: next() }; } return result; };
  const power = (): Node => binary(percent, ['^']);
  const product = (): Node => binary(power, ['*', '/']);
  const addition = (): Node => binary(product, ['+', '-']);
  const concatenate = (): Node => binary(addition, ['&']);
  const comparison = (): Node => binary(concatenate, ['=', '<>', '<', '>', '<=', '>=']);
  const node = comparison(); if (position !== tokens.length) fail(); return node;
}

function scalar(value: Value): Scalar {
  if (value && typeof value === 'object' && !(value instanceof FormulaError)) { if (value.kind === 'reference') return value.value; return fail(); }
  return value;
}
function checked(value: Value): Exclude<Scalar, FormulaError> { const result = scalar(value); if (result instanceof FormulaError) throw result; return result; }
function number(value: Value): number {
  const result = checked(value);
  if (typeof result === 'number') return result;
  if (result == null || result === '') return 0;
  if (typeof result === 'boolean') return result ? 1 : 0;
  if (!result.trim() || !/^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(result.trim())) return fail();
  const parsed = Number(result); return Number.isFinite(parsed) ? parsed : fail('#NUM!');
}
function text(value: Value): string { const result = checked(value); return result == null ? '' : typeof result === 'boolean' ? (result ? 'TRUE' : 'FALSE') : String(result); }
function truth(value: Value): boolean {
  const result = checked(value);
  if (typeof result === 'string') { if (/^(TRUE|FALSE)$/i.test(result)) return result.toUpperCase() === 'TRUE'; return fail(); }
  return Boolean(result);
}
function values(value: Value): Scalar[] { return value && typeof value === 'object' && !(value instanceof FormulaError) ? value.kind === 'range' ? value.rows.flat() : [value.value] : [value]; }
function matrix(value: Value, requireRange = false): Scalar[][] {
  if (value && typeof value === 'object' && !(value instanceof FormulaError)) return value.kind === 'range' ? value.rows : [[value.value]];
  checked(value);
  if (requireRange) return fail();
  return [[value as Scalar]];
}
function vector(value: Value): Scalar[] {
  const rows = matrix(value);
  if (rows.length !== 1 && rows[0].length !== 1) return fail('#N/A');
  return rows.flat();
}
function wildcardPattern(pattern: string): RegExp {
  let source = '^';
  for (let index = 0; index < pattern.length; index++) {
    const char = pattern[index];
    if (char === '~' && index + 1 < pattern.length && /[~*?]/.test(pattern[index + 1])) source += pattern[++index].replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    else if (char === '*') source += '.*';
    else if (char === '?') source += '.';
    else source += char.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp(`${source}$`, 'is');
}
function lookupDifference(candidate: Scalar, lookup: Exclude<Scalar, FormulaError>): number | null {
  if (candidate instanceof FormulaError) throw candidate;
  if (lookup === null) lookup = 0;
  if (candidate === null) candidate = typeof lookup === 'string' ? '' : typeof lookup === 'boolean' ? false : 0;
  return typeof candidate === typeof lookup ? compare(candidate, lookup) : null;
}
function lookupIndex(lookup: Exclude<Scalar, FormulaError>, entries: Scalar[], mode: 'exact' | 'wildcard' | 'smaller' | 'larger', reverse = false, lastApproximate = false): number {
  const wildcard = mode === 'wildcard' && typeof lookup === 'string' ? wildcardPattern(lookup) : undefined;
  let closest = -1;
  for (let step = 0; step < entries.length; step++) {
    const index = reverse ? entries.length - step - 1 : step;
    const candidate = entries[index];
    const difference = lookupDifference(candidate, lookup);
    const exact = wildcard ? typeof candidate === 'string' && wildcard.test(candidate) : difference === 0;
    if (exact && (!lastApproximate || mode === 'exact' || mode === 'wildcard')) return index;
    if (difference === null || (mode !== 'smaller' && mode !== 'larger')) continue;
    if ((mode === 'smaller' && difference > 0) || (mode === 'larger' && difference < 0)) continue;
    const improvement = closest < 0 ? 0 : compare(candidate, entries[closest]);
    if (closest < 0 || (mode === 'smaller' ? improvement > 0 : improvement < 0) || (lastApproximate && improvement === 0)) closest = index;
  }
  return closest;
}
function binaryLookupIndex(lookup: Exclude<Scalar, FormulaError>, entries: Scalar[], mode: number, descending: boolean): number {
  if (mode === 2) return fail(); // Wildcards cannot define the ordering required by a binary search.
  let low = 0; let high = entries.length - 1;
  while (low <= high) {
    const middle = Math.floor((low + high) / 2);
    const difference = lookupDifference(entries[middle], lookup);
    if (difference === 0) return middle;
    const order = difference ?? compare(entries[middle], lookup);
    if (descending ? order > 0 : order < 0) low = middle + 1; else high = middle - 1;
  }
  if (mode === 0) return -1;
  const index = mode === -1 ? descending ? low : high : descending ? high : low;
  return index >= 0 && index < entries.length && lookupDifference(entries[index], lookup) !== null ? index : -1;
}
function compare(left: Value, right: Value): number {
  let a = checked(left); let b = checked(right);
  if (a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  if (typeof a !== typeof b) { const rank = (value: unknown) => typeof value === 'number' ? 0 : typeof value === 'string' ? 1 : 2; return rank(a) - rank(b); }
  if (typeof a === 'string' && typeof b === 'string') { a = a.toLocaleLowerCase(); b = b.toLocaleLowerCase(); }
  return a === b ? 0 : a < b ? -1 : 1;
}
const finite = (result: number) => Number.isFinite(result) ? result : fail('#NUM!');
function shiftDecimal(value: number, digits: number) {
  const [mantissa, exponent] = value.toExponential().split('e');
  return Number(`${mantissa}e${Number(exponent) + digits}`);
}
function criteriaMatcher(criteria: Value): (value: Scalar) => boolean {
  const criterion = checked(criteria);
  if (typeof criterion !== 'string') return value => !(value instanceof FormulaError) && compare(value, criterion) === 0;
  const match = criterion.match(/^(<>|<=|>=|=|<|>)(.*)$/s);
  const op = match?.[1] || '='; const desired = match?.[2] ?? criterion;
  const numeric = desired.trim() !== '' && /^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(desired.trim()) ? Number(desired) : null;
  const wildcard = wildcardPattern(desired);
  return value => {
    if (value instanceof FormulaError) return false;
    let comparison: number;
    if (numeric !== null) { if (typeof value === 'number') comparison = value - numeric; else if (value === null && numeric === 0) comparison = 0; else if (typeof value === 'string' && value.trim() && Number.isFinite(Number(value))) comparison = Number(value) - numeric; else return op === '<>'; }
    else if (op === '=' || op === '<>') { const matches = typeof value === 'string' ? wildcard.test(value) : value === null ? desired === '' : typeof value === 'boolean' && desired.toUpperCase() === String(value).toUpperCase(); return op === '=' ? matches : !matches; }
    else { if (typeof value !== 'string') return false; comparison = compare(value, desired); }
    return op === '=' ? comparison === 0 : op === '<>' ? comparison !== 0 : op === '<' ? comparison < 0 : op === '>' ? comparison > 0 : op === '<=' ? comparison <= 0 : comparison >= 0;
  };
}
function serialDate(year: number, month: number, day: number) {
  const date = new Date(0); date.setUTCHours(0, 0, 0, 0); date.setUTCFullYear(year, month - 1, day);
  const serial = (date.getTime() - Date.UTC(1899, 11, 31)) / 86400000;
  return finite(serial + (date.getTime() >= Date.UTC(1900, 2, 1) ? 1 : 0));
}

/** Safe spreadsheet parser with cached, cycle-checked evaluation. Formula text is never executed as JavaScript. */
export function calculateCells(cells: Record<string, string>, context?: FormulaContext): (key: string) => string | number {
  const currentSheetId = context?.currentSheetId || '__current';
  const sheets = context?.sheets?.length ? context.sheets : [{ id: currentSheetId, name: 'Sheet 1', cells }];
  const byId = new Map(sheets.map(sheet => [sheet.id, sheet]));
  const byName = new Map(sheets.map(sheet => [sheet.name.toLocaleLowerCase(), sheet.id]));
  const cache = new Map<string, Scalar>();
  const resolveSheet = (name: string | undefined, fallback: string) => name === undefined ? fallback : byName.get(name.toLocaleLowerCase()) || fail('#REF!');
  const calculate = (key: string, sheetId: string, ancestors: Set<string>): Scalar => {
    let cacheKey = '';
    try {
      key = canonicalRef(key); cacheKey = `${sheetId}\0${key}`;
      if (cache.has(cacheKey)) return cache.get(cacheKey)!;
      if (ancestors.has(cacheKey)) return fail('#CYCLE!');
      if (ancestors.size >= 200) return fail('#DEPTH!');
      const sheetCells = sheetId === currentSheetId ? cells : byId.get(sheetId)?.cells;
      if (!sheetCells) return fail('#REF!');
      const raw = String(sheetCells[key] ?? '');
      if (!raw) return null;
      if (!raw.startsWith('=')) {
        if (raw.startsWith("'")) return raw.slice(1);
        if (ERROR_PATTERN.test(raw) && raw.match(ERROR_PATTERN)?.[0].length === raw.length) return new FormulaError(raw.toUpperCase());
        return isStoredNumber(raw) ? Number(raw) : /^(TRUE|FALSE)$/i.test(raw) ? raw.toUpperCase() === 'TRUE' : raw;
      }
      const visiting = new Set(ancestors); visiting.add(cacheKey);
      const evaluate = (node: Node, evaluationDepth = 0): Value => {
        if (evaluationDepth > 500) return fail('#DEPTH!');
        const run = (child: Node) => evaluate(child, evaluationDepth + 1);
        if (node.kind === 'literal') return typeof node.value === 'number' ? finite(node.value) : node.value;
        if (node.kind === 'reference') return { kind: 'reference', value: calculate(node.ref, resolveSheet(node.sheet, sheetId), visiting) };
        if (node.kind === 'range') {
          const startSheet = resolveSheet(node.start.sheet, sheetId); const endSheet = resolveSheet(node.end.sheet, sheetId);
          if (startSheet !== endSheet) return fail('#REF!');
          const start = pointFor(node.start.ref); const end = pointFor(node.end.ref);
          const rowCount = Math.abs(end.row - start.row) + 1; const colCount = Math.abs(end.col - start.col) + 1;
          if (rowCount * colCount > 100000) return fail('#NUM!');
          const rows: Scalar[][] = [];
          for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row++) {
            const line: Scalar[] = [];
            for (let col = Math.min(start.col, end.col); col <= Math.max(start.col, end.col); col++) line.push(calculate(`${columnName(col)}${row + 1}`, startSheet, visiting));
            rows.push(line);
          }
          return { kind: 'range', rows };
        }
        if (node.kind === 'unary') { const value = number(run(node.node)); return node.op === '-' ? -value : node.op === '%' ? value / 100 : value; }
        if (node.kind === 'binary') {
          const left = run(node.left); const right = run(node.right);
          if (node.op === '&') return text(left) + text(right);
          if (['=', '<>', '<', '>', '<=', '>='].includes(node.op)) { const difference = compare(left, right); return node.op === '=' ? difference === 0 : node.op === '<>' ? difference !== 0 : node.op === '<' ? difference < 0 : node.op === '>' ? difference > 0 : node.op === '<=' ? difference <= 0 : difference >= 0; }
          const a = number(left); const b = number(right);
          if (node.op === '/' && b === 0) return fail('#DIV/0!');
          return finite(node.op === '+' ? a + b : node.op === '-' ? a - b : node.op === '*' ? a * b : node.op === '/' ? a / b : a ** b);
        }
        const name = node.name.replace(/^_XLFN\.(?:_XLWS\.)?/, '');
        const { args } = node;
        const arity = (min: number, max = min) => { if (args.length < min || args.length > max) fail(); };
        const arg = (index: number, fallback: Scalar = null) => args[index] ? run(args[index]) : fallback;
        if (name === 'IF') { arity(2, 3); return truth(arg(0)) ? arg(1) : args.length === 3 ? arg(2) : false; }
        if (name === 'IFERROR') { arity(2); try { const value = arg(0); checked(value); return value; } catch (error) { if (!(error instanceof FormulaError)) throw error; return arg(1); } }
        if (['SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA', 'MEDIAN'].includes(name)) {
          arity(name === 'SUM' ? 0 : 1, 255); const numeric: number[] = []; let count = 0;
          for (const nodeArg of args) {
            let value: Value;
            try { value = run(nodeArg); } catch (error) { if (name !== 'COUNT' && name !== 'COUNTA') throw error; value = error instanceof FormulaError ? error : new FormulaError('#VALUE!'); }
            const referenced = value !== null && typeof value === 'object' && !(value instanceof FormulaError);
            for (const item of values(value)) {
              if (item !== null) count++;
              if (name === 'COUNTA') continue;
              if (item instanceof FormulaError) { if (name === 'COUNT' && referenced) continue; throw item; }
              if (typeof item === 'number') numeric.push(item);
              else if (!referenced && item !== null) { if (name === 'COUNT') { try { numeric.push(number(item)); } catch { /* COUNT ignores text that cannot become a number. */ } } else numeric.push(number(item)); }
            }
          }
          if (name === 'COUNTA') return count;
          if (name === 'COUNT') return numeric.length;
          if (name === 'SUM') return finite(numeric.reduce((sum, value) => sum + value, 0));
          if (!numeric.length) return name === 'MIN' || name === 'MAX' ? 0 : fail('#DIV/0!');
          if (name === 'AVERAGE') return finite(numeric.reduce((sum, value) => sum + value, 0) / numeric.length);
          if (name === 'MIN') return numeric.reduce((min, value) => Math.min(min, value), Infinity);
          if (name === 'MAX') return numeric.reduce((max, value) => Math.max(max, value), -Infinity);
          numeric.sort((a, b) => a - b); const middle = Math.floor(numeric.length / 2); return numeric.length % 2 ? numeric[middle] : (numeric[middle - 1] + numeric[middle]) / 2;
        }
        if (name === 'COUNTIF' || name === 'SUMIF') {
          arity(2, name === 'SUMIF' ? 3 : 2);
          const input = arg(0); const entries = values(input); const matches = criteriaMatcher(arg(1));
          if (name === 'COUNTIF') return entries.filter(matches).length;
          const sums = args.length === 3 ? values(arg(2)) : entries;
          if (sums.length !== entries.length) return fail();
          return finite(entries.reduce<number>((sum, value, index) => { if (!matches(value)) return sum; const add = sums[index]; if (add instanceof FormulaError) throw add; return sum + (typeof add === 'number' ? add : 0); }, 0));
        }
        if (name === 'COUNTIFS' || name === 'SUMIFS') {
          const start = name === 'SUMIFS' ? 1 : 0;
          arity(start + 2, start + 254);
          if ((args.length - start) % 2 !== 0) return fail();
          const sumRows = name === 'SUMIFS' ? matrix(arg(0), true) : undefined;
          let height = sumRows?.length; let width = sumRows?.[0].length;
          const conditions: Array<{ cells: Scalar[]; matches: (value: Scalar) => boolean }> = [];
          for (let index = start; index < args.length; index += 2) {
            const rows = matrix(arg(index), true);
            height ??= rows.length; width ??= rows[0].length;
            if (rows.length !== height || rows[0].length !== width) return fail();
            conditions.push({ cells: rows.flat(), matches: criteriaMatcher(arg(index + 1)) });
          }
          const sums = sumRows?.flat(); let result = 0;
          for (let index = 0; index < conditions[0].cells.length; index++) {
            if (!conditions.every(condition => condition.matches(condition.cells[index]))) continue;
            if (!sums) { result++; continue; }
            const value = sums[index]; if (value instanceof FormulaError) throw value;
            if (typeof value === 'number') result += value;
          }
          return finite(result);
        }
        if (name === 'INDEX') {
          arity(2, 3);
          const rows = matrix(arg(0));
          let row = Math.trunc(number(arg(1))); let column = args.length === 3 ? Math.trunc(number(arg(2))) : rows[0].length > 1 ? 0 : 1;
          if (args.length === 2 && rows.length === 1) { column = row; row = 1; }
          if (row < 0 || column < 0) return fail();
          if (row > rows.length || column > rows[0].length) return fail('#REF!');
          if (row === 0 || column === 0) {
            const selected = row === 0 ? column === 0 ? rows : rows.map(line => [line[column - 1]]) : [rows[row - 1]];
            return selected.length === 1 && selected[0].length === 1 ? { kind: 'reference', value: selected[0][0] } : { kind: 'range', rows: selected };
          }
          return { kind: 'reference', value: rows[row - 1][column - 1] };
        }
        if (name === 'MATCH') {
          arity(2, 3);
          const lookup = checked(arg(0)); const entries = vector(arg(1)); const matchType = number(arg(2, 1));
          if (![0, 1, -1].includes(matchType)) return fail();
          const index = lookupIndex(lookup, entries, matchType === 0 ? 'wildcard' : matchType === 1 ? 'smaller' : 'larger', false, matchType !== 0);
          return index < 0 ? fail('#N/A') : index + 1;
        }
        if (name === 'VLOOKUP' || name === 'HLOOKUP') {
          arity(3, 4);
          const lookup = checked(arg(0)); const rows = matrix(arg(1)); const ordinal = Math.trunc(number(arg(2)));
          const vertical = name === 'VLOOKUP'; const approximate = truth(arg(3, true));
          if (ordinal < 1) return fail();
          if (ordinal > (vertical ? rows[0].length : rows.length)) return fail('#REF!');
          const index = lookupIndex(lookup, vertical ? rows.map(line => line[0]) : rows[0], approximate ? 'smaller' : 'wildcard', false, approximate);
          if (index < 0) return fail('#N/A');
          return { kind: 'reference', value: vertical ? rows[index][ordinal - 1] : rows[ordinal - 1][index] };
        }
        if (name === 'XLOOKUP') {
          arity(3, 6);
          const lookup = checked(arg(0)); const lookupRows = matrix(arg(1)); const returnRows = matrix(arg(2));
          if (lookupRows.length !== 1 && lookupRows[0].length !== 1) return fail();
          const vertical = lookupRows[0].length === 1;
          if (vertical ? returnRows.length !== lookupRows.length : returnRows[0].length !== lookupRows[0].length) return fail();
          const entries = lookupRows.flat();
          const omitted = (index: number) => !args[index] || (args[index].kind === 'literal' && args[index].value === null);
          const matchMode = omitted(4) ? 0 : number(arg(4)); const searchMode = omitted(5) ? 1 : number(arg(5));
          if (![0, -1, 1, 2].includes(matchMode) || ![1, -1, 2, -2].includes(searchMode)) return fail();
          const index = Math.abs(searchMode) === 2 ? binaryLookupIndex(lookup, entries, matchMode, searchMode === -2) : lookupIndex(lookup, entries, matchMode === 0 ? 'exact' : matchMode === 2 ? 'wildcard' : matchMode === -1 ? 'smaller' : 'larger', searchMode === -1);
          if (index < 0) return !omitted(3) ? arg(3) : fail('#N/A');
          const selected = vertical ? [returnRows[index]] : returnRows.map(line => [line[index]]);
          return selected.length === 1 && selected[0].length === 1 ? { kind: 'reference', value: selected[0][0] } : { kind: 'range', rows: selected };
        }
        if (name === 'AND' || name === 'OR') {
          arity(1, 255); const bools: boolean[] = [];
          for (const nodeArg of args) { const value = run(nodeArg); const referenced = value !== null && typeof value === 'object' && !(value instanceof FormulaError); for (const item of values(value)) { if (referenced && (item === null || typeof item === 'string')) continue; bools.push(truth(item)); } }
          if (!bools.length) return fail(); return name === 'AND' ? bools.every(Boolean) : bools.some(Boolean);
        }
        if (name === 'NOT') { arity(1); return !truth(arg(0)); }
        if (['ROUND', 'ROUNDUP', 'ROUNDDOWN'].includes(name)) {
          arity(2); const value = number(arg(0)); const digits = Math.trunc(number(arg(1)));
          if (Math.abs(digits) > 308) return digits > 0 ? value : 0;
          const scaled = shiftDecimal(Math.abs(value), digits);
          if (!Number.isFinite(scaled)) return value;
          const result = name === 'ROUND' ? Math.floor(scaled + 0.5 + Number.EPSILON * scaled) : name === 'ROUNDUP' ? Math.ceil(scaled) : Math.floor(scaled);
          return finite(Math.sign(value) * shiftDecimal(result, -digits));
        }
        if (name === 'ABS' || name === 'SQRT') { arity(1); const value = number(arg(0)); if (name === 'SQRT' && value < 0) return fail('#NUM!'); return name === 'ABS' ? Math.abs(value) : Math.sqrt(value); }
        if (name === 'POWER' || name === 'MOD') { arity(2); const a = number(arg(0)); const b = number(arg(1)); if (name === 'MOD' && b === 0) return fail('#DIV/0!'); return finite(name === 'POWER' ? a ** b : a - b * Math.floor(a / b)); }
        if (['LEN', 'TRIM', 'UPPER', 'LOWER'].includes(name)) { arity(1); const value = text(arg(0)); return name === 'LEN' ? value.length : name === 'TRIM' ? value.replace(/^ +| +$/g, '').replace(/ +/g, ' ') : name === 'UPPER' ? value.toUpperCase() : value.toLowerCase(); }
        if (name === 'LEFT' || name === 'RIGHT') { arity(1, 2); const value = text(arg(0)); const length = Math.trunc(number(arg(1, 1))); if (length < 0) return fail(); return name === 'LEFT' ? value.slice(0, length) : length === 0 ? '' : value.slice(-length); }
        if (name === 'MID') { arity(3); const value = text(arg(0)); const start = Math.trunc(number(arg(1))); const length = Math.trunc(number(arg(2))); if (start < 1 || length < 0) return fail(); return value.slice(start - 1, start - 1 + length); }
        if (name === 'CONCAT' || name === 'CONCATENATE') { arity(1, 255); return args.flatMap(nodeArg => values(run(nodeArg))).map(value => text(value)).join(''); }
        if (name === 'TODAY') { arity(0); const date = new Date(); return serialDate(date.getFullYear(), date.getMonth() + 1, date.getDate()); }
        if (name === 'DATE') { arity(3); let year = Math.trunc(number(arg(0))); const month = Math.trunc(number(arg(1))); const day = Math.trunc(number(arg(2))); if (year < 0 || year >= 10000) return fail('#NUM!'); if (year < 1900) year += 1900; const result = serialDate(year, month, 1) + day - 1; return result >= 0 && result <= 2958465 ? result : fail('#NUM!'); }
        return fail('#NAME?');
      };
      const result = checked(evaluate(parse(raw.slice(1))));
      const final = typeof result === 'number' ? finite(Number(result.toPrecision(15))) : result === null ? 0 : result;
      cache.set(cacheKey, final); return final;
    } catch (error) { const result = error instanceof FormulaError ? error : new FormulaError('#VALUE!'); if (cacheKey) cache.set(cacheKey, result); return result; }
  };
  return key => { const value = calculate(key, currentSheetId, new Set()); return value instanceof FormulaError ? value.message : value === null ? '' : typeof value === 'boolean' ? value ? 'TRUE' : 'FALSE' : value; };
}

/** Shift only real cell references: preserve strings, sheet names, and absolute row/column markers. */
export function shiftFormulaReferences(formula: string, rowDelta: number, colDelta: number): string {
  if (!formula.startsWith('=')) return formula;
  return formula.replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'|\$?[A-Za-z]+\$?[1-9]\d*/g, (match: string, offset: number) => {
    if (match.startsWith('"') || match.startsWith("'")) return match;
    const before = formula[offset - 1] || ''; const after = formula.slice(offset + match.length);
    if (/[\p{L}\p{N}_.]/u.test(before) || /^[\p{L}\p{N}_.]/u.test(after) || /^\s*[!(]/.test(after)) return match;
    const ref = match.match(/^(\$?)([A-Za-z]+)(\$?)(\d+)$/)!;
    try {
      const point = pointFor(match); const row = point.row + (ref[3] ? 0 : rowDelta); const col = point.col + (ref[1] ? 0 : colDelta);
      if (row < 0 || col < 0 || row >= MAX_ROW || col >= MAX_COL) return '#REF!';
      return `${ref[1]}${columnName(col)}${ref[3]}${row + 1}`;
    } catch { return match; }
  });
}
