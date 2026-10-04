import type { SheetTab } from '../types';

type Scalar = string | number | boolean | null | FormulaError;
type Value = Scalar | { kind: 'reference'; value: Scalar } | { kind: 'range'; rows: Scalar[][] };
type Token = { kind: 'number' | 'string' | 'name' | 'quoted' | 'symbol' | 'error'; text: string };
type Node = { kind: 'literal'; value: Scalar } | { kind: 'reference'; sheet?: string; ref: string } | { kind: 'range'; start: Extract<Node, { kind: 'reference' }>; end: Extract<Node, { kind: 'reference' }> } | { kind: 'unary'; op: string; node: Node } | { kind: 'binary'; op: string; left: Node; right: Node } | { kind: 'call'; name: string; args: Node[] };
export interface FormulaContext { sheets: SheetTab[]; currentSheetId: string; }
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
  let pattern = '^';
  for (let i = 0; i < desired.length; i++) {
    const char = desired[i];
    if (char === '~' && i + 1 < desired.length) pattern += desired[++i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    else if (char === '*') pattern += '.*'; else if (char === '?') pattern += '.'; else pattern += char.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  const wildcard = new RegExp(`${pattern}$`, 'i');
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
        const { name, args } = node;
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
