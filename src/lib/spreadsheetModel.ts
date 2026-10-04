import type { CellStyle, SheetContent, SheetTab } from '../types';

export interface WorkbookState { sheets: SheetTab[]; activeSheetId: string; }
const record = (value: unknown): Record<string, unknown> => value && typeof value === 'object' && !Array.isArray(value) ? value as Record<string, unknown> : {};

/** Rename complete sheet qualifiers while leaving string literals and other sheet names alone. */
export function renameSheetReferences(formula: string, before: string, after: string): string {
  if (!formula.startsWith('=')) return formula;
  return formula.replace(/"(?:[^"]|"")*"|('(?:[^']|'')*'|[\p{L}_][\p{L}\p{N}_.]*)\s*!/gu, (match: string, qualifier?: string) => {
    if (!qualifier) return match;
    const name = qualifier.startsWith("'") ? qualifier.slice(1, -1).replaceAll("''", "'") : qualifier;
    return name.toLocaleLowerCase() === before.toLocaleLowerCase() ? `'${after.replaceAll("'", "''")}'!` : match;
  });
}

/** Keep the complete workbook while accepting files saved by the original one-sheet editor. */
export function normalizeWorkbook(content: unknown): WorkbookState {
  const source = record(content);
  const entries = Array.isArray(source.sheets) && source.sheets.length ? source.sheets : [{ id: 'sheet-1', name: source.name, cells: source.cells }];
  const usedNames = new Set<string>();
  const usedIds = new Set<string>();
  const originalNames = new Set<string>();
  const renames: Array<[string, string]> = [];
  const sheets: SheetTab[] = entries.map((entry, index) => {
    const sheet = record(entry);
    const baseName = (String(sheet.name || `Sheet ${index + 1}`).replace(/[\\/?*\[\]:\u0000-\u001f]/g, '').trim().replace(/^'+|'+$/g, '') || `Sheet ${index + 1}`).slice(0, 31);
    let name = baseName;
    for (let suffix = 2; usedNames.has(name.toLocaleLowerCase()); suffix++) {
      const tail = ` (${suffix})`; name = baseName.slice(0, 31 - tail.length) + tail;
    }
    usedNames.add(name.toLocaleLowerCase());
    if (typeof sheet.name === 'string') {
      const originalKey = sheet.name.toLocaleLowerCase();
      if (!originalNames.has(originalKey) && sheet.name !== name) renames.push([sheet.name, name]);
      originalNames.add(originalKey);
    }
    const baseId = typeof sheet.id === 'string' && sheet.id ? sheet.id : `sheet-${index + 1}`;
    let id = baseId;
    for (let suffix = 2; usedIds.has(id); suffix++) id = `${baseId}-${suffix}`;
    usedIds.add(id);
    const cells = Object.fromEntries(Object.entries(record(sheet.cells)).map(([key, value]) => [key, value == null ? '' : String(value)]));
    return {
      ...sheet, id, name, cells,
      ...(sheet.styles ? { styles: Object.fromEntries(Object.entries(record(sheet.styles)).map(([key, style]) => [key, { ...record(style) } as CellStyle])) } : {}),
      ...(sheet.columnWidths ? { columnWidths: { ...record(sheet.columnWidths) } as Record<string, number> } : {}),
    } as SheetTab;
  });
  if (renames.length) for (const sheet of sheets) {
    sheet.cells = Object.fromEntries(Object.entries(sheet.cells).map(([key, value]) => {
      // One pass over original qualifiers avoids chained renames when sanitized names collide.
      if (value.startsWith('=')) value = value.replace(/"(?:[^"]|"")*"|('(?:[^']|'')*'|[\p{L}_][\p{L}\p{N}_.]*)\s*!/gu, (match: string, qualifier?: string) => {
        if (!qualifier) return match;
        const name = qualifier.startsWith("'") ? qualifier.slice(1, -1).replaceAll("''", "'") : qualifier;
        const rename = renames.find(([old]) => old.toLocaleLowerCase() === name.toLocaleLowerCase());
        return rename ? `'${rename[1].replaceAll("'", "''")}'!` : match;
      });
      return [key, value];
    }));
  }
  const activeSheetId = typeof source.activeSheetId === 'string' && sheets.some(sheet => sheet.id === source.activeSheetId) ? source.activeSheetId : sheets[0].id;
  return { sheets, activeSheetId };
}

/** Mirror the active tab for compatibility with older saved files and integrations. */
export function packWorkbook(sheets: SheetTab[], activeSheetId: string): SheetContent {
  const normalized = normalizeWorkbook({ sheets, activeSheetId });
  const active = normalized.sheets.find(sheet => sheet.id === normalized.activeSheetId)!;
  return { ...normalized, cells: { ...active.cells }, name: active.name };
}
