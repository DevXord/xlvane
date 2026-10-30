// Surgical insertion and deletion of worksheet rows in existing XLSX files (docs/decisions.md
// D-B04). Built on the editor of edit.ts: only the parts that hold row numbers or references are
// rewritten, everything else keeps its exact content. References follow the rules of
// sml/ref-adjust.ts (docs/facts/ref-adjust.md R-01 … R-14); anything that cannot be rewritten
// safely makes the operation fail with an error instead of leaving a wrong reference behind.
import {
  PackageEditor,
  inputBytes,
  removeContentType,
  removeRelationships,
  requestFullCalculation,
  scan,
  setAttr,
  splice,
  type EditOptions,
  type EditResult,
  type Element,
} from './edit';
import { XlvError, type XlvWarning } from './errors';
import { resolveLimits, type ReadLimits } from './limits';
import type { Sheet, Workbook } from './model';
import { openPackage, type OpcPackage } from './opc/package';
import { readWorkbook } from './read';
import { formatCellRef, formatRange, MAX_ROWS, parseCellRef, parseRange } from './sml/cell-ref';
import { shiftFormula } from './sml/formula-shift';
import { classifyRelationship } from './sml/namespaces';
import { adjustFormula, adjustRange, type ReferenceChange } from './sml/ref-adjust';
import { readOptions } from './util/options';
import { WarningCollector } from './util/warnings';
import { escapeAttr, escapeText } from './xml/writer';
import { openZip } from './zip/reader';

/**
 * Options for {@link insertRows}.
 *
 * @example
 * ```ts
 * import type { InsertRowsOptions } from 'xlvane';
 *
 * // Two empty rows before row 5 of "Sales", formatted like row 4.
 * const options: InsertRowsOptions = { sheet: 'Sales', at: 5, count: 2, style: 'above' };
 * ```
 */
export interface InsertRowsOptions extends EditOptions {
  /** Sheet name (exact). */
  readonly sheet: string;
  /** The new rows are inserted before this row (1-based). */
  readonly at: number;
  /** How many rows. Default 1. */
  readonly count?: number;
  /**
   * Formatting of the new rows: `'above'` copies the row format, height and cell formats of the row
   * above, `'below'` those of the row that is pushed down, `'none'` (default) leaves them empty.
   */
  readonly style?: 'above' | 'below' | 'none';
}

/**
 * Options for {@link deleteRows}.
 *
 * @example
 * ```ts
 * import type { DeleteRowsOptions } from 'xlvane';
 *
 * // Rows 5 and 6 of "Sales".
 * const options: DeleteRowsOptions = { sheet: 'Sales', at: 5, count: 2 };
 * ```
 */
export interface DeleteRowsOptions extends EditOptions {
  /** Sheet name (exact). */
  readonly sheet: string;
  /** The first row to delete (1-based). */
  readonly at: number;
  /** How many rows. Default 1. */
  readonly count?: number;
}

/** Relationship types of a worksheet whose parts do not depend on row positions or are handled here. */
const SAFE_SHEET_RELATIONSHIPS = new Set(['hyperlink', 'comments', 'vmlDrawing', 'printerSettings', 'image', 'customProperty', 'drawing', 'table']);

/** Elements of a worksheet whose content refers to cells and that xlvane does not rewrite. */
const REFUSED_ELEMENTS: Readonly<Record<string, string>> = {
  oleObjects: 'embedded objects',
  controls: 'form controls',
  scenarios: 'scenarios',
  cellWatches: 'cell watches',
};

const WORKAROUND = 'Insert or delete the rows in a spreadsheet application, or remove that content first.';

interface RowChange {
  readonly sheet: string;
  readonly at: number;
  /** Positive to insert, negative to delete. */
  readonly count: number;
  readonly sheets: readonly string[];
}

/** A character that cannot be part of a sheet name, used as the "sheet" of defined names. */
const NO_SHEET = '\u0003';

function refChange(c: RowChange, formulaSheet: string): ReferenceChange {
  return { sheet: c.sheet, formulaSheet, axis: 'row', at: c.at, count: c.count, sheets: c.sheets };
}

/** New number of a row, or `null` when the row is deleted or pushed beyond the last row. */
function moveRow(n: number, c: RowChange): number | null {
  if (c.count > 0) {
    if (n < c.at) return n;
    return n + c.count > MAX_ROWS ? null : n + c.count;
  }
  const end = c.at - c.count - 1;
  if (n < c.at) return n;
  return n > end ? n + c.count : null;
}

function decodeEntities(s: string): string {
  return s.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|amp|lt|gt|quot|apos);/g, (_m, e: string) => {
    if (e === 'amp') return '&';
    if (e === 'lt') return '<';
    if (e === 'gt') return '>';
    if (e === 'quot') return '"';
    if (e === 'apos') return "'";
    return String.fromCodePoint(e.startsWith('#x') ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
  });
}

function removeAttr(tag: string, name: string): string {
  return tag.replace(new RegExp(`\\s${name}\\s*=\\s*("[^"]*"|'[^']*')`), '');
}

function content(text: string, e: Element): string {
  return e.start.empty ? '' : text.slice(e.start.end, e.end.start);
}

function notSupported(message: string, location: { part?: string; sheet?: string; cell?: string }, details?: Record<string, string | number>): XlvError {
  return new XlvError('XLV_E_NOT_SUPPORTED', message, { location, hint: WORKAROUND, ...(details ? { details } : {}) });
}

/**
 * Adjusts the formula text of an element (XML-escaped, as in the file). Returns the new escaped
 * text, the same text when nothing changes, or `null` when the formula cannot be adjusted although
 * it may refer to the changed sheet.
 */
function adjustText(raw: string, c: RowChange, formulaSheet: string): string | null {
  if (raw.includes('<')) return null;
  const text = decodeEntities(raw);
  const out = adjustFormula(text, refChange(c, formulaSheet));
  if (out === null) {
    // A formula on another sheet without any sheet-qualified reference cannot refer to the changed sheet.
    return formulaSheet.toLowerCase() !== c.sheet.toLowerCase() && !text.includes('!') ? raw : null;
  }
  return out === text ? raw : escapeText(out);
}

/**
 * Applies non-overlapping replacements in one pass (linear in the text size, unlike edit.ts
 * `splice`, which matters for sheets with many moved cells). An edit inside an earlier replaced
 * span, such as a removed row, is dropped with it; an insertion at the start of a removed element
 * is kept.
 */
function applyEdits(text: string, edits: readonly [number, number, string][]): string {
  const sorted = [...edits].sort((a, b) => a[0] - b[0] || Number(b[0] === b[1]) - Number(a[0] === a[1]) || b[1] - a[1]);
  const out: string[] = [];
  let pos = 0;
  for (const [start, end, replacement] of sorted) {
    // An edit that starts inside an already replaced span (a removed element) is dropped with it.
    if (start < pos) continue;
    out.push(text.slice(pos, start), replacement);
    pos = end;
  }
  out.push(text.slice(pos));
  return out.join('');
}

interface SheetEdit {
  readonly text: string;
  readonly hasFormulas: boolean;
}

interface CellInfo {
  readonly col: number;
  readonly hasR: boolean;
  readonly element: Element;
  f: Element | null;
}

interface RowInfo {
  readonly number: number;
  readonly hasR: boolean;
  readonly element: Element;
  readonly cells: CellInfo[];
}

const KEEP_DEPTH3 = new Set(['row', 'mergeCell', 'hyperlink', 'cfRule', 'dataValidation', 'brk', 'ignoredError', 'protectedRange']);

/** Rewrites one worksheet part: references everywhere, and row positions on the changed sheet. */
function rewriteSheet(text: string, part: string, sheet: Sheet, c: RowChange, target: boolean, style: 'above' | 'below' | 'none', limits: ReadLimits): SheetEdit {
  const { elements } = scan(
    text,
    part,
    limits,
    (local, depth, parent) =>
      depth === 2 ||
      (depth === 3 && KEEP_DEPTH3.has(local)) ||
      (depth === 4 && ((local === 'c' && parent === 'row') || (local === 'formula' && parent === 'cfRule') || ((local === 'formula1' || local === 'formula2') && parent === 'dataValidation'))) ||
      (depth === 5 && local === 'f' && parent === 'c'),
  );
  const edits: [number, number, string][] = [];
  const tagOf = (e: Element): string => text.slice(e.start.start, e.start.end);
  const replaceTag = (e: Element, tag: string): void => {
    edits.push([e.start.start, e.start.end, tag]);
  };
  const removeElement = (e: Element): void => {
    edits.push([e.start.start, e.end.end, '']);
  };
  const where = (cell?: string): { part: string; sheet: string; cell?: string } => (cell === undefined ? { part, sheet: sheet.name } : { part, sheet: sheet.name, cell });

  if (target) {
    for (const e of elements) {
      if (e.depth !== 2) continue;
      const what = REFUSED_ELEMENTS[e.local];
      if (what !== undefined) throw notSupported(`Sheet "${sheet.name}" has ${what}, which xlvane cannot move with its rows yet.`, where(), { element: e.local });
      if (e.local === 'extLst' && /sqref|<xm:f|sparkline|slicer/i.test(text.slice(e.start.start, e.end.end))) {
        throw notSupported(`Sheet "${sheet.name}" has extended content (such as data bars, sparklines or slicers) that refers to cells, which xlvane cannot move with its rows yet.`, where(), { element: 'extLst' });
      }
    }
  } else {
    const ext = elements.find((e) => e.depth === 2 && e.local === 'extLst');
    if (ext && /<xm:f|sparkline/i.test(text.slice(ext.start.start, ext.end.end)) && text.slice(ext.start.start, ext.end.end).toLowerCase().includes(c.sheet.toLowerCase())) {
      throw notSupported(`Sheet "${sheet.name}" has extended content (such as sparklines) that may refer to sheet "${c.sheet}", which xlvane cannot rewrite yet.`, where(), { element: 'extLst' });
    }
  }

  // Rows and cells, numbered as the reader numbers them (D-105).
  const rows: RowInfo[] = [];
  let current: RowInfo | null = null;
  let currentRow = 0;
  let lastCol = 0;
  let lastCell: CellInfo | null = null;
  const formulas: { cell: CellInfo; row: RowInfo }[] = [];
  const invalidRef = (ref: string): never => {
    throw new XlvError('XLV_E_INVALID_CELL_REF', `Sheet "${sheet.name}" contains the invalid reference "${ref.slice(0, 20)}", so its rows cannot be moved safely.`, {
      location: where(),
      hint: 'Open and re-save the file in a spreadsheet application, which drops the damaged cells, then try again.',
      details: { ref: ref.slice(0, 40) },
    });
  };
  let hasFormulas = false;
  for (const e of elements) {
    if (e.depth === 3 && e.local === 'row') {
      const r = e.attrs.get('r');
      const n = r === undefined ? currentRow + 1 : /^[0-9]{1,7}$/.test(r.trim()) ? Number(r.trim()) : Number.NaN;
      if (!Number.isInteger(n) || n < 1 || n > MAX_ROWS) invalidRef(r ?? String(n));
      currentRow = n;
      current = { number: n, hasR: r !== undefined, element: e, cells: [] };
      rows.push(current);
      lastCol = 0;
    } else if (e.depth === 4 && e.local === 'c' && current) {
      const r = e.attrs.get('r');
      const address = r === undefined || r === '' ? null : parseCellRef(r);
      if (r !== undefined && r !== '' && address === null) invalidRef(r);
      const col = address?.col ?? lastCol + 1;
      lastCol = col;
      lastCell = { col, hasR: r !== undefined && r !== '', element: e, f: null };
      current.cells.push(lastCell);
    } else if (e.depth === 5 && e.local === 'f' && lastCell && current) {
      lastCell.f = e;
      hasFormulas = true;
      formulas.push({ cell: lastCell, row: current });
    } else if (e.depth === 4 && (e.local === 'formula' || e.local === 'formula1' || e.local === 'formula2')) {
      const raw = content(text, e);
      const out = adjustText(raw, c, sheet.name);
      if (out === null) {
        const owner = e.local === 'formula' ? 'conditional format' : 'data validation';
        throw notSupported(`A ${owner} formula of sheet "${sheet.name}" cannot be adjusted safely: "${decodeEntities(raw).slice(0, 60)}".`, where(), { formula: decodeEntities(raw).slice(0, 80) });
      }
      if (out !== raw) edits.push([e.start.end, e.end.start, out]);
    } else if (e.depth === 3 && e.local === 'hyperlink') {
      const location = e.attrs.get('location');
      let tag = tagOf(e);
      if (location !== undefined) {
        const out = adjustFormula(location, refChange(c, sheet.name));
        if (out !== null && out !== location) tag = setAttr(tag, 'location', out);
      }
      if (target) {
        const ref = adjustRange(e.attrs.get('ref') ?? '', refChange(c, sheet.name));
        if (ref === '') {
          removeElement(e);
          continue;
        }
        if (ref !== null && ref !== e.attrs.get('ref')) tag = setAttr(tag, 'ref', ref);
      }
      if (tag !== tagOf(e)) replaceTag(e, tag);
    }
  }

  // Cell formulas, with shared formulas expanded where they cannot stay shared.
  const masters = new Map<string, { row: number; col: number; text: string }>();
  for (const { cell, row } of formulas) {
    const f = cell.f as Element;
    if (f.attrs.get('t') === 'shared' && f.attrs.has('ref') && f.attrs.has('si')) masters.set(f.attrs.get('si') ?? '', { row: row.number, col: cell.col, text: content(text, f) });
  }
  const expand = new Set<string>();
  for (const [si, m] of masters) {
    const out = adjustText(m.text, c, sheet.name);
    if (target || out === null || out !== m.text) expand.add(si);
  }
  const deletedRow = (n: number): boolean => c.count < 0 && moveRow(n, c) === null;
  for (const { cell, row } of formulas) {
    const f = cell.f as Element;
    if (target && deletedRow(row.number)) continue;
    const ref = formatCellRef(row.number, cell.col);
    const kind = f.attrs.get('t');
    const failed = (formula: string): never => {
      throw notSupported(`The formula in ${ref} of sheet "${sheet.name}" cannot be adjusted safely: "${formula.slice(0, 60)}".`, where(ref), { formula: formula.slice(0, 80) });
    };
    if (kind === 'shared') {
      const si = f.attrs.get('si') ?? '';
      const master = masters.get(si);
      if (master === undefined || !expand.has(si)) continue;
      const source = decodeEntities(master.text);
      const own = shiftFormula(source, row.number - master.row, cell.col - master.col);
      if (own === null) failed(source);
      const adjusted = adjustFormula(own as string, refChange(c, sheet.name));
      if (adjusted === null) failed(own as string);
      let tag = removeAttr(removeAttr(removeAttr(tagOf(f), 't'), 'ref'), 'si');
      if (tag.endsWith('/>')) tag = `${tag.slice(0, -2).trimEnd()}>`;
      edits.push([f.start.start, f.end.end, `${tag}${escapeText(adjusted as string)}</${f.start.qname}>`]);
      continue;
    }
    if (kind === 'dataTable') {
      if (target) throw notSupported(`Cell ${ref} of sheet "${sheet.name}" is part of a data table, which xlvane cannot move.`, where(ref));
      continue;
    }
    const raw = content(text, f);
    const out = adjustText(raw, c, sheet.name);
    if (out === null) failed(decodeEntities(raw));
    let tag = tagOf(f);
    if (kind === 'array' && target) {
      const range = parseRange(f.attrs.get('ref') ?? '');
      if (range !== null) {
        const end = c.at - c.count - 1;
        const split = c.count > 0 ? c.at > range.start.row && c.at <= range.end.row : !(range.start.row >= c.at && range.end.row <= end) && range.end.row >= c.at && range.start.row <= end;
        if (split) throw notSupported(`The array formula in ${f.attrs.get('ref') ?? ref} of sheet "${sheet.name}" would be split by the change.`, where(ref));
        const moved = adjustRange(f.attrs.get('ref') ?? '', refChange(c, sheet.name));
        if (moved !== null && moved !== '' && moved !== f.attrs.get('ref')) tag = setAttr(tag, 'ref', moved);
      }
    }
    if (tag !== tagOf(f)) replaceTag(f, tag);
    if (out !== raw) edits.push([f.start.end, f.end.start, out as string]);
  }

  if (!target) return { text: edits.length === 0 ? text : applyEdits(text, edits), hasFormulas };

  // Rows and cells of the changed sheet.
  const force = rows.some((r) => !r.hasR || r.cells.some((x) => !x.hasR));
  for (const row of rows) {
    const n = moveRow(row.number, c);
    if (n === null) {
      if (c.count > 0 && row.cells.some((x) => !x.element.start.empty)) {
        throw new XlvError('XLV_E_LIMIT_EXCEEDED', `Inserting ${String(c.count)} rows would push row ${String(row.number)} of sheet "${sheet.name}", which holds data, past the last row of the sheet.`, {
          location: where(),
          hint: `A sheet has ${String(MAX_ROWS)} rows. Delete the rows at the end of the sheet first, or insert fewer rows.`,
          details: { limit: 'rows', value: MAX_ROWS },
        });
      }
      removeElement(row.element);
      continue;
    }
    if (n === row.number && !force) continue;
    replaceTag(row.element, setAttr(tagOf(row.element), 'r', String(n)));
    for (const cell of row.cells) replaceTag(cell.element, setAttr(tagOf(cell.element), 'r', formatCellRef(n, cell.col)));
  }
  if (c.count > 0 && style !== 'none') {
    const source = rows.find((r) => r.number === (style === 'above' ? c.at - 1 : c.at));
    if (source) {
      const prefix = source.element.start.qname.slice(0, source.element.start.qname.length - 3);
      const keep = ['s', 'customFormat', 'ht', 'customHeight'].filter((a) => source.element.attrs.has(a));
      let created = '';
      for (let i = 0; i < c.count; i++) {
        const n = c.at + i;
        const attrs = keep.map((a) => ` ${a}="${escapeAttr(source.element.attrs.get(a) ?? '')}"`).join('');
        const cells = source.cells.filter((x) => (x.element.attrs.get('s') ?? '0') !== '0').map((x) => `<${prefix}c r="${formatCellRef(n, x.col)}" s="${escapeAttr(x.element.attrs.get('s') ?? '0')}"/>`).join('');
        created += cells === '' ? `<${prefix}row r="${String(n)}"${attrs}/>` : `<${prefix}row r="${String(n)}"${attrs}>${cells}</${prefix}row>`;
      }
      const next = rows.find((r) => r.number >= c.at);
      if (next) edits.push([next.element.start.start, next.element.start.start, created]);
      else {
        const sheetData = elements.find((e) => e.depth === 2 && e.local === 'sheetData');
        if (sheetData && !sheetData.start.empty) edits.push([sheetData.end.start, sheetData.end.start, created]);
        else if (sheetData) edits.push([sheetData.start.start, sheetData.start.end, `${tagOf(sheetData).slice(0, -2).trimEnd()}>${created}</${sheetData.start.qname}>`]);
      }
    }
  }

  // Elements of the changed sheet that hold ranges.
  const containers = new Map<Element, { kept: number; total: number }>();
  let parent2: Element | null = null;
  for (const e of elements) {
    if (e.depth === 2) {
      parent2 = e;
      if (e.local === 'dimension' || e.local === 'autoFilter') {
        const ref = e.attrs.get('ref');
        if (ref === undefined) continue;
        const out = adjustRange(ref, refChange(c, sheet.name));
        if (out === null || out === ref) continue;
        if (e.local === 'dimension') replaceTag(e, setAttr(tagOf(e), 'ref', out === '' ? 'A1' : (out.split(' ')[0] ?? 'A1')));
        else if (out === '') removeElement(e);
        else replaceTag(e, setAttr(tagOf(e), 'ref', out));
      } else if (e.local === 'conditionalFormatting') {
        const sqref = e.attrs.get('sqref') ?? '';
        const out = adjustRange(sqref, refChange(c, sheet.name));
        if (out === null) throw notSupported(`A conditional format of sheet "${sheet.name}" has a range xlvane cannot adjust: "${sqref.slice(0, 60)}".`, where());
        if (out === '') removeElement(e);
        else if (out !== sqref) replaceTag(e, setAttr(tagOf(e), 'sqref', out));
      }
      continue;
    }
    if (e.depth !== 3 || parent2 === null) continue;
    const container = parent2;
    const count = containers.get(container) ?? { kept: 0, total: 0 };
    containers.set(container, count);
    if (e.local === 'mergeCell') {
      count.total++;
      const ref = e.attrs.get('ref') ?? '';
      const range = parseRange(ref);
      // R-14: a merge whose top-left cell is deleted, or that shrinks to one cell, is dropped.
      const out = range === null ? null : c.count < 0 && deletedRow(range.start.row) ? '' : adjustRange(ref, refChange(c, sheet.name));
      if (out === null) {
        count.kept++;
        continue;
      }
      const moved = out === '' ? null : parseRange(out);
      if (moved === null || (moved.start.row === moved.end.row && moved.start.col === moved.end.col)) {
        removeElement(e);
        continue;
      }
      count.kept++;
      if (out !== ref) replaceTag(e, setAttr(tagOf(e), 'ref', out));
    } else if (e.local === 'dataValidation' || e.local === 'ignoredError' || e.local === 'protectedRange') {
      count.total++;
      const sqref = e.attrs.get('sqref') ?? '';
      const out = adjustRange(sqref, refChange(c, sheet.name));
      if (out === null) throw notSupported(`Sheet "${sheet.name}" has a ${e.local} range xlvane cannot adjust: "${sqref.slice(0, 60)}".`, where());
      if (out === '') {
        removeElement(e);
        continue;
      }
      count.kept++;
      if (out !== sqref) replaceTag(e, setAttr(tagOf(e), 'sqref', out));
    } else if (e.local === 'hyperlink') {
      // The element itself was rewritten or removed above; only its container is counted here.
      count.total++;
      if (adjustRange(e.attrs.get('ref') ?? '', refChange(c, sheet.name)) !== '') count.kept++;
    } else if (e.local === 'brk' && container.local === 'rowBreaks') {
      count.total++;
      const id = Number(e.attrs.get('id') ?? 'NaN');
      if (!Number.isInteger(id)) {
        count.kept++;
        continue;
      }
      const n = moveRow(id, c);
      if (n === null) {
        removeElement(e);
        continue;
      }
      count.kept++;
      if (n !== id) replaceTag(e, setAttr(tagOf(e), 'id', String(n)));
    }
  }
  for (const [container, n] of containers) {
    if (n.total === 0 || n.kept === n.total) continue;
    // Elements removed: drop an emptied container, or update its count.
    const counted = ['mergeCells', 'dataValidations', 'rowBreaks'].includes(container.local);
    if (n.kept === 0) {
      removeElement(container);
    } else if (counted) {
      let tag = tagOf(container);
      if (container.attrs.has('count')) tag = setAttr(tag, 'count', String(n.kept));
      if (container.local === 'rowBreaks' && container.attrs.has('manualBreakCount')) tag = setAttr(tag, 'manualBreakCount', String(n.kept));
      if (tag !== tagOf(container)) replaceTag(container, tag);
    }
  }

  return { text: applyEdits(text, edits), hasFormulas };
}

/** Rewrites the notes of the changed sheet: the Comments part and the note shapes of its VML part. */
function rewriteNotes(edit: PackageEditor, pkg: OpcPackage, sheetPart: string, sheetText: string, sheet: Sheet, c: RowChange): void {
  const rels = pkg.relationships(sheetPart);
  const comments = rels.find((r) => !r.external && classifyRelationship(r.type)?.name === 'comments');
  if (comments !== undefined) {
    const text = edit.text(comments.target);
    if (text !== undefined) {
      const { elements } = scan(text, comments.target, edit.limits, (local, depth) => depth === 3 && local === 'comment');
      const edits: [number, number, string][] = [];
      for (const e of elements) {
        const ref = e.attrs.get('ref') ?? '';
        const out = adjustRange(ref, refChange(c, sheet.name));
        if (out === null || out === ref) continue;
        if (out === '') edits.push([e.start.start, e.end.end, '']);
        else edits.push([e.start.start, e.start.end, setAttr(text.slice(e.start.start, e.start.end), 'ref', out)]);
      }
      if (edits.length > 0) edit.write(comments.target, splice(text, edits));
    }
  }
  // The VML part of the notes is the one named by legacyDrawing (W2-48).
  const id = /<(?:[A-Za-z_][\w.-]*:)?legacyDrawing\b[^>]*?\br:id\s*=\s*["']([^"']*)["']/.exec(sheetText)?.[1];
  const vml = id === undefined ? undefined : rels.find((r) => r.id === id && !r.external);
  if (vml === undefined) return;
  const text = edit.text(vml.target);
  if (text === undefined) return;
  let changed = false;
  const out = text.replace(/<v:shape\b[\s\S]*?<\/v:shape>/g, (shape) => {
    if (!/ObjectType\s*=\s*["']Note["']/.test(shape)) return shape;
    const row = /<x:Row>\s*(\d+)\s*<\/x:Row>/.exec(shape);
    if (row === null) return shape;
    const n = moveRow(Number(row[1]) + 1, c);
    changed = true;
    if (n === null) return '';
    const moveAnchor = (v: number): number => {
      const one = v + 1;
      if (c.count > 0) return one >= c.at ? v + c.count : v;
      const end = c.at - c.count - 1;
      return one < c.at ? v : one > end ? v + c.count : c.at - 1;
    };
    return shape
      .replace(/<x:Row>\s*\d+\s*<\/x:Row>/, `<x:Row>${String(n - 1)}</x:Row>`)
      .replace(/<x:Anchor>([^<]*)<\/x:Anchor>/, (all, list: string) => {
        const v = list.split(',').map((s) => s.trim());
        if (v.length !== 8 || !v.every((s) => /^\d+$/.test(s))) return all;
        const nums = v.map(Number);
        nums[2] = moveAnchor(nums[2] ?? 0);
        nums[6] = Math.max(nums[2] ?? 0, moveAnchor(nums[6] ?? 0));
        return `<x:Anchor>${nums.join(', ')}</x:Anchor>`;
      });
  });
  if (changed) edit.write(vml.target, out);
}

/**
 * Moves the pictures of the changed sheet's drawing with their rows (TD-38): the `row` of each
 * `from` and `to` marker of two-cell and one-cell anchors; absolute anchors do not depend on
 * rows. Drawings with charts, shapes, groups or other objects are refused, and so is a deletion
 * that would cut a picture (a deleted row inside the rows it covers or under its anchor cell).
 */
function rewriteDrawing(edit: PackageEditor, pkg: OpcPackage, sheetPart: string, sheetText: string, sheet: Sheet, c: RowChange): void {
  const id = /<(?:[A-Za-z_][\w.-]*:)?drawing\b[^>]*?\br:id\s*=\s*["']([^"']*)["']/.exec(sheetText)?.[1];
  const rels = pkg.relationships(sheetPart);
  const rel = id === undefined ? undefined : rels.find((r) => r.id === id && !r.external && classifyRelationship(r.type)?.name === 'drawing');
  if (rel === undefined) return;
  const text = edit.text(rel.target);
  if (text === undefined) return;
  const where = { part: rel.target, sheet: sheet.name };
  for (const r of pkg.relationships(rel.target)) {
    const name = classifyRelationship(r.type)?.name;
    if (!r.external && name !== 'image' && name !== 'hyperlink') throw notSupported(`The drawing of sheet "${sheet.name}" has a related part of type "${name ?? r.type}" (${r.target}), which xlvane cannot move with its rows yet.`, where, { type: r.type });
  }
  const anchors = /<([A-Za-z_][\w.-]*:)?(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[\s\S]*?<\/\1\2>/g;
  const rest = text.replace(anchors, (anchor, p: string | undefined) => (new RegExp(`<${p ?? ''}pic[\\s>]`).test(anchor) ? '' : anchor));
  if (/graphicFrame|<[^>]*:sp[\s>]|grpSp|cxnSp|contentPart|AlternateContent|twoCellAnchor|oneCellAnchor|absoluteAnchor/.test(rest)) {
    throw notSupported(`The drawing of sheet "${sheet.name}" has charts, shapes or other objects besides pictures, which xlvane cannot move with its rows yet.`, where);
  }
  const end = c.at - c.count - 1;
  // A marker at the start of row m + 1 (0-based m): its new row, as a range end moves (R-03, R-08).
  const moveMarker = (m: number): number => {
    const one = m + 1;
    if (c.count > 0) return one >= c.at ? m + c.count : m;
    return one < c.at ? m : one > end ? m + c.count : c.at - 1;
  };
  let changed = false;
  const out = text.replace(anchors, (anchor, p: string | undefined, kind: string) => {
    if (kind === 'absoluteAnchor') return anchor;
    const pre = p ?? '';
    const marker = (local: 'from' | 'to'): { row: number; rowOff: number } | null => {
      const m = new RegExp(`<${pre}${local}>[\\s\\S]*?<${pre}row>\\s*(\\d+)\\s*</${pre}row>\\s*<${pre}rowOff>\\s*([^<]*?)\\s*</${pre}rowOff>`).exec(anchor);
      return m === null ? null : { row: Number(m[1]), rowOff: Number(m[2]) || 0 };
    };
    const from = marker('from');
    const to = kind === 'twoCellAnchor' ? marker('to') : null;
    if (from === null || (kind === 'twoCellAnchor' && to === null)) throw notSupported(`A picture of sheet "${sheet.name}" has a position xlvane cannot read, so its rows cannot be moved.`, where);
    const first = from.row + 1;
    const last = to === null ? first : to.rowOff === 0 && to.row > from.row ? to.row : to.row + 1;
    if (c.count < 0 && first <= end && last >= c.at) throw notSupported(`Deleting rows ${String(c.at)} to ${String(end)} of sheet "${sheet.name}" would cut a picture (rows ${String(first)} to ${String(last)}), which xlvane does not do.`, where, { cell: formatCellRef(first, 1) });
    const rows = [moveMarker(from.row), to === null ? null : moveMarker(to.row)];
    if (rows.some((r) => r !== null && r > MAX_ROWS - 1)) throw notSupported(`Inserting rows would push a picture of sheet "${sheet.name}" beyond the last row.`, where);
    if (rows[0] === from.row && (to === null || rows[1] === to.row)) return anchor;
    changed = true;
    let k = 0;
    return anchor.replace(new RegExp(`(<${pre}row>)\\s*\\d+\\s*(</${pre}row>)`, 'g'), (all, open: string, close: string) => {
      const value = rows[k++];
      return value === null || value === undefined ? all : `${open}${String(value)}${close}`;
    });
  });
  if (changed) edit.write(rel.target, out);
}

/**
 * Moves, grows or shrinks the tables of a sheet with its rows (TD-11): the table `ref` and its
 * `autoFilter` follow as ranges do (adjustRange), formulas of calculated columns and totals are
 * adjusted like cell formulas, and structured references are left as they are. A deletion that
 * would remove a table's header row, totals row or all of its data rows is refused. For tables
 * of other sheets (`target` false) only the formulas are adjusted.
 */
function rewriteTables(edit: PackageEditor, pkg: OpcPackage, sheetPart: string, sheet: Sheet, c: RowChange, target: boolean): void {
  for (const rel of pkg.relationships(sheetPart)) {
    if (rel.external || classifyRelationship(rel.type)?.name !== 'table') continue;
    const text = edit.text(rel.target);
    if (text === undefined) continue;
    const where = { part: rel.target, sheet: sheet.name };
    const { root, elements } = scan(text, rel.target, edit.limits, (local, depth, parent) => (depth === 2 && local === 'autoFilter') || (depth === 4 && parent === 'tableColumn' && (local === 'calculatedColumnFormula' || local === 'totalsRowFormula')));
    const name = root.attrs.get('displayName') ?? root.attrs.get('name') ?? '';
    const edits: [number, number, string][] = [];
    if (target) {
      const ref = root.attrs.get('ref') ?? '';
      const range = parseRange(ref.trim().toUpperCase());
      if (range === null) throw notSupported(`Table "${name}" of sheet "${sheet.name}" has a range xlvane cannot read ("${ref.slice(0, 20)}").`, where);
      const header = (root.attrs.get('headerRowCount') ?? '1').trim() !== '0' ? 1 : 0;
      const totals = /^[1-9]/.test((root.attrs.get('totalsRowCount') ?? '0').trim()) ? 1 : 0;
      if (c.count < 0) {
        const end = c.at - c.count - 1;
        const cuts = (row: number): boolean => row >= c.at && row <= end;
        const dataFirst = range.start.row + header;
        const dataLast = range.end.row - totals;
        if ((header === 1 && cuts(range.start.row)) || (totals === 1 && cuts(range.end.row)) || (c.at <= dataFirst && end >= dataLast && range.start.row <= end && range.end.row >= c.at)) {
          throw notSupported(`Deleting rows ${String(c.at)} to ${String(end)} of sheet "${sheet.name}" would remove the header row, the totals row or every data row of table "${name}" (${formatRange(range)}).`, where, { table: name });
        }
      }
      const newRef = adjustRange(ref, refChange(c, sheet.name));
      if (newRef === null || newRef === '') throw notSupported(`Table "${name}" of sheet "${sheet.name}" cannot be moved with its rows.`, where, { table: name });
      const moved = parseRange(newRef);
      if (moved !== null && moved.end.row - moved.start.row < range.end.row - range.start.row && c.count > 0) throw notSupported(`Inserting rows would push table "${name}" beyond the last row.`, where, { table: name });
      if (newRef !== ref) edits.push([root.start.start, root.start.end, setAttr(text.slice(root.start.start, root.start.end), 'ref', newRef)]);
    }
    for (const e of elements) {
      if (e.local === 'autoFilter') {
        if (!target) continue;
        const ref = e.attrs.get('ref');
        if (ref === undefined) continue;
        const out = adjustRange(ref, refChange(c, sheet.name));
        if (out === null || out === '') throw notSupported(`The filter of table "${name}" of sheet "${sheet.name}" cannot be moved with its rows.`, where, { table: name });
        if (out !== ref) edits.push([e.start.start, e.start.end, setAttr(text.slice(e.start.start, e.start.end), 'ref', out)]);
        continue;
      }
      const raw = content(text, e);
      const out = adjustText(raw, c, sheet.name);
      if (out === null) throw notSupported(`A formula of table "${name}" of sheet "${sheet.name}" cannot be adjusted safely: "${decodeEntities(raw).slice(0, 60)}".`, where, { table: name });
      if (out !== raw) edits.push([e.start.end, e.end.start, out]);
    }
    if (edits.length > 0) edit.write(rel.target, applyEdits(text, edits));
  }
}

/** Rewrites the defined names of the workbook part. */
function rewriteNames(edit: PackageEditor, workbookPart: string, c: RowChange): void {
  const text = edit.text(workbookPart);
  if (text === undefined) return;
  const { elements } = scan(text, workbookPart, edit.limits, (local, depth, parent) => depth === 3 && local === 'definedName' && parent === 'definedNames');
  const edits: [number, number, string][] = [];
  for (const e of elements) {
    const raw = content(text, e);
    const out = adjustText(raw, c, NO_SHEET);
    if (out === null) {
      throw notSupported(`The defined name "${e.attrs.get('name') ?? ''}" cannot be adjusted safely: "${decodeEntities(raw).slice(0, 60)}".`, { part: workbookPart }, { name: e.attrs.get('name') ?? '' });
    }
    if (out !== raw) edits.push([e.start.end, e.end.start, out]);
  }
  if (edits.length > 0) edit.write(workbookPart, splice(text, edits));
}

function checkOptions(options: unknown, insert: boolean): { sheet: string; at: number; count: number; style: 'above' | 'below' | 'none'; limits: ReadLimits } {
  const what = insert ? 'insertRows options' : 'deleteRows options';
  const o = readOptions(options, insert ? ['sheet', 'at', 'count', 'style', 'limits'] : ['sheet', 'at', 'count', 'limits'], what);
  const sheet = o.get('sheet');
  if (typeof sheet !== 'string') throw new XlvError('XLV_E_INVALID_OPTION', `${what} need the sheet name.`, { hint: "Pass { sheet: 'Sales', at: 5 }.", details: { option: 'sheet' } });
  const at = o.get('at');
  if (typeof at !== 'number' || !Number.isInteger(at) || at < 1 || at > MAX_ROWS) {
    throw new XlvError('XLV_E_INVALID_OPTION', `The option "at" must be a row number from 1 to ${String(MAX_ROWS)}.`, { hint: 'For example { at: 5 } for row 5.', details: { option: 'at' } });
  }
  const count = o.get('count') ?? 1;
  if (typeof count !== 'number' || !Number.isInteger(count) || count < 1 || at + count - 1 > MAX_ROWS) {
    throw new XlvError('XLV_E_INVALID_OPTION', `The option "count" must be a whole number of at least 1, and the rows must end at row ${String(MAX_ROWS)} or before.`, {
      hint: 'For example { count: 2 } for two rows.',
      details: { option: 'count' },
    });
  }
  const style = o.get('style') ?? 'none';
  if (style !== 'above' && style !== 'below' && style !== 'none') {
    throw new XlvError('XLV_E_INVALID_OPTION', `The option "style" must be "above", "below" or "none", not ${JSON.stringify(style)}.`, { hint: "For example { style: 'above' } to copy the formatting of the row above.", details: { option: 'style' } });
  }
  return { sheet, at, count, style, limits: resolveLimits(o.get('limits')) };
}

function changeRows(input: unknown, sheetName: string, at: number, count: number, style: 'above' | 'below' | 'none', limits: ReadLimits): EditResult {
  const bytes = inputBytes(input);
  const workbook: Workbook = readWorkbook(bytes, { limits });
  const sheet = workbook.getSheet(sheetName);
  if (sheet.kind !== 'worksheet' || sheet.part === null) {
    throw new XlvError('XLV_E_INVALID_INPUT', `Sheet "${sheet.name}" is a ${sheet.kind} without rows.`, { location: { sheet: sheet.name }, hint: 'Insert or delete rows of worksheets only.' });
  }
  const pkg = openPackage(openZip(bytes, limits), limits);
  const edit = new PackageEditor(bytes, limits, pkg);
  const warnings = new WarningCollector();
  const officeDocument = pkg.relationships(null).find((r) => classifyRelationship(r.type)?.name === 'officeDocument');
  const workbookPart = officeDocument?.target ?? '/xl/workbook.xml';
  const c: RowChange = { sheet: sheet.name, at, count, sheets: workbook.sheets.map((s) => s.name) };

  for (const rel of pkg.relationships(sheet.part)) {
    if (rel.external) continue;
    const name = classifyRelationship(rel.type)?.name ?? null;
    if (name === null || !SAFE_SHEET_RELATIONSHIPS.has(name)) {
      throw notSupported(`Sheet "${sheet.name}" has a related part of type "${name ?? rel.type}" (${rel.target}), which xlvane cannot move with its rows yet.`, { part: sheet.part, sheet: sheet.name }, { type: rel.type });
    }
  }
  // Charts and pivot caches anywhere may refer to the sheet by name.
  const lowerName = sheet.name.toLowerCase();
  for (const name of pkg.partNames) {
    if (!/\/charts\/[^/]+\.xml$|pivotCacheDefinition[^/]*\.xml$/i.test(name)) continue;
    const text = edit.text(name);
    if (text !== undefined && text.toLowerCase().includes(lowerName)) {
      throw notSupported(`The part ${name} (a chart or pivot table) may refer to sheet "${sheet.name}", and xlvane cannot rewrite it yet.`, { part: name, sheet: sheet.name });
    }
  }

  let hasFormulas = false;
  for (const s of workbook.sheets) {
    if (s.kind !== 'worksheet' || s.part === null) continue;
    const text = edit.text(s.part);
    if (text === undefined) continue;
    const target = s.part === sheet.part;
    const out = rewriteSheet(text, s.part, s, c, target, style, limits);
    hasFormulas ||= out.hasFormulas;
    if (out.text !== text) edit.write(s.part, out.text);
    rewriteTables(edit, pkg, s.part, s, c, target);
    if (target) {
      rewriteNotes(edit, pkg, s.part, text, s, c);
      rewriteDrawing(edit, pkg, s.part, text, s, c);
    }
  }
  rewriteNames(edit, workbookPart, c);

  if (hasFormulas) {
    requestFullCalculation(edit, workbookPart);
    warnings.add('XLV_W_FORMULAS_NOT_RECALCULATED', 'The workbook contains formulas. xlvane does not recalculate them: cached results may be out of date until the file is opened in a spreadsheet application, which is asked to recalculate on load.', { part: workbookPart });
    const rels = pkg.relationships(workbookPart);
    const chains = rels.filter((r) => !r.external && classifyRelationship(r.type)?.name === 'calcChain');
    for (const chain of chains) {
      edit.remove(chain.target);
      removeContentType(edit, chain.target);
    }
    if (chains.length > 0) removeRelationships(edit, workbookPart, new Set(chains.map((x) => x.id)));
  }
  return { bytes: edit.build(), warnings: warnings.list() as XlvWarning[] };
}

/**
 * Inserts empty rows into a sheet of an existing XLSX file and returns the new file. Rows below
 * move down, and every reference follows them as in a spreadsheet application: formulas on all
 * sheets (ranges that the new rows fall into grow), merged cells, hyperlinks, notes, the
 * autofilter, conditional formats, data validation, page breaks, defined names and tables (which
 * grow when rows are inserted inside them). Every other part of the file stays exactly as it was.
 * Sheets with pictures, charts, pivot tables or other content that xlvane cannot move are refused
 * with `XLV_E_NOT_SUPPORTED`, and so are
 * formulas it cannot rewrite safely (docs/facts/ref-adjust.md).
 *
 * @example
 * ```ts
 * import { insertRows } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * // Two rows before row 5 of "Sales", formatted like row 4.
 * const { bytes: updated, warnings } = insertRows(bytes, { sheet: 'Sales', at: 5, count: 2, style: 'above' });
 * console.log(updated.length, warnings.map((w) => w.code));
 * ```
 */
export function insertRows(input: Uint8Array | ArrayBuffer, options: InsertRowsOptions): EditResult {
  const o = checkOptions(options, true);
  return changeRows(input, o.sheet, o.at, o.count, o.style, o.limits);
}

/**
 * Deletes rows from a sheet of an existing XLSX file and returns the new file. Rows below move up,
 * and references follow as in a spreadsheet application: ranges shrink, references to deleted
 * cells become `#REF!`, merged cells whose top-left cell is deleted are removed, and so are notes,
 * hyperlinks and other ranges that lie entirely in the deleted rows. Refusals are as for
 * {@link insertRows}.
 *
 * @example
 * ```ts
 * import { deleteRows } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * // Rows 5 and 6 of "Sales".
 * const { bytes: updated } = deleteRows(bytes, { sheet: 'Sales', at: 5, count: 2 });
 * ```
 */
export function deleteRows(input: Uint8Array | ArrayBuffer, options: DeleteRowsOptions): EditResult {
  const o = checkOptions(options, false);
  return changeRows(input, o.sheet, o.at, -o.count, 'none', o.limits);
}
