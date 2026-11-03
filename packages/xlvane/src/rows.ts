// Surgical insertion and deletion of worksheet rows in existing XLSX files (docs/decisions.md
// D-B04). Built on the editor of edit.ts: only the parts that hold row numbers or references are
// rewritten, everything else keeps its exact content. References follow the rules of
// sml/ref-adjust.ts (docs/facts/ref-adjust.md R-01 … R-14); anything that cannot be rewritten
// safely makes the operation fail with an error instead of leaving a wrong reference behind.
import {
  IntList,
  PackageEditor,
  TextBuilder,
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
import { CHART_CONTENT_TYPE } from './charts';
import { XlvError, type XlvWarning } from './errors';
import { DRAWING_CONTENT_TYPE } from './images';
import { resolveLimits, type ReadLimits } from './limits';
import type { Sheet, Workbook } from './model';
import { openPackage, type OpcPackage } from './opc/package';
import { partKey } from './opc/part-name';
import { readWorkbook, styledCellCount } from './read';
import { formatCellRef, formatRange, MAX_ROWS, parseCellRef, parseRange, type CellRange } from './sml/cell-ref';
import { shiftFormula } from './sml/formula-shift';
import { classifyRelationship, REL_NS, SML_NS } from './sml/namespaces';
import { adjustFormula, adjustRange, type ReferenceChange } from './sml/ref-adjust';
import { X14_CF_URI, X14_DV_URI, X14_NS, XM_NS } from './sheet-rules';
import { readOptions } from './util/options';
import { WarningCollector } from './util/warnings';
import { parseXml } from './xml/parser';
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
   * The copied cell formats are cells that count against `limits.maxCells` with the cells of the
   * workbook (`XLV_E_LIMIT_EXCEEDED` above it).
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
/**
 * Media type of pivot cache definition parts, as Microsoft Excel writes it in
 * `[Content_Types].xml` (Apache POI corpus files poi/ExcelPivotTableSample.xlsx, poi/54436.xlsx,
 * poi/DataTableCities.xlsx; B6). The relationship from the workbook ends in
 * `/pivotCacheDefinition` in the same files.
 */
const PIVOT_CACHE_DEFINITION_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml';

const SAFE_SHEET_RELATIONSHIPS = new Set(['hyperlink', 'comments', 'vmlDrawing', 'printerSettings', 'image', 'customProperty', 'drawing', 'table']);

/** Elements of a worksheet whose content refers to cells and that xlvane does not rewrite. */
const REFUSED_ELEMENTS: Readonly<Record<string, string>> = {
  oleObjects: 'embedded objects',
  controls: 'form controls',
  scenarios: 'scenarios',
  cellWatches: 'cell watches',
  customSheetViews: 'custom sheet views',
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

/** Decodes the predefined entities and character references of XML text; a reference to a code point that does not exist stays as written. */
function decodeEntities(s: string): string {
  return s.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|amp|lt|gt|quot|apos);/g, (m, e: string) => {
    if (e === 'amp') return '&';
    if (e === 'lt') return '<';
    if (e === 'gt') return '>';
    if (e === 'quot') return '"';
    if (e === 'apos') return "'";
    const cp = e.startsWith('#x') ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
    return cp <= 0x10ffff ? String.fromCodePoint(cp) : m;
  });
}

/**
 * Whether the XML text of a part may refer to a sheet: its name, ignoring case, appears in the
 * text after entity and character references are decoded, as written or in the quoted form of a
 * formula, where each "'" is doubled (docs/facts/formulas.md F2-11). Searching the raw XML missed names such
 * as "R&D" (written "R&amp;D") and "O'Brien" (written "'O''Brien'!", B6). It may find the name
 * where it is not a reference, which only refuses more.
 */
function mentionsSheet(xml: string, sheet: string): boolean {
  const text = decodeEntities(xml).toLowerCase();
  const name = sheet.toLowerCase();
  return text.includes(name) || (name.includes("'") && text.includes(name.replace(/'/g, "''")));
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
  const out = adjustDecoded(text, c, formulaSheet);
  if (out === null) return null;
  return out === text ? raw : escapeText(out);
}

/** {@link adjustText} for formula text that is already decoded, such as an attribute value. */
function adjustDecoded(text: string, c: RowChange, formulaSheet: string): string | null {
  const out = adjustFormula(text, refChange(c, formulaSheet));
  if (out === null) {
    // A formula on another sheet without any sheet-qualified reference cannot refer to the changed sheet.
    return formulaSheet.toLowerCase() !== c.sheet.toLowerCase() && !text.includes('!') ? text : null;
  }
  return out;
}

/** `ext@uri` values of the worksheet extensions whose ranges and formulas are rewritten (SR-40, D-F04). */
const KNOWN_EXTENSIONS = new Set([X14_CF_URI, X14_DV_URI].map((u) => u.toLowerCase()));
const isKnownExtension = (e: Element): boolean => KNOWN_EXTENSIONS.has((e.attrs.get('uri') ?? '').trim().toLowerCase());
/** Start tags named `f` or `sqref`, with any prefix: the formulas and ranges an extension can hold. */
const FORMULA_OR_RANGE_TAG = /<(?:[\w.-]+:)?(?:f|sqref)[\s/>]/g;
const CFVO_PARENTS = new Set(['colorScale', 'dataBar', 'iconSet']);
const within = (outer: Element, inner: Element): boolean => inner.start.start >= outer.start.start && inner.end.end <= outer.end.end;

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
  /** The new text, as one string or as pieces for {@link PackageEditor.write}; the same string when nothing changed. */
  readonly text: string | string[];
  readonly hasFormulas: boolean;
}

const KEEP_DEPTH3 = new Set(['mergeCell', 'hyperlink', 'cfRule', 'dataValidation', 'brk', 'ignoredError', 'protectedRange']);

/** Integers per row of {@link WorksheetScan.rows}: number, start tag start and end, element end, flags, first cell, cell count. */
const ROW_FIELDS = 7;
/** Integers per cell of {@link WorksheetScan.cells}: column, start tag start and end, element end, flags. */
const CELL_FIELDS = 5;
/** Integers per formula of {@link WorksheetScan.formulas}: row, column, start tag start and end, end tag start and end, kind. */
const FORMULA_FIELDS = 7;
/** Row and cell flags. */
const HAS_R = 1;
const EMPTY = 2;
/** Formula kinds (`f@t`); any other value is handled as a normal formula. */
const NORMAL = 0;
const SHARED = 1;
const ARRAY = 2;
const DATA_TABLE = 3;

/**
 * A worksheet part as rewriteSheet needs it: the elements that hold references, as objects (they
 * do not grow with the cells), and the rows, cells and cell formulas as numbers, so that a sheet
 * of millions of cells costs a few integers per cell instead of objects (the element scanner kept
 * an object per cell and ran out of memory on 4.9 million cells).
 */
interface WorksheetScan {
  readonly elements: Element[];
  readonly rows: IntList;
  readonly rowCount: number;
  readonly cells: IntList;
  readonly formulas: IntList;
  readonly formulaCount: number;
  /** A row or cell without a reference: then every row and cell gets one. */
  readonly missingRefs: boolean;
  /** The first invalid row or cell reference, reported where the element scanner reported it. */
  readonly invalidRef: string | null;
}

function scanWorksheet(text: string, part: string, limits: ReadLimits): WorksheetScan {
  const keep = (local: string, depth: number, parent: string, ns: string): boolean =>
    depth === 2 ||
    (depth === 3 && (KEEP_DEPTH3.has(local) || (local === 'ext' && parent === 'extLst'))) ||
    (depth === 4 && ((local === 'formula' && parent === 'cfRule') || ((local === 'formula1' || local === 'formula2') && parent === 'dataValidation'))) ||
    (depth === 5 && local === 'cfvo' && CFVO_PARENTS.has(parent)) ||
    // The [MS-XLSX] extension: containers, items, rules, the x14:id of a classic rule, xm:sqref and xm:f (SR-40, SR-41).
    (ns === X14_NS && ((depth === 4 && (local === 'conditionalFormattings' || local === 'dataValidations')) || (depth === 5 && (local === 'conditionalFormatting' || local === 'dataValidation')) || (depth === 6 && (local === 'cfRule' || (local === 'id' && parent === 'ext'))))) ||
    (ns === XM_NS && ((depth === 6 && local === 'sqref') || (depth >= 7 && local === 'f')));
  const elements: Element[] = [];
  /** Per open element: the kept Element, or a code for a recorded row (1), cell (2) or formula (3), or null. */
  const open: (Element | 1 | 2 | 3 | null)[] = [];
  const names: string[] = [];
  let root = false;
  const rows = new IntList();
  const cells = new IntList();
  const formulas = new IntList();
  let rowCount = 0;
  let formulaCount = 0;
  let currentRow = 0;
  let lastCol = 0;
  let hasRow = false;
  let hasCell = false;
  let missingRefs = false;
  let invalidRef: string | null = null;
  parseXml(
    text,
    {
      start(ns, local, attrs, span) {
        const depth = names.length + 1;
        const parent = names[names.length - 1] ?? '';
        names.push(local);
        if (depth === 1) {
          root = true;
          open.push(null);
          return;
        }
        if (depth === 3 && local === 'row') {
          const r = attrs.get('r');
          const n = r === undefined ? currentRow + 1 : /^[0-9]{1,7}$/.test(r.trim()) ? Number(r.trim()) : Number.NaN;
          if (!Number.isInteger(n) || n < 1 || n > MAX_ROWS) invalidRef ??= r ?? String(n);
          currentRow = Number.isInteger(n) ? n : currentRow;
          lastCol = 0;
          if (r === undefined) missingRefs = true;
          rows.push(n);
          rows.push(span.start);
          rows.push(span.end);
          rows.push(span.end);
          rows.push((r === undefined ? 0 : HAS_R) | (span.empty ? EMPTY : 0));
          rows.push(cells.length / CELL_FIELDS);
          rows.push(0);
          rowCount++;
          hasRow = true;
          hasCell = false;
          open.push(1);
          return;
        }
        if (depth === 4 && local === 'c' && parent === 'row' && hasRow) {
          const r = attrs.get('r');
          const address = r === undefined || r === '' ? null : parseCellRef(r);
          if (r !== undefined && r !== '' && address === null) invalidRef ??= r;
          const col = address?.col ?? lastCol + 1;
          lastCol = col;
          const hasR = r !== undefined && r !== '';
          if (!hasR) missingRefs = true;
          cells.push(col);
          cells.push(span.start);
          cells.push(span.end);
          cells.push(span.end);
          cells.push((hasR ? HAS_R : 0) | (span.empty ? EMPTY : 0));
          const count = (rowCount - 1) * ROW_FIELDS + 6;
          rows.set(count, rows.get(count) + 1);
          hasCell = true;
          open.push(2);
          return;
        }
        if (depth === 5 && local === 'f' && parent === 'c' && hasCell) {
          const t = attrs.get('t');
          formulas.push(currentRow);
          formulas.push(lastCol);
          formulas.push(span.start);
          formulas.push(span.end);
          formulas.push(span.end);
          formulas.push(span.end);
          formulas.push(t === 'shared' ? SHARED : t === 'array' ? ARRAY : t === 'dataTable' ? DATA_TABLE : NORMAL);
          formulaCount++;
          open.push(3);
          return;
        }
        if (keep(local, depth, parent, ns)) {
          const element: Element = { local, ns, depth, attrs, start: span, end: span };
          elements.push(element);
          open.push(element);
        } else open.push(null);
      },
      end(_ns, _local, span) {
        names.pop();
        const what = open.pop();
        if (what === 1) {
          rows.set((rowCount - 1) * ROW_FIELDS + 3, span.end);
        } else if (what === 2) {
          cells.set(cells.length - 2, span.end);
        } else if (what === 3) {
          formulas.set(formulas.length - 3, span.start);
          formulas.set(formulas.length - 2, span.end);
        } else if (what !== null && what !== undefined) what.end = span;
      },
      text() {},
    },
    { part, limits },
  );
  if (!root) throw new XlvError('XLV_E_PART_INVALID', 'The part has no root element.', { location: { part }, hint: 'Re-save the file in a spreadsheet application.' });
  return { elements, rows, rowCount, cells, formulas, formulaCount, missingRefs, invalidRef };
}

/**
 * The attributes of a start tag given as source text (already checked by the XML parser), with
 * references decoded. Used for the few tags whose attributes are needed again, instead of keeping
 * the attributes of every cell and formula.
 */
function tagAttributes(tag: string): Map<string, string> {
  const attrs = new Map<string, string>();
  const pattern = /\s+([^\s=/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/y;
  pattern.lastIndex = tag.search(/[\s/>]/);
  for (let m = pattern.exec(tag); m !== null; m = pattern.exec(tag)) attrs.set(m[1] ?? '', decodeEntities((m[2] ?? m[3] ?? '').replace(/[\t\n\r]/g, ' ')));
  return attrs;
}

/** Rewrites one worksheet part: references everywhere, and row positions on the changed sheet. */
function rewriteSheet(text: string, part: string, sheet: Sheet, c: RowChange, target: boolean, style: 'above' | 'below' | 'none', limits: ReadLimits, workbookCells: number): SheetEdit {
  const scanned = scanWorksheet(text, part, limits);
  const { elements, rows, rowCount, cells, formulas, formulaCount } = scanned;
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
    }
  }
  // The conditional format and data validation extensions of [MS-XLSX] are moved when each of their
  // f and sqref elements is an xm:f or xm:sqref at the depth the schema gives (D-F04); one in
  // another layout (another namespace, mc:AlternateContent) is treated like other extensions.
  const knownExts = new Set<Element>();
  {
    let current: Element | null = null;
    let found = 0;
    const close = (): void => {
      if (current !== null && (text.slice(current.start.start, current.end.end).match(FORMULA_OR_RANGE_TAG) ?? []).length === found) knownExts.add(current);
    };
    for (const e of elements) {
      if (e.depth === 3 && e.local === 'ext') {
        close();
        current = isKnownExtension(e) ? e : null;
        found = 0;
      } else if (current !== null && e.ns === XM_NS && ((e.depth === 6 && e.local === 'sqref') || (e.depth >= 7 && e.local === 'f')) && within(current, e)) found++;
    }
    close();
  }
  // Other worksheet extensions: refused on the changed sheet when they refer to cells, and on other
  // sheets when they may refer to the changed one.
  for (const ext of elements) {
    if (ext.depth !== 3 || ext.local !== 'ext' || knownExts.has(ext)) continue;
    const body = text.slice(ext.start.start, ext.end.end);
    if (target && /sqref|<xm:f|sparkline|slicer/i.test(body)) {
      throw notSupported(`Sheet "${sheet.name}" has extended content (such as sparklines or slicers) that refers to cells, which xlvane cannot move with its rows yet.`, where(), { element: 'extLst' });
    }
    if (!target && /<xm:f|sparkline/i.test(body) && mentionsSheet(body, c.sheet)) {
      throw notSupported(`Sheet "${sheet.name}" has extended content (such as sparklines) that may refer to sheet "${c.sheet}", which xlvane cannot rewrite yet.`, where(), { element: 'extLst' });
    }
  }

  // Rows and cells are numbered as the reader numbers them (D-105).
  if (scanned.invalidRef !== null) {
    const ref = scanned.invalidRef;
    throw new XlvError('XLV_E_INVALID_CELL_REF', `Sheet "${sheet.name}" contains the invalid reference "${ref.slice(0, 20)}", so its rows cannot be moved safely.`, {
      location: where(),
      hint: 'Open and re-save the file in a spreadsheet application, which drops the damaged cells, then try again.',
      details: { ref: ref.slice(0, 40) },
    });
  }
  const hasFormulas = formulaCount > 0;
  let extension: Element | null = null;
  for (const e of elements) {
    if (e.depth === 3 && e.local === 'ext') extension = knownExts.has(e) ? e : null;
    if (e.depth === 4 && (e.local === 'formula' || e.local === 'formula1' || e.local === 'formula2')) {
      const raw = content(text, e);
      const out = adjustText(raw, c, sheet.name);
      if (out === null) {
        const owner = e.local === 'formula' ? 'conditional format' : 'data validation';
        throw notSupported(`A ${owner} formula of sheet "${sheet.name}" cannot be adjusted safely: "${decodeEntities(raw).slice(0, 60)}".`, where(), { formula: decodeEntities(raw).slice(0, 80) });
      }
      if (out !== raw) edits.push([e.start.end, e.end.start, out]);
    } else if (e.depth === 5 && e.local === 'cfvo') {
      // A threshold is a formula (SR-12), such as a reference to a cell (R-18).
      const val = e.attrs.get('val');
      if (val === undefined) continue;
      const out = adjustDecoded(val, c, sheet.name);
      if (out === null) throw notSupported(`A conditional format threshold of sheet "${sheet.name}" cannot be adjusted safely: "${val.slice(0, 60)}".`, where(), { formula: val.slice(0, 80) });
      if (out !== val) replaceTag(e, setAttr(tagOf(e), 'val', out));
    } else if (e.ns === XM_NS && e.local === 'f' && extension !== null && within(extension, e)) {
      // xm:f of an extension rule, threshold or validation (SR-42, SR-44; R-19).
      const raw = content(text, e);
      const out = adjustText(raw, c, sheet.name);
      if (out === null) {
        const owner = (extension.attrs.get('uri') ?? '').toLowerCase() === X14_DV_URI.toLowerCase() ? 'data validation' : 'conditional format';
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

  // Cell formulas, with shared formulas expanded where they cannot stay shared. The changes are
  // worked out while the new text is written, one formula at a time, so they are never all held.
  const masters = new Map<string, { row: number; col: number; text: string }>();
  for (let i = 0; i < formulaCount; i++) {
    const at = i * FORMULA_FIELDS;
    if (formulas.get(at + 6) !== SHARED) continue;
    const attrs = tagAttributes(text.slice(formulas.get(at + 2), formulas.get(at + 3)));
    if (attrs.get('t') === 'shared' && attrs.has('ref') && attrs.has('si')) masters.set(attrs.get('si') ?? '', { row: formulas.get(at), col: formulas.get(at + 1), text: text.slice(formulas.get(at + 3), formulas.get(at + 4)) });
  }
  const expand = new Set<string>();
  for (const [si, m] of masters) {
    const out = adjustText(m.text, c, sheet.name);
    if (target || out === null || out !== m.text) expand.add(si);
  }
  const deletedRow = (n: number): boolean => c.count < 0 && moveRow(n, c) === null;
  /** The changes of formula `i`: [start, end, text] replacements in document order. */
  const formulaChanges = (i: number): [number, number, string][] => {
    const at = i * FORMULA_FIELDS;
    const row = formulas.get(at);
    const col = formulas.get(at + 1);
    const tagStart = formulas.get(at + 2);
    const tagEnd = formulas.get(at + 3);
    const endStart = formulas.get(at + 4);
    const end = formulas.get(at + 5);
    const kind = formulas.get(at + 6);
    if (target && deletedRow(row)) return [];
    const ref = formatCellRef(row, col);
    const failed = (formula: string): never => {
      throw notSupported(`The formula in ${ref} of sheet "${sheet.name}" cannot be adjusted safely: "${formula.slice(0, 60)}".`, where(ref), { formula: formula.slice(0, 80) });
    };
    const tag = text.slice(tagStart, tagEnd);
    if (kind === SHARED) {
      const si = tagAttributes(tag).get('si') ?? '';
      const master = masters.get(si);
      if (master === undefined || !expand.has(si)) return [];
      const source = decodeEntities(master.text);
      const own = shiftFormula(source, row - master.row, col - master.col);
      if (own === null) failed(source);
      const adjusted = adjustFormula(own as string, refChange(c, sheet.name));
      if (adjusted === null) failed(own as string);
      let plain = removeAttr(removeAttr(removeAttr(tag, 't'), 'ref'), 'si');
      if (plain.endsWith('/>')) plain = `${plain.slice(0, -2).trimEnd()}>`;
      const qname = /^<([^\s/>]+)/.exec(tag)?.[1] ?? 'f';
      return [[tagStart, end, `${plain}${escapeText(adjusted as string)}</${qname}>`]];
    }
    if (kind === DATA_TABLE) {
      if (target) throw notSupported(`Cell ${ref} of sheet "${sheet.name}" is part of a data table, which xlvane cannot move.`, where(ref));
      return [];
    }
    const raw = text.slice(tagEnd, endStart);
    const out = adjustText(raw, c, sheet.name);
    if (out === null) failed(decodeEntities(raw));
    const changes: [number, number, string][] = [];
    if (kind === ARRAY && target) {
      const arrayRef = tagAttributes(tag).get('ref') ?? '';
      const range = parseRange(arrayRef);
      if (range !== null) {
        const last = c.at - c.count - 1;
        const split = c.count > 0 ? c.at > range.start.row && c.at <= range.end.row : !(range.start.row >= c.at && range.end.row <= last) && range.end.row >= c.at && range.start.row <= last;
        if (split) throw notSupported(`The array formula in ${arrayRef === '' ? ref : arrayRef} of sheet "${sheet.name}" would be split by the change.`, where(ref));
        const moved = adjustRange(arrayRef, refChange(c, sheet.name));
        if (moved !== null && moved !== '' && moved !== arrayRef) changes.push([tagStart, tagEnd, setAttr(tag, 'ref', moved)]);
      }
    }
    if (out !== raw) changes.push([tagEnd, endStart, out as string]);
    return changes;
  };

  let force = false;
  if (target) {
    force = scanned.missingRefs;
    // Rows pushed past the last row must be empty.
    for (let i = 0; i < rowCount && c.count > 0; i++) {
      const at = i * ROW_FIELDS;
      const number = rows.get(at);
      if (moveRow(number, c) !== null) continue;
      for (let j = 0; j < rows.get(at + 6); j++) {
        if ((cells.get((rows.get(at + 5) + j) * CELL_FIELDS + 4) & EMPTY) === 0) {
          throw new XlvError('XLV_E_LIMIT_EXCEEDED', `Inserting ${String(c.count)} rows would push row ${String(number)} of sheet "${sheet.name}", which holds data, past the last row of the sheet.`, {
            location: where(),
            hint: `A sheet has ${String(MAX_ROWS)} rows. Delete the rows at the end of the sheet first, or insert fewer rows.`,
            details: { limit: 'rows', value: MAX_ROWS },
          });
        }
      }
    }
    const findRow = (test: (n: number) => boolean): number => {
      for (let i = 0; i < rowCount; i++) if (test(rows.get(i * ROW_FIELDS))) return i;
      return -1;
    };
    if (c.count > 0 && style !== 'none') {
      const source = findRow((n) => n === (style === 'above' ? c.at - 1 : c.at));
      if (source >= 0) {
        const at = source * ROW_FIELDS;
        const rowTag = text.slice(rows.get(at + 1), rows.get(at + 2));
        const rowAttrs = tagAttributes(rowTag);
        const rowQname = /^<([^\s/>]+)/.exec(rowTag)?.[1] ?? 'row';
        const prefix = rowQname.slice(0, rowQname.length - 3);
        const keep = ['s', 'customFormat', 'ht', 'customHeight'].filter((a) => rowAttrs.has(a));
        const formats: { col: number; s: string }[] = [];
        for (let j = 0; j < rows.get(at + 6); j++) {
          const cell = (rows.get(at + 5) + j) * CELL_FIELDS;
          const s = tagAttributes(text.slice(cells.get(cell + 1), cells.get(cell + 2))).get('s') ?? '0';
          if (s !== '0') formats.push({ col: cells.get(cell), s });
        }
        // The new cells are counted before they are built: copying the formats of a wide row into
        // many rows could otherwise ask for billions of cells.
        const added = formats.length * c.count;
        if (workbookCells + added > limits.maxCells) {
          throw new XlvError('XLV_E_LIMIT_EXCEEDED', `Inserting ${String(c.count)} rows formatted like row ${String(rows.get(at))} of sheet "${sheet.name}" would add ${String(added)} formatted cells, and the workbook would hold more cells than the limit of ${String(limits.maxCells)}.`, {
            location: where(),
            hint: `Cells that hold only a style count too. Insert fewer rows, use { style: 'none' } and format the rows afterwards, or, if there is enough memory, raise the limit, for example { limits: { maxCells: ${String(workbookCells + added)} } }.`,
            details: { limit: 'maxCells', value: limits.maxCells },
          });
        }
        const attrs = keep.map((a) => ` ${a}="${escapeAttr(rowAttrs.get(a) ?? '')}"`).join('');
        const created = new TextBuilder();
        for (let i = 0; i < c.count; i++) {
          const n = c.at + i;
          if (formats.length === 0) {
            created.add(`<${prefix}row r="${String(n)}"${attrs}/>`);
            continue;
          }
          created.add(`<${prefix}row r="${String(n)}"${attrs}>`);
          for (const f of formats) created.add(`<${prefix}c r="${formatCellRef(n, f.col)}" s="${escapeAttr(f.s)}"/>`);
          created.add(`</${prefix}row>`);
        }
        const createdText = created.finish().join('');
        const next = findRow((n) => n >= c.at);
        if (next >= 0) edits.push([rows.get(next * ROW_FIELDS + 1), rows.get(next * ROW_FIELDS + 1), createdText]);
        else {
          const sheetData = elements.find((e) => e.depth === 2 && e.local === 'sheetData');
          if (sheetData && !sheetData.start.empty) edits.push([sheetData.end.start, sheetData.end.start, createdText]);
          else if (sheetData) edits.push([sheetData.start.start, sheetData.start.end, `${tagOf(sheetData).slice(0, -2).trimEnd()}>${createdText}</${sheetData.start.qname}>`]);
        }
      }
    }

    // Elements of the changed sheet that hold ranges.
    const removedIds = new Set<string>();
    // The x14:id values of the rules of each classic conditional format, collected in one pass.
    const linkIds = new Map<Element, string[]>();
    let format: Element | null = null;
    for (const e of elements) {
      if (e.depth === 2) format = e.local === 'conditionalFormatting' ? e : null;
      else if (format !== null && e.depth === 6 && e.ns === X14_NS && e.local === 'id') {
        const list = linkIds.get(format) ?? [];
        list.push(content(text, e).trim().toLowerCase());
        linkIds.set(format, list);
      }
    }
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
          if (out === '') {
            removeElement(e);
            // The extension parts of its rules go too (the x14:id link, SR-41).
            for (const id of linkIds.get(e) ?? []) removedIds.add(id);
          } else if (out !== sqref) replaceTag(e, setAttr(tagOf(e), 'sqref', out));
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

    rewriteExtensionRanges(text, elements, knownExts, c, sheet, removedIds, edits, where);
  } else if (edits.length === 0) {
    // Another sheet: most often no formula refers to the changed one, and the text stays as it is.
    let changed = false;
    for (let i = 0; i < formulaCount && !changed; i++) changed = formulaChanges(i).length > 0;
    if (!changed) return { text, hasFormulas };
  }

  // The new text in one pass, in document order: the edits collected above, interleaved with the
  // rows and cells (new references) and the formulas (adjusted), written as they are produced.
  // As in applyEdits, edits are sorted by start, insertions first, and an edit that starts inside
  // a replaced span is dropped with it.
  const general = [...edits].sort((a, b) => a[0] - b[0] || Number(b[0] === b[1]) - Number(a[0] === a[1]) || b[1] - a[1]);
  const out = new TextBuilder();
  let pos = 0;
  let g = 0;
  const flush = (limit: number): void => {
    while (g < general.length) {
      const [start, end, replacement] = general[g]!;
      if (!(start < limit || (start === limit && end === start))) break;
      g++;
      if (start < pos) continue;
      out.add(text.slice(pos, start));
      out.add(replacement);
      pos = end;
    }
  };
  const replace = (start: number, end: number, replacement: string): void => {
    flush(start);
    if (start < pos) return;
    out.add(text.slice(pos, start));
    out.add(replacement);
    pos = end;
  };
  let f = 0;
  /** Formulas that start before `limit`: adjusted, or skipped when `skip`. */
  const formulasBefore = (limit: number, skip: boolean): void => {
    for (; f < formulaCount && formulas.get(f * FORMULA_FIELDS + 2) < limit; f++) {
      if (skip) continue;
      for (const [start, end, replacement] of formulaChanges(f)) replace(start, end, replacement);
    }
  };
  for (let i = 0; i < rowCount; i++) {
    const at = i * ROW_FIELDS;
    const number = rows.get(at);
    const rowStart = rows.get(at + 1);
    const rowTagEnd = rows.get(at + 2);
    const rowEnd = rows.get(at + 3);
    formulasBefore(rowStart, false);
    const n = target ? moveRow(number, c) : number;
    if (n === null) {
      // A deleted row, or an empty row pushed past the last row.
      replace(rowStart, rowEnd, '');
      formulasBefore(rowEnd, true);
      continue;
    }
    const moved = target && (n !== number || force);
    if (moved) replace(rowStart, rowTagEnd, setAttr(text.slice(rowStart, rowTagEnd), 'r', String(n)));
    const first = rows.get(at + 5);
    for (let j = 0; j < rows.get(at + 6); j++) {
      const cell = (first + j) * CELL_FIELDS;
      const cellStart = cells.get(cell + 1);
      formulasBefore(cellStart, false);
      if (moved) replace(cellStart, cells.get(cell + 2), setAttr(text.slice(cellStart, cells.get(cell + 2)), 'r', formatCellRef(n, cells.get(cell))));
      formulasBefore(cells.get(cell + 3), false);
    }
    formulasBefore(rowEnd, false);
  }
  formulasBefore(text.length + 1, false);
  flush(text.length + 1);
  out.add(text.slice(pos));
  const pieces = out.finish();
  // Sort states are rare and small: they are moved on the whole text (D-B01).
  if (target && /sortState/.test(text)) return { text: rewriteSortStates(pieces.join(''), part, sheet, c, limits), hasFormulas };
  return { text: pieces, hasFormulas };
}

/**
 * Moves the ranges (`xm:sqref`) of the extension conditional formats and data validations of the
 * changed sheet (SR-40, SR-44; R-20; D-F04). An item whose cells are all deleted is removed, as is
 * an extension rule linked to a removed classic rule (SR-41); emptied containers, `ext` and
 * `extLst` elements are removed, and `x14:dataValidations@count` is updated.
 */
function rewriteExtensionRanges(text: string, elements: readonly Element[], knownExts: ReadonlySet<Element>, c: RowChange, sheet: Sheet, removedIds: ReadonlySet<string>, edits: [number, number, string][], where: () => { part: string; sheet: string }): void {
  // The extension elements as a tree, built in one pass in document order (elements nest by span).
  interface Node {
    readonly e: Element;
    readonly children: Node[];
  }
  const roots: Node[] = [];
  const stack: Node[] = [];
  for (const e of elements) {
    const relevant = (e.depth === 2 && e.local === 'extLst') || (e.depth === 3 && e.local === 'ext') || (e.ns === X14_NS && e.depth >= 4 && e.depth <= 6) || (e.ns === XM_NS && e.depth === 6 && e.local === 'sqref');
    if (!relevant) continue;
    while (stack.length > 0 && !within(stack[stack.length - 1]!.e, e)) stack.pop();
    const node: Node = { e, children: [] };
    (stack.length > 0 ? stack[stack.length - 1]!.children : roots).push(node);
    stack.push(node);
  }
  const tagOf = (e: Element): string => text.slice(e.start.start, e.start.end);
  const remove = (e: Element): void => {
    edits.push([e.start.start, e.end.end, '']);
  };
  for (const list of roots) {
    if (list.e.local !== 'extLst') continue;
    let extsKept = 0;
    for (const ext of list.children) {
      if (ext.e.local !== 'ext' || !knownExts.has(ext.e)) {
        extsKept++;
        continue;
      }
      let containersKept = 0;
      let containersTotal = 0;
      for (const container of ext.children) {
        if (container.e.depth !== 4 || container.e.ns !== X14_NS) continue;
        containersTotal++;
        let kept = 0;
        let total = 0;
        for (const item of container.children) {
          if (item.e.depth !== 5 || item.e.ns !== X14_NS) continue;
          total++;
          let gone = false;
          const sqref = item.children.find((n) => n.e.ns === XM_NS && n.e.local === 'sqref')?.e;
          if (sqref) {
            const raw = decodeEntities(content(text, sqref));
            const normal = raw.trim().split(/\s+/).join(' ');
            // An empty range is left as it is; adjustRange would read it as deleted.
            const out = normal === '' ? '' : adjustRange(normal, refChange(c, sheet.name));
            if (out === null) throw notSupported(`An extension ${item.e.local === 'dataValidation' ? 'data validation' : 'conditional format'} of sheet "${sheet.name}" has a range xlvane cannot adjust: "${normal.slice(0, 60)}".`, where());
            if (out === '' && normal !== '') gone = true;
            else if (out !== normal) edits.push([sqref.start.end, sqref.end.start, escapeText(out)]);
          }
          if (!gone && removedIds.size > 0) {
            const rules = item.children.filter((n) => n.e.ns === X14_NS && n.e.local === 'cfRule').map((n) => n.e);
            const linked = rules.filter((r) => removedIds.has((r.attrs.get('id') ?? '').trim().toLowerCase()));
            if (linked.length > 0 && linked.length === rules.length) gone = true;
            else for (const r of linked) remove(r);
          }
          if (gone) remove(item.e);
          else kept++;
        }
        if (kept === 0 && total > 0) {
          // An emptied container goes; if no container stays, the ext goes below.
          remove(container.e);
          continue;
        }
        containersKept++;
        if (kept < total && container.e.attrs.has('count')) edits.push([container.e.start.start, container.e.start.end, setAttr(tagOf(container.e), 'count', String(kept))]);
      }
      if (containersTotal > 0 && containersKept === 0) remove(ext.e);
      else extsKept++;
    }
    if (list.children.length > 0 && extsKept === 0) remove(list.e);
  }
}

/**
 * Moves the sort ranges of the changed sheet (`sortState` at sheet level or inside `autoFilter`,
 * and their `sortCondition`s) as ranges move (D-B01). A sort range that is deleted entirely is
 * removed; one that cannot be adjusted is refused. Runs on the text after the other edits, so
 * a sort state inside a removed filter is already gone.
 */
function rewriteSortStates(text: string, part: string, sheet: Sheet, c: RowChange, limits: ReadLimits): string {
  const { elements } = scan(text, part, limits, (local) => local === 'sortState' || local === 'sortCondition');
  const edits: [number, number, string][] = [];
  const removed: [number, number][] = [];
  for (const e of elements) {
    if (removed.some(([a, b]) => e.start.start >= a && e.end.end <= b)) continue;
    const ref = e.attrs.get('ref');
    if (ref === undefined) continue;
    const out = adjustRange(ref, refChange(c, sheet.name));
    if (out === null) throw notSupported(`A sort range of sheet "${sheet.name}" cannot be adjusted: "${ref.slice(0, 60)}".`, { part, sheet: sheet.name });
    if (out === '') {
      edits.push([e.start.start, e.end.end, '']);
      removed.push([e.start.start, e.end.end]);
    } else if (out !== ref) edits.push([e.start.start, e.start.end, setAttr(text.slice(e.start.start, e.start.end), 'ref', out)]);
  }
  return edits.length === 0 ? text : applyEdits(text, edits);
}

/** A `from` or `to` marker of a drawing anchor: its row, row offset and where the row value is. */
interface AnchorMarker {
  readonly row: number;
  readonly rowOff: number;
  /** Offsets of the row value (between `<row>` and `</row>`) in the anchor text. */
  readonly rowStart: number;
  readonly rowEnd: number;
}

/**
 * Reads the `row` and `rowOff` of the first `from` or `to` marker of an anchor with `indexOf`
 * only, so that the work is linear in the anchor's size whatever its content (a regular
 * expression here backtracked cubically on runs of white space, D-R03).
 */
function readMarker(anchor: string, pre: string, local: 'from' | 'to'): AnchorMarker | null {
  const open = `<${pre}${local}>`;
  const start = anchor.indexOf(open);
  if (start < 0) return null;
  const close = anchor.indexOf(`</${pre}${local}>`, start + open.length);
  if (close < 0) return null;
  const child = (name: string): { value: string; start: number; end: number } | null => {
    const tag = `<${pre}${name}>`;
    const a = anchor.indexOf(tag, start + open.length);
    if (a < 0 || a > close) return null;
    const b = anchor.indexOf(`</${pre}${name}>`, a + tag.length);
    if (b < 0 || b > close) return null;
    return { value: anchor.slice(a + tag.length, b), start: a + tag.length, end: b };
  };
  const row = child('row');
  const rowOff = child('rowOff');
  if (row === null || rowOff === null || !/^\s*\d+\s*$/.test(row.value)) return null;
  return { row: Number(row.value), rowOff: Number(rowOff.value.trim()) || 0, rowStart: row.start, rowEnd: row.end };
}

/**
 * Calls `change` for each `<v:shape …>…</v:shape>` of a VML part and joins the results, finding
 * the elements with `indexOf` so that the work is linear even in a part that is not well-formed
 * (a lazy regular expression here was quadratic on unclosed shapes, D-R03).
 */
function eachVmlShape(text: string, change: (shape: string) => string): string {
  const out: string[] = [];
  let pos = 0;
  let from = 0;
  for (;;) {
    const at = text.indexOf('<v:shape', from);
    if (at < 0) break;
    const next = text.charCodeAt(at + 8);
    // `<v:shapetype` and other names that only start with "v:shape" are not shapes.
    if (!(next === 0x20 || next === 0x3e || next === 0x2f || next === 0x09 || next === 0x0a || next === 0x0d)) {
      from = at + 8;
      continue;
    }
    const end = text.indexOf('</v:shape>', at + 8);
    if (end < 0) break;
    const stop = end + '</v:shape>'.length;
    out.push(text.slice(pos, at), change(text.slice(at, stop)));
    pos = stop;
    from = stop;
  }
  out.push(text.slice(pos));
  return out.join('');
}

/** The relationship ids of a worksheet's `drawing` and `legacyDrawing` elements. */
interface DrawingIds {
  readonly drawing: string | undefined;
  readonly legacyDrawing: string | undefined;
}

/**
 * Finds the first `drawing` and `legacyDrawing` elements of the worksheet (SpreadsheetML
 * namespace, any depth, so also inside `mc:AlternateContent`) with the element scanner, so that
 * comments are skipped and the time is linear in the part's size (D-H01); the id is the
 * attribute `id` in the relationships namespace, whatever its prefix.
 */
function drawingIds(text: string, part: string, limits: ReadLimits): DrawingIds {
  const { elements } = scan(text, part, limits, (local, depth, _parent, ns) => depth >= 2 && (local === 'drawing' || local === 'legacyDrawing') && (ns === SML_NS.transitional || ns === SML_NS.strict));
  const idOf = (local: string): string | undefined => {
    const e = elements.find((x) => x.local === local);
    return e?.attrs.get(`{${REL_NS.transitional}}id`) ?? e?.attrs.get(`{${REL_NS.strict}}id`);
  };
  return { drawing: idOf('drawing'), legacyDrawing: idOf('legacyDrawing') };
}

/** Rewrites the notes of the changed sheet: the Comments part and the note shapes of its VML part. */
function rewriteNotes(edit: PackageEditor, pkg: OpcPackage, sheetPart: string, ids: DrawingIds, sheet: Sheet, c: RowChange): void {
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
      if (edits.length > 0) edit.write(comments.target, applyEdits(text, edits));
    }
  }
  // The VML part of the notes is the one named by legacyDrawing (W2-48).
  const id = ids.legacyDrawing;
  const vml = id === undefined ? undefined : rels.find((r) => r.id === id && !r.external);
  if (vml === undefined) return;
  const text = edit.text(vml.target);
  if (text === undefined) return;
  let changed = false;
  const out = eachVmlShape(text, (shape) => {
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
function rewriteDrawing(edit: PackageEditor, pkg: OpcPackage, sheetPart: string, ids: DrawingIds, sheet: Sheet, c: RowChange): void {
  const id = ids.drawing;
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
  // The anchors are the children of `wsDr`, located with the element scanner (linear in the
  // part's size and bounded by the read limits, D-D05); only anchors holding a picture move.
  const { elements } = scan(text, rel.target, edit.limits, (_local, depth) => depth === 2);
  const anchorKinds = new Set(['twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor']);
  const others = (): XlvError => notSupported(`The drawing of sheet "${sheet.name}" has charts, shapes or other objects besides pictures, which xlvane cannot move with its rows yet.`, where);
  for (const e of elements) {
    if (!anchorKinds.has(e.local) || e.start.empty) throw others();
    const pre = e.start.qname.slice(0, e.start.qname.length - e.local.length);
    if (!new RegExp(`<${pre.replace(/[.]/g, '\\.')}pic[\\s>]`).test(text.slice(e.start.start, e.end.end))) throw others();
  }
  const end = c.at - c.count - 1;
  // A marker at the start of row m + 1 (0-based m): its new row, as a range end moves (R-03, R-08).
  // A `to` marker with no offset that lies on the boundary below the last covered row (0-based
  // m = 1-based last row) moves as that last row does, so rows inserted just below the picture
  // leave it as it is, as LibreOffice 24.2.7.2 does (D-D05).
  const moveMarker = (m: number, boundary = false): number => {
    const one = boundary ? m : m + 1;
    if (c.count > 0) return one >= c.at ? m + c.count : m;
    return one < c.at ? m : one > end ? m + c.count : c.at - 1;
  };
  const edits: [number, number, string][] = [];
  for (const e of elements) {
    const kind = e.local;
    if (kind === 'absoluteAnchor') continue;
    const anchor = text.slice(e.start.start, e.end.end);
    const pre = e.start.qname.slice(0, e.start.qname.length - kind.length);
    const marker = (local: 'from' | 'to'): AnchorMarker | null => readMarker(anchor, pre, local);
    const from = marker('from');
    const to = kind === 'twoCellAnchor' ? marker('to') : null;
    if (from === null || (kind === 'twoCellAnchor' && to === null)) throw notSupported(`A picture of sheet "${sheet.name}" has a position xlvane cannot read, so its rows cannot be moved.`, where);
    const first = from.row + 1;
    const last = to === null ? first : to.rowOff === 0 && to.row > from.row ? to.row : to.row + 1;
    if (c.count < 0 && first <= end && last >= c.at) throw notSupported(`Deleting rows ${String(c.at)} to ${String(end)} of sheet "${sheet.name}" would cut a picture (rows ${String(first)} to ${String(last)}), which xlvane does not do.`, where, { cell: formatCellRef(first, 1) });
    const rows = [moveMarker(from.row), to === null ? null : moveMarker(to.row, to.rowOff === 0 && to.row > from.row)];
    if (rows.some((r) => r !== null && r > MAX_ROWS - 1)) throw notSupported(`Inserting rows would push a picture of sheet "${sheet.name}" beyond the last row.`, where);
    if (rows[0] === from.row && (to === null || rows[1] === to.row)) continue;
    const markerEdits: [number, number, string][] = [[from.rowStart, from.rowEnd, String(rows[0])]];
    if (to !== null && rows[1] !== null && rows[1] !== undefined) markerEdits.push([to.rowStart, to.rowEnd, String(rows[1])]);
    edits.push([e.start.start, e.end.end, applyEdits(anchor, markerEdits)]);
  }
  if (edits.length > 0) edit.write(rel.target, applyEdits(text, edits));
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
    let derivedFilter: { from: string; to: string } | null = null;
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
      // A filter in the form xlvane and Excel write (header and data rows, TD-08) is derived from
      // the new table range, so rows inserted just above the totals row join it (D-D05).
      if (header === 1 && moved !== null) {
        const filterOf = (r: CellRange): string => formatRange({ start: r.start, end: { col: r.end.col, row: r.end.row - totals } });
        derivedFilter = { from: filterOf(range), to: filterOf(moved) };
      }
      if (newRef !== ref) edits.push([root.start.start, root.start.end, setAttr(text.slice(root.start.start, root.start.end), 'ref', newRef)]);
    }
    for (const e of elements) {
      if (e.local === 'autoFilter') {
        if (!target) continue;
        const ref = e.attrs.get('ref');
        if (ref === undefined) continue;
        const out = derivedFilter !== null && ref.trim().toUpperCase() === derivedFilter.from ? derivedFilter.to : adjustRange(ref, refChange(c, sheet.name));
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
  // Charts and pivot caches anywhere may refer to the sheet by name. Chart parts are found by
  // their content type (CH-01) and as the targets of the chart relationships of drawing parts
  // (CH-04), whatever their names; pivot cache definitions by their content type and as the
  // targets of the pivot cache relationships of the workbook (B6); other XML parts in a charts
  // folder and parts named like pivot cache definitions by their names, as before. Drawings are
  // found by their content type and as the targets of the drawing relationships of the sheets.
  const charts = new Set<string>();
  const drawings = new Map<string, string>();
  const unmatched = (source: string, r: { type: string; target: string }): XlvError =>
    notSupported(`The part ${source} has a relationship to ${r.target}, which is not a part of the package (its name may be written differently in the archive), so xlvane cannot check whether a chart there refers to sheet "${sheet.name}".`, { part: source, sheet: sheet.name }, { type: r.type, target: r.target });
  for (const name of pkg.partNames) {
    const type = pkg.contentType(name);
    if (type === CHART_CONTENT_TYPE || type === PIVOT_CACHE_DEFINITION_CONTENT_TYPE) charts.add(partKey(name));
    else if (type === DRAWING_CONTENT_TYPE) drawings.set(partKey(name), name);
  }
  for (const s of workbook.sheets) {
    if (s.part === null) continue;
    for (const r of pkg.relationships(s.part)) {
      if (r.external || classifyRelationship(r.type)?.name !== 'drawing') continue;
      // A drawing that cannot be found may hold charts that refer to the sheet (B7).
      if (!pkg.hasPart(r.target)) throw unmatched(s.part, r);
      drawings.set(partKey(r.target), r.target);
    }
  }
  for (const name of drawings.values()) {
    for (const r of pkg.relationships(name)) {
      if (r.external || classifyRelationship(r.type)?.name !== 'chart') continue;
      // A chart part whose name the archive spells differently (B7) is not a part, so it was missed.
      if (!pkg.hasPart(r.target)) throw unmatched(name, r);
      charts.add(partKey(r.target));
    }
  }
  for (const r of pkg.relationships(workbookPart)) {
    if (r.external || classifyRelationship(r.type)?.name !== 'pivotCacheDefinition') continue;
    if (!pkg.hasPart(r.target)) throw unmatched(workbookPart, r);
    charts.add(partKey(r.target));
  }
  for (const name of pkg.partNames) {
    if (!charts.has(partKey(name)) && !/\/charts\/[^/]+\.xml$|pivotCacheDefinition[^/]*\.xml$/i.test(name)) continue;
    const text = edit.text(name);
    if (text !== undefined && mentionsSheet(text, sheet.name)) {
      throw notSupported(`The part ${name} (a chart or pivot table) may refer to sheet "${sheet.name}", and xlvane cannot rewrite it yet.`, { part: name, sheet: sheet.name });
    }
  }

  let hasFormulas = false;
  // The cells of the workbook as read (with a value or only a style), for the formatted cells of
  // inserted rows, which count against maxCells too.
  let workbookCells = 0;
  for (const s of workbook.sheets) workbookCells += s.cellCount + styledCellCount(s);
  for (const s of workbook.sheets) {
    if (s.kind !== 'worksheet' || s.part === null) continue;
    const text = edit.text(s.part);
    if (text === undefined) continue;
    const target = s.part === sheet.part;
    const out = rewriteSheet(text, s.part, s, c, target, style, limits, workbookCells);
    hasFormulas ||= out.hasFormulas;
    if (out.text !== text) edit.write(s.part, out.text);
    rewriteTables(edit, pkg, s.part, s, c, target);
    if (target) {
      const ids = drawingIds(text, s.part, edit.limits);
      rewriteNotes(edit, pkg, s.part, ids, s, c);
      rewriteDrawing(edit, pkg, s.part, ids, s, c);
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
 * autofilter, conditional formats (also those of the [MS-XLSX] extension, and thresholds that
 * refer to cells), data validation (also of the extension), page breaks, defined names, tables
 * (which grow when rows are inserted inside them) and pictures placed on the sheet. Every other
 * part of the file stays exactly as it was.
 *
 * Refused with `XLV_E_NOT_SUPPORTED`, leaving the file unchanged: a drawing that holds charts,
 * shapes or anything besides pictures; embedded objects, form controls, scenarios, cell watches,
 * custom sheet views, other extended content that refers to cells (sparklines, slicers) and
 * other related parts xlvane does not know; charts or pivot tables anywhere in the workbook that
 * may refer to the sheet; data tables; an array formula the change would split; and formulas it
 * cannot rewrite safely (docs/facts/ref-adjust.md). Rows formatted with `style` add a formatted
 * cell for each formatted cell of the row they copy; when the workbook would then hold more cells
 * than `limits.maxCells`, the call throws `XLV_E_LIMIT_EXCEEDED` before building them.
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
 * {@link insertRows}; in addition a deletion that would cut a picture, or remove a table's header
 * row, totals row or all of its data rows, is refused.
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
