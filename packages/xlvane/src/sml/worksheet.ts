// Streams a worksheet part through the SAX parser and collects its cells (ECMA-376 Part 1
// §18.3.1.4 c, §18.3.1.73 row, §18.3.1.40 f, §18.3.1.96 v, §18.3.1.55 mergeCell,
// §18.3.1.35 dimension; docs/facts/ecma-376.md E-20 … E-39).
import { XlvError } from '../errors';
import type { ReadLimits } from '../limits';
import type { Cell, CellType, CellValue, ColumnInfo, Conformance, DateSystem, Formula, FormulaKind, FrozenPanes, RowInfo } from '../model';
import type { WarningCollector } from '../util/warnings';
import { parseXml, type XmlHandler } from '../xml/parser';
import { formatCellRef, formatRange, MAX_COLUMNS, MAX_ROWS, parseCellRef, parseRange, type CellRange } from './cell-ref';
import { compileFormula, type CompiledFormula } from './formula-shift';
import { parseIsoDateTime, serialToDate } from './dates';
import { SML_NS } from './namespaces';
import { BUILTIN_NUMBER_FORMATS, isDateFormat, isDateFormatId } from './numfmt';
import type { Styles } from './styles';
import { parseXsdBoolean } from './workbook';
import { decodeXstring } from './xstring';

/** Everything a worksheet needs from the rest of the package. */
export interface WorksheetContext {
  readonly part: string;
  readonly sheet: string;
  readonly limits: ReadLimits;
  readonly warnings: WarningCollector;
  readonly sharedStrings: readonly string[];
  readonly styles: Styles;
  readonly dateSystem: DateSystem;
  readonly conformance: Conformance;
  /** Number of cells still allowed by `maxCells`, shared by all sheets of a workbook. */
  readonly budget: { remaining: number };
  /** Number formats already classified, by number format id; shared by all sheets. */
  readonly formatCache: Map<number, NumberFormat>;
}

/** A cell's number format and whether it displays dates or times. */
export interface NumberFormat {
  readonly code: string | null;
  readonly isDate: boolean;
}

/**
 * Format codes longer than this are not classified (and so never treated as dates); this
 * bounds the work an oversized code can cause. Implementation choice, see docs/decisions.md D-115.
 */
const MAX_FORMAT_CODE_LENGTH = 1024;

/** The cells and ranges of one worksheet. */
export interface WorksheetData {
  /** Row number → (column number → cell). */
  readonly rows: Map<number, Map<number, Cell>>;
  readonly dimension: string | null;
  readonly merges: string[];
  readonly cellCount: number;
  /** `col` elements in file order (W2-03). */
  readonly columns: ColumnInfo[];
  /** Rows with a height or hidden flag (W2-04). */
  readonly rowInfo: Map<number, RowInfo>;
  /** Frozen panes of the first sheet view (W2-05). */
  readonly freeze: FrozenPanes | null;
  /** `autoFilter@ref` (W2-06). */
  readonly autoFilter: string | null;
}

const FORMULA_KINDS: readonly FormulaKind[] = ['normal', 'shared', 'array', 'dataTable'];

interface PendingCell {
  row: number;
  col: number;
  ref: string;
  t: string;
  styleIndex: number;
  v: string | null;
  inline: string | null;
  formulaAttrs: ReadonlyMap<string, string> | null;
  formulaText: string;
}

/** Parses a non-negative integer of at most `maxDigits` digits; `null` otherwise. */
function parseUint(text: string, maxDigits: number): number | null {
  const s = text.trim();
  if (s.length === 0 || s.length > maxDigits) return null;
  let n = 0;
  for (let i = 0; i < s.length; i++) {
    const d = s.charCodeAt(i) - 48;
    if (d < 0 || d > 9) return null;
    n = n * 10 + d;
  }
  return n;
}

/**
 * Parses the lexical form of `xsd:double` restricted to finite values: optional sign, digits
 * with an optional fraction, optional exponent. Returns `null` for anything else (including
 * INF and NaN, which have no JSON representation).
 */
export function parseXsdDouble(text: string): number | null {
  const s = text.trim();
  if (s.length === 0 || s.length > 64) return null;
  let i = 0;
  if (s[i] === '+' || s[i] === '-') i++;
  let digits = 0;
  while (i < s.length && s.charCodeAt(i) >= 48 && s.charCodeAt(i) <= 57) {
    i++;
    digits++;
  }
  if (s[i] === '.') {
    i++;
    while (i < s.length && s.charCodeAt(i) >= 48 && s.charCodeAt(i) <= 57) {
      i++;
      digits++;
    }
  }
  if (digits === 0) return null;
  if (s[i] === 'e' || s[i] === 'E') {
    i++;
    if (s[i] === '+' || s[i] === '-') i++;
    const start = i;
    while (i < s.length && s.charCodeAt(i) >= 48 && s.charCodeAt(i) <= 57) i++;
    if (i === start) return null;
  }
  if (i !== s.length) return null;
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

/** Parses a worksheet part. */
export function parseWorksheet(xml: string, ctx: WorksheetContext): WorksheetData {
  const { handler, finish } = createWorksheetParser(ctx);
  parseXml(xml, handler, { part: ctx.part, limits: ctx.limits });
  return finish();
}

/**
 * The worksheet parser as an XML handler, for whole-text and streaming parsing. With `onRow`
 * (streaming, docs/decisions.md D-402), each row's cells are passed on when its `</row>` is
 * read and then dropped, so memory does not grow with the sheet; rows come in file order, and
 * shared formulas whose master comes later in the file cannot be expanded.
 */
export function createWorksheetParser(
  ctx: WorksheetContext,
  onRow?: (row: number, cells: readonly Cell[], info: RowInfo | undefined) => void,
): { handler: XmlHandler; finish(): WorksheetData } {
  const rows = new Map<number, Map<number, Cell>>();
  const merges: string[] = [];
  const formats = ctx.formatCache;
  const path: string[] = [];
  let ns: string | null = null;
  let dimension: string | null = null;
  let cellCount = 0;
  const masters = new Map<number, SharedMaster>();
  const dependents: SharedDependent[] = [];
  const columns: ColumnInfo[] = [];
  const rowInfo = new Map<number, RowInfo>();
  let freeze: FrozenPanes | null = null;
  let autoFilter: string | null = null;
  let sheetViews = 0;
  let currentRow = 0;
  let lastCol = 0;
  let cell: PendingCell | null = null;
  /** Where character data currently goes. */
  let target: 'v' | 'f' | 'is' | null = null;
  let phoneticDepth = 0;
  /** Raw text of the inline-string `t` element being read. */
  let inlineText = '';

  const loc = (ref?: string) => (ref === undefined ? { part: ctx.part, sheet: ctx.sheet } : { part: ctx.part, sheet: ctx.sheet, cell: ref });

  function invalidRef(text: string, what: string): never {
    throw new XlvError('XLV_E_INVALID_CELL_REF', `${what} "${text}" is not a valid reference inside the sheet grid.`, {
      location: { part: ctx.part, sheet: ctx.sheet },
      hint: `References must lie between A1 and ${formatCellRef(MAX_ROWS, MAX_COLUMNS)}. The file is damaged or was written by a faulty program; re-save it in a spreadsheet application.`,
      details: { ref: text },
    });
  }

  // Keyed by number format id and shared by all sheets, so each format code is classified
  // once per workbook however many style indexes or sheets refer to it.
  const numberFormat = (styleIndex: number): NumberFormat => {
    const id = ctx.styles.numFmtId(styleIndex);
    let format = formats.get(id);
    if (format === undefined) {
      const code = ctx.styles.customFormatCode(id) ?? BUILTIN_NUMBER_FORMATS.get(id) ?? null;
      let isDate: boolean;
      if (code === null) isDate = isDateFormatId(id);
      else if (code.length > MAX_FORMAT_CODE_LENGTH) {
        isDate = false;
        ctx.warnings.add(
          'XLV_W_INVALID_VALUE',
          `Number format ${String(id)} is ${String(code.length)} characters long; it is not checked for dates.`,
          { part: ctx.part, sheet: ctx.sheet },
        );
      } else isDate = isDateFormat(code);
      format = { code: id === 0 || code === 'General' ? null : code, isDate };
      formats.set(id, format);
    }
    return format;
  };

  const finishCell = (p: PendingCell): void => {
    const warn = (message: string): void => ctx.warnings.add('XLV_W_INVALID_VALUE', message, loc(p.ref));
    const format = numberFormat(p.styleIndex);
    let type: CellType = 'empty';
    let value: CellValue = null;
    const raw = p.v;
    switch (p.t) {
      case 'n': {
        if (raw === null) break;
        const n = parseXsdDouble(raw);
        if (n === null) {
          warn(`A numeric cell holds "${raw.slice(0, 40)}", which is not a number; it is returned as text.`);
          type = 'string';
          value = raw;
        } else if (format.isDate) {
          const date = serialToDate(n, ctx.dateSystem, ctx.conformance);
          if (date === null) {
            warn(`The date serial ${String(n)} is outside the range of the ${ctx.dateSystem} date system; it is returned as a number.`);
            type = 'number';
            value = n;
          } else {
            type = 'date';
            value = date;
          }
        } else {
          type = 'number';
          value = n;
        }
        break;
      }
      case 's': {
        if (raw === null) break;
        const index = parseUint(raw, 10);
        const text = index === null ? undefined : ctx.sharedStrings[index];
        if (text === undefined) warn(`Shared string index "${raw.slice(0, 40)}" does not exist; the cell is treated as empty.`);
        else {
          type = 'string';
          value = text;
        }
        break;
      }
      case 'str':
        if (raw !== null) {
          type = 'string';
          value = decodeXstring(raw);
        }
        break;
      case 'inlineStr':
        break;
      case 'b': {
        if (raw === null) break;
        const b = parseXsdBoolean(raw);
        if (b === null) {
          warn(`A boolean cell holds "${raw.slice(0, 40)}"; it is returned as text.`);
          type = 'string';
          value = raw;
        } else {
          type = 'boolean';
          value = b;
        }
        break;
      }
      case 'e':
        if (raw !== null) {
          type = 'error';
          value = raw.trim();
        }
        break;
      case 'd': {
        if (raw === null) break;
        const date = parseIsoDateTime(raw.trim());
        if (date === null) {
          warn(`A date cell holds "${raw.slice(0, 40)}", which is not an ISO 8601 date; it is returned as text.`);
          type = 'string';
          value = raw;
        } else {
          type = 'date';
          value = date;
        }
        break;
      }
      default:
        warn(`The cell type "${p.t.slice(0, 20)}" is unknown; the value is returned as text.`);
        if (raw !== null) {
          type = 'string';
          value = raw;
        }
    }
    // Inline strings: `is` holds the value. Accept `is` without t="inlineStr" too (E-22).
    if (value === null && p.inline !== null) {
      type = 'string';
      value = p.inline;
    }

    let formula: Formula | null = null;
    if (p.formulaAttrs !== null) {
      const rawKind = p.formulaAttrs.get('t') ?? 'normal';
      let kind = FORMULA_KINDS.find((k) => k === rawKind) ?? 'normal';
      const si = p.formulaAttrs.get('si');
      const sharedIndex = si === undefined ? null : parseUint(si, 10);
      const ref = p.formulaAttrs.get('ref') ?? null;
      // A formula is shared only when both t="shared" and si are present (E-28).
      if (kind === 'shared' && sharedIndex === null) kind = 'normal';
      // ST_Formula is an ST_Xstring (W2-02), so _xHHHH_ escapes are decoded.
      let text: string | null = decodeXstring(p.formulaText);
      if (kind === 'shared' && sharedIndex !== null) {
        if (ref === null) {
          // A dependent cell of a shared formula: its own text, if any, is ignored (E-28). The
          // text is derived from the master after the whole sheet is read (masters may come later).
          text = null;
          dependents.push({ row: p.row, col: p.col, ref: p.ref, si: sharedIndex });
        } else if (!masters.has(sharedIndex)) {
          // The first formula of a group is the master (E-28); later ones with ref are ignored.
          masters.set(sharedIndex, { row: p.row, col: p.col, text, range: parseRange(ref) });
        }
      }
      formula = { kind, text, ref, sharedIndex: kind === 'shared' ? sharedIndex : null };
      if (p.v === null && p.inline === null) {
        ctx.warnings.add(
          'XLV_W_FORMULA_NO_CACHED_VALUE',
          'Formula cells without a cached value are returned with value null; xlvane does not calculate formulas.',
          loc(p.ref),
        );
      }
    }
    if (value === null && formula === null) return;

    let rowCells = rows.get(p.row);
    if (rowCells === undefined) {
      rowCells = new Map();
      rows.set(p.row, rowCells);
    }
    if (rowCells.has(p.col)) {
      ctx.warnings.add('XLV_W_DUPLICATE_CELL', 'The sheet defines the same cell more than once; the last definition is used.', loc(p.ref));
    } else {
      if (--ctx.budget.remaining < 0) {
        throw new XlvError('XLV_E_LIMIT_EXCEEDED', 'The workbook has more cells than the limit.', {
          location: loc(p.ref),
          hint: `If you trust this file, raise the limit, for example { limits: { maxCells: ${String(ctx.limits.maxCells * 4)} } }.`,
          details: { limit: 'maxCells', value: ctx.limits.maxCells },
        });
      }
      cellCount++;
    }
    rowCells.set(p.col, {
      ref: p.ref,
      row: p.row,
      col: p.col,
      type,
      value,
      formula,
      styleIndex: p.styleIndex,
      numFmt: format.code,
      style: ctx.styles.cellStyle(p.styleIndex, format.code),
    });
  };

  const warnNotExpanded = (ref: string): void =>
    ctx.warnings.add(
      'XLV_W_SHARED_FORMULA_NOT_EXPANDED',
      'Some cells reuse a shared formula that could not be moved to them (missing master, cell outside the master range, or formula text xlvane cannot translate safely); their formula.text is null and their cached values are available.',
      loc(ref),
    );
  /** Streaming: expands the pending shared formulas and hands the buffered cells on. */
  const flushRows = (): void => {
    if (!onRow) return;
    // Row heights are handed on with the row instead of being collected (constant memory).
    const info = rowInfo.get(currentRow);
    rowInfo.delete(currentRow);
    if (rows.size === 0) return;
    expandSharedFormulas(rows, masters, dependents, warnNotExpanded);
    dependents.length = 0;
    const out: Cell[] = [];
    for (const r of [...rows.keys()].sort((a, b) => a - b)) {
      const cells = rows.get(r);
      if (cells) for (const c of [...cells.keys()].sort((a, b) => a - b)) out.push(cells.get(c) as Cell);
    }
    rows.clear();
    if (out.length > 0) onRow(currentRow, out, info);
  };

  const handler: XmlHandler = {
      start(elementNs, local, attrs) {
        if (ns === null) {
          if (local !== 'worksheet' || (elementNs !== SML_NS.strict && elementNs !== SML_NS.transitional)) {
            throw new XlvError('XLV_E_PART_INVALID', `The worksheet part has the root element <${local}>.`, {
              location: { part: ctx.part, sheet: ctx.sheet },
              hint: 'The sheet part is not a SpreadsheetML worksheet. Re-save the file in a spreadsheet application.',
              details: { element: local, namespace: elementNs },
            });
          }
          ns = elementNs;
        }
        const parent = path[path.length - 1];
        path.push(elementNs === ns ? local : '');
        if (elementNs !== ns) return;
        const depth = path.length;
        if (cell !== null) {
          if (phoneticDepth > 0 || local === 'rPh') {
            phoneticDepth++;
            return;
          }
          if (depth === 5 && local === 'v') target = 'v';
          else if (depth === 5 && local === 'f') {
            target = 'f';
            cell.formulaAttrs = attrs;
          } else if (depth === 5 && local === 'is') cell.inline ??= '';
          else if (local === 't' && cell.inline !== null && (parent === 'is' || parent === 'r')) target = 'is';
          return;
        }
        if (depth === 2 && local === 'dimension') {
          const ref = attrs.get('ref') ?? '';
          const range = parseRange(ref);
          if (range === null) ctx.warnings.add('XLV_W_INVALID_VALUE', `The sheet dimension "${ref.slice(0, 40)}" is not a valid range.`, loc());
          else dimension = formatRange(range);
        } else if (depth === 3 && local === 'mergeCell' && parent === 'mergeCells') {
          const ref = attrs.get('ref') ?? '';
          const range = parseRange(ref);
          if (range === null) ctx.warnings.add('XLV_W_INVALID_VALUE', `The merged range "${ref.slice(0, 40)}" is not a valid range.`, loc());
          else merges.push(formatRange(range));
        } else if (depth === 2 && local === 'autoFilter') {
          const ref = attrs.get('ref');
          const range = ref === undefined ? null : parseRange(ref);
          if (range === null) {
            if (ref !== undefined) ctx.warnings.add('XLV_W_INVALID_VALUE', `The autofilter range "${ref.slice(0, 40)}" is not a valid range.`, loc());
          } else autoFilter = formatRange(range);
        } else if (depth === 3 && local === 'sheetView' && parent === 'sheetViews') {
          sheetViews++;
        } else if (depth === 4 && local === 'pane' && parent === 'sheetView' && sheetViews === 1) {
          freeze = parsePane(attrs, (message) => ctx.warnings.add('XLV_W_INVALID_VALUE', message, loc()));
        } else if (depth === 3 && local === 'col' && parent === 'cols') {
          const col = parseColumn(attrs);
          if (col === null) ctx.warnings.add('XLV_W_INVALID_VALUE', 'A column definition (col) has invalid min, max or width attributes; it is ignored.', loc());
          else columns.push(col);
        } else if (depth === 3 && local === 'row' && parent === 'sheetData') {
          const r = attrs.get('r');
          if (r === undefined) currentRow++;
          else {
            const n = parseUint(r, 7);
            if (n === null || n < 1 || n > MAX_ROWS) invalidRef(r, 'Row number');
            currentRow = n;
          }
          if (currentRow > MAX_ROWS) invalidRef(String(currentRow), 'Row number');
          lastCol = 0;
          const ht = attrs.get('ht');
          const height = ht === undefined ? null : parseXsdDouble(ht);
          const hiddenRaw = attrs.get('hidden');
          const hidden = hiddenRaw === undefined ? false : parseXsdBoolean(hiddenRaw) === true;
          if (ht !== undefined && (height === null || height < 0)) {
            ctx.warnings.add('XLV_W_INVALID_VALUE', `The row height "${ht.slice(0, 20)}" is not a valid number; it is ignored.`, loc());
          }
          const validHeight = height !== null && height >= 0 ? height : null;
          if (validHeight !== null || hidden) rowInfo.set(currentRow, { height: validHeight, hidden });
        } else if (depth === 4 && local === 'c' && parent === 'row' && path[1] === 'sheetData') {
          const r = attrs.get('r');
          let row = currentRow;
          let col: number;
          if (r === undefined) {
            col = lastCol + 1;
            if (col > MAX_COLUMNS || row < 1) invalidRef(formatCellRefSafe(row, col), 'Cell');
          } else {
            const address = parseCellRef(r);
            if (address === null) return invalidRef(r, 'Cell');
            row = address.row;
            col = address.col;
          }
          lastCol = col;
          const s = attrs.get('s');
          let styleIndex = 0;
          if (s !== undefined) {
            const parsed = parseUint(s, 10);
            if (parsed === null) ctx.warnings.add('XLV_W_INVALID_VALUE', `The style index "${s.slice(0, 20)}" is not a number; using 0.`, loc());
            else styleIndex = parsed;
          }
          cell = {
            row,
            col,
            ref: formatCellRef(row, col),
            t: attrs.get('t') ?? 'n',
            styleIndex,
            v: null,
            inline: null,
            formulaAttrs: null,
            formulaText: '',
          };
        }
      },
      end(elementNs, local) {
        path.pop();
        if (onRow && local === 'row' && elementNs === ns && path.length === 2) flushRows();
        if (elementNs !== ns || cell === null) return;
        if (phoneticDepth > 0) {
          phoneticDepth--;
          return;
        }
        if (local === 'c' && path.length === 3) {
          const done = cell;
          cell = null;
          target = null;
          finishCell(done);
        } else if (local === 't' && target === 'is') {
          // Each `t` is decoded as a whole: an _xHHHH_ escape cannot span two text events.
          cell.inline = (cell.inline ?? '') + decodeXstring(inlineText);
          inlineText = '';
          target = null;
        } else if ((local === 'v' || local === 'f') && path.length === 4) {
          target = null;
        }
      },
      text(value) {
        if (cell === null || target === null) return;
        if (target === 'v') cell.v = (cell.v ?? '') + value;
        else if (target === 'f') cell.formulaText += value;
        else inlineText += value;
      },
    };
  const finish = (): WorksheetData => {
    if (onRow) flushRows();
    else expandSharedFormulas(rows, masters, dependents, warnNotExpanded);
    return { rows, dimension, merges, cellCount, columns, rowInfo, freeze, autoFilter };
  };
  return { handler, finish };
}

/** The master cell of a shared formula group (E-28). */
interface SharedMaster {
  readonly row: number;
  readonly col: number;
  readonly text: string;
  readonly range: CellRange | null;
}

/** A cell that reuses a shared formula. */
interface SharedDependent {
  readonly row: number;
  readonly col: number;
  readonly ref: string;
  readonly si: number;
}

/**
 * Gives every dependent cell of a shared formula the master formula moved to its position
 * (docs/decisions.md D-107, docs/facts/formulas.md). Each master is tokenized once; a dependent
 * gets a formula whose `text` is rendered on first access, so a long master formula reused by
 * many cells costs memory only for the texts actually read. Dependents that cannot be expanded
 * keep `text: null` and are reported through `warn`.
 */
function expandSharedFormulas(
  rows: Map<number, Map<number, Cell>>,
  masters: ReadonlyMap<number, SharedMaster>,
  dependents: readonly SharedDependent[],
  warn: (ref: string) => void,
): void {
  const compiled = new Map<number, CompiledFormula | null>();
  for (const d of dependents) {
    const rowCells = rows.get(d.row);
    const cell = rowCells?.get(d.col);
    // Skip entries replaced by a later definition of the same cell.
    if (rowCells === undefined || cell === undefined || cell.formula?.kind !== 'shared' || cell.formula.ref !== null || cell.formula.sharedIndex !== d.si) {
      continue;
    }
    const master = masters.get(d.si);
    const range = master?.range;
    // Cells outside the master's range are implementation-defined (E-28); they are not expanded.
    if (master === undefined || range == null || d.row < range.start.row || d.row > range.end.row || d.col < range.start.col || d.col > range.end.col) {
      warn(d.ref);
      continue;
    }
    let formula = compiled.get(d.si);
    if (formula === undefined) {
      formula = compileFormula(master.text);
      compiled.set(d.si, formula);
    }
    const dRow = d.row - master.row;
    const dCol = d.col - master.col;
    if (formula === null || !formula.canShift(dRow, dCol)) {
      warn(d.ref);
      continue;
    }
    rowCells.set(d.col, { ...cell, formula: lazySharedFormula(formula, dRow, dCol, d.si) });
  }
}

function lazySharedFormula(compiled: CompiledFormula, dRow: number, dCol: number, si: number): Formula {
  let text: string | undefined;
  return {
    kind: 'shared',
    get text(): string {
      text ??= compiled.render(dRow, dCol);
      return text;
    },
    ref: null,
    sharedIndex: si,
  };
}

/** Parses a `col` element (CT_Col, W2-03); `null` if min/max are missing or out of range. */
function parseColumn(attrs: ReadonlyMap<string, string>): ColumnInfo | null {
  const min = parseUint(attrs.get('min') ?? '', 5);
  const max = parseUint(attrs.get('max') ?? '', 5);
  if (min === null || max === null || min < 1 || max < min || max > MAX_COLUMNS) return null;
  const rawWidth = attrs.get('width');
  const width = rawWidth === undefined ? null : parseXsdDouble(rawWidth);
  if (rawWidth !== undefined && (width === null || width < 0)) return null;
  const hidden = attrs.get('hidden');
  return { min, max, width, hidden: hidden !== undefined && parseXsdBoolean(hidden) === true };
}

/**
 * Reads frozen panes from a `pane` element (CT_Pane, W2-05): only the states `frozen` and
 * `frozenSplit` freeze rows and columns; xSplit and ySplit then count the frozen columns and rows.
 */
function parsePane(attrs: ReadonlyMap<string, string>, warn: (message: string) => void): FrozenPanes | null {
  const state = attrs.get('state') ?? 'split';
  if (state !== 'frozen' && state !== 'frozenSplit') return null;
  const count = (name: string, max: number): number | null => {
    const raw = attrs.get(name);
    if (raw === undefined) return 0;
    const n = parseXsdDouble(raw);
    if (n === null || !Number.isInteger(n) || n < 0 || n >= max) {
      warn(`The frozen pane attribute ${name}="${raw.slice(0, 20)}" is not a valid count; the frozen panes are ignored.`);
      return null;
    }
    return n;
  };
  const cols = count('xSplit', MAX_COLUMNS);
  const rows = count('ySplit', MAX_ROWS);
  if (cols === null || rows === null || (rows === 0 && cols === 0)) return null;
  return { rows, cols };
}

function formatCellRefSafe(row: number, col: number): string {
  return row >= 1 && row <= MAX_ROWS && col >= 1 && col <= MAX_COLUMNS ? formatCellRef(row, col) : `R${row}C${col}`;
}
