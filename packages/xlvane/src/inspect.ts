import { XlvError, type XlvWarning } from './errors';
import type { ReadLimits } from './limits';
import type { Cell, CellType, Conformance, DateSystem, Sheet, SheetKind, SheetState, Workbook, WorkbookFormat } from './model';
import { readWorkbook } from './read';
import { columnLetters, formatRange, parseRange, type CellRange } from './sml/cell-ref';
import { suggest } from './util/suggest';
import { WarningCollector } from './util/warnings';

/**
 * Options for {@link inspect}.
 *
 * @example
 * ```ts
 * import { inspect, type InspectOptions } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const options: InspectOptions = { headerScanRows: 50, maxColumns: 20 };
 * console.log(inspect(bytes, options).sheets.length);
 * ```
 */
export interface InspectOptions {
  /** Resource limits used when `input` is a file; see {@link ReadLimits}. */
  readonly limits?: Partial<ReadLimits>;
  /** Maximum number of columns described per sheet. Default 200. */
  readonly maxColumns?: number;
  /** Number of leading non-empty rows searched for a header row. Default 20. */
  readonly headerScanRows?: number;
  /** Maximum number of example formulas per sheet. Default 10. */
  readonly formulaSamples?: number;
}

/**
 * A JSON-safe cell value: dates are ISO 8601 strings without a time zone.
 *
 * @example
 * ```ts
 * import { inspect, type InspectValue } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const samples: readonly InspectValue[] = inspect(bytes).sheets[0]?.columns[0]?.samples ?? [];
 * console.log(samples.map((value) => typeof value));
 * ```
 */
export type InspectValue = string | number | boolean;

/**
 * Counts of cell types in a column.
 *
 * @example
 * ```ts
 * import { inspect, type InspectTypeCounts } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const counts: InspectTypeCounts | undefined = inspect(bytes).sheets[0]?.columns[0]?.types;
 * console.log(counts?.number, counts?.string, counts?.empty);
 * ```
 */
export interface InspectTypeCounts {
  readonly number: number;
  readonly string: number;
  readonly boolean: number;
  readonly date: number;
  readonly error: number;
  readonly empty: number;
}

/**
 * Description of one column of the used range.
 *
 * @example
 * ```ts
 * import { inspect, type InspectColumn } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const columns: readonly InspectColumn[] = inspect(bytes).sheets[0]?.columns ?? [];
 * for (const column of columns) console.log(column.letter, column.header, column.dominantType);
 * ```
 */
export interface InspectColumn {
  /** 1-based column number. */
  readonly index: number;
  readonly letter: string;
  /** Text of the header cell, or `null` when no header row was detected. */
  readonly header: string | null;
  /** The most frequent non-empty type in the data rows, or `"empty"`. */
  readonly dominantType: CellType;
  /** Type counts over the data rows (below the header row, within the used range). */
  readonly types: InspectTypeCounts;
  /** Up to three distinct example values from the data rows. */
  readonly samples: readonly InspectValue[];
}

/**
 * Description of one sheet.
 *
 * @example
 * ```ts
 * import { inspect, type InspectSheet } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const hidden: InspectSheet[] = inspect(bytes).sheets.filter((sheet) => sheet.state !== 'visible');
 * console.log(hidden.map((sheet) => sheet.name));
 * ```
 */
export interface InspectSheet {
  readonly name: string;
  readonly index: number;
  readonly kind: SheetKind;
  readonly state: SheetState;
  /** Range declared in the file, or `null`. */
  readonly dimension: string | null;
  /** Smallest range containing every cell with a value or formula, or `null` for an empty sheet. */
  readonly usedRange: string | null;
  /** Rows in the used range. */
  readonly rowCount: number;
  /** Columns in the used range. */
  readonly columnCount: number;
  /** Cells with a value or formula. */
  readonly cellCount: number;
  readonly formulaCount: number;
  readonly mergedRanges: readonly string[];
  /** Detected header row (1-based) and how sure the detection is, or `null`. */
  readonly headerRow: { readonly row: number; readonly confidence: 'high' | 'low' } | null;
  readonly columns: readonly InspectColumn[];
  /** `true` when the used range has more columns than `maxColumns`. */
  readonly columnsTruncated: boolean;
  /** Example formulas (text `null` only for shared formulas that could not be translated). */
  readonly formulas: readonly { readonly cell: string; readonly formula: string | null }[];
  /** Cells whose text has formatted runs (`Cell.richText`). */
  readonly richTextCellCount: number;
  /** Up to ten of those cells, e.g. `["A1", "B4"]`. */
  readonly richTextCells: readonly string[];
  /** Columns with a default style (`ColumnInfo.style`), e.g. `["B", "D:F"]`; at most `maxColumns` entries. */
  readonly styledColumns: readonly string[];
  /** Number of hyperlinks (`Sheet.hyperlinks`). */
  readonly hyperlinkCount: number;
  /** Up to ten hyperlinks: range, external target and place in the workbook. Targets are never opened. */
  readonly hyperlinks: readonly { readonly ref: string; readonly target: string | null; readonly location: string | null }[];
  /** Number of notes (`Sheet.notes`). */
  readonly noteCount: number;
  /** Up to ten cells that have a note. */
  readonly notes: readonly string[];
}

/**
 * A defined name as reported by {@link inspect}.
 *
 * @example
 * ```ts
 * import { inspect, type InspectDefinedName } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const names: readonly InspectDefinedName[] = inspect(bytes).definedNames;
 * console.log(names.map((n) => `${n.name} = ${n.refersTo}`));
 * ```
 */
export interface InspectDefinedName {
  readonly name: string;
  /** Formula text without "=", cut to 100 characters. */
  readonly refersTo: string;
  /** The sheet the name belongs to, or `null` for a workbook-level name. */
  readonly sheet: string | null;
  readonly hidden: boolean;
}

/**
 * The result of {@link inspect}. Version 1 of the schema documented in docs/inspect-json.md.
 *
 * @example
 * ```ts
 * import { inspect, type InspectReport } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const report: InspectReport = inspect(bytes);
 * console.log(report.schemaVersion, report.sheetCount, report.warnings.length);
 * ```
 */
export interface InspectReport {
  readonly schemaVersion: 1;
  readonly format: WorkbookFormat;
  readonly conformance: Conformance;
  readonly dateSystem: DateSystem;
  readonly hasMacros: boolean;
  readonly externalLinkCount: number;
  /**
   * Document properties that are set (`Workbook.properties`); dates as ISO 8601 strings in UTC.
   * Empty when the file has none.
   */
  readonly properties: Readonly<Record<string, string>>;
  /** Defined names in file order, at most 100. */
  readonly definedNames: readonly InspectDefinedName[];
  /** `true` when the workbook has more than 100 defined names. */
  readonly definedNamesTruncated: boolean;
  readonly sheetCount: number;
  readonly sheets: readonly InspectSheet[];
  readonly warnings: readonly XlvWarning[];
}

const OPTION_NAMES = ['limits', 'maxColumns', 'headerScanRows', 'formulaSamples'] as const;
/** Upper bound on the steps spent looking for values hidden under merged ranges, per sheet. */
const MAX_MERGE_CHECK_WORK = 5_000_000;
/** Bounds of the lists that are not controlled by an option. */
const MAX_DEFINED_NAMES = 100;
const MAX_RICH_TEXT_CELLS = 10;
const MAX_HYPERLINK_SAMPLES = 10;
const TYPE_ORDER: readonly Exclude<CellType, 'empty'>[] = ['number', 'string', 'date', 'boolean', 'error'];

interface ResolvedOptions {
  readonly maxColumns: number;
  readonly headerScanRows: number;
  readonly formulaSamples: number;
}

function resolveOptions(options: unknown): { limits: Partial<ReadLimits> | undefined; resolved: ResolvedOptions } {
  if (options === undefined) return { limits: undefined, resolved: { maxColumns: 200, headerScanRows: 20, formulaSamples: 10 } };
  if (options === null || typeof options !== 'object' || Array.isArray(options)) {
    throw new XlvError('XLV_E_INVALID_OPTION', 'Options must be an object.', {
      hint: 'Pass an object such as { maxColumns: 50 }, or leave the options out.',
    });
  }
  const record = options as Record<string, unknown>;
  for (const key of Object.keys(record)) {
    if (!(OPTION_NAMES as readonly string[]).includes(key)) {
      const close = suggest(key, OPTION_NAMES);
      throw new XlvError('XLV_E_INVALID_OPTION', `Unknown option "${key}".`, {
        hint: `${close.length > 0 ? `Did you mean "${close[0] ?? ''}"? ` : ''}Known options: ${OPTION_NAMES.join(', ')}.`,
        details: { option: key, known: [...OPTION_NAMES] },
      });
    }
  }
  const count = (name: 'maxColumns' | 'headerScanRows' | 'formulaSamples', fallback: number, min: number): number => {
    if (!Object.hasOwn(record, name)) return fallback;
    const value = record[name];
    if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < min) {
      throw new XlvError('XLV_E_INVALID_OPTION', `Option "${name}" must be a whole number of at least ${String(min)}.`, {
        hint: `Pass for example { ${name}: ${String(fallback)} }, or leave it out to use the default.`,
        details: { option: name },
      });
    }
    return value;
  };
  return {
    limits: Object.hasOwn(record, 'limits') ? (record['limits'] as Partial<ReadLimits>) : undefined,
    resolved: {
      maxColumns: count('maxColumns', 200, 1),
      headerScanRows: count('headerScanRows', 20, 1),
      formulaSamples: count('formulaSamples', 10, 0),
    },
  };
}

function isWorkbook(input: unknown): input is Workbook {
  return (
    input !== null &&
    typeof input === 'object' &&
    !ArrayBuffer.isView(input) &&
    !(input instanceof ArrayBuffer) &&
    typeof (input as Workbook).getSheet === 'function' &&
    Array.isArray((input as Workbook).sheets)
  );
}

/** Formats a value for JSON: dates as `YYYY-MM-DDTHH:mm:ss.sss` (no time zone). */
function jsonValue(value: string | number | boolean | Date): InspectValue {
  if (value instanceof Date) return value.toISOString().slice(0, 23);
  if (typeof value === 'string' && value.length > 100) return `${value.slice(0, 100)}…`;
  return value;
}

/**
 * Describes the structure of a workbook as plain JSON data: sheets and their visibility, used
 * ranges, detected header rows, column types with example values, merged cells, formulas and
 * warnings. Accepts the file contents or a workbook returned by `readWorkbook`.
 *
 * @example
 * ```ts
 * import { inspect } from 'xlvane';
 *
 * declare const bytes: Uint8Array; // the contents of an .xlsx file
 * const report = inspect(bytes);
 * for (const sheet of report.sheets) {
 *   console.log(sheet.name, sheet.usedRange, sheet.columns.map((c) => `${c.header}: ${c.dominantType}`));
 * }
 * console.log(JSON.stringify(report, null, 2));
 * ```
 */
export function inspect(input: Uint8Array | ArrayBuffer | Workbook, options?: InspectOptions): InspectReport {
  const { limits, resolved } = resolveOptions(options);
  const workbook = isWorkbook(input) ? input : readWorkbook(input, limits === undefined ? {} : { limits });
  const warnings = new WarningCollector();
  const sheets = workbook.sheets.map((sheet) => inspectSheet(sheet, resolved, warnings));
  return {
    schemaVersion: 1,
    format: workbook.format,
    conformance: workbook.conformance,
    dateSystem: workbook.dateSystem,
    hasMacros: workbook.hasMacros,
    externalLinkCount: workbook.externalLinkCount,
    properties: propertiesJson(workbook),
    definedNames: workbook.definedNames.slice(0, MAX_DEFINED_NAMES).map((n) => ({
      name: n.name,
      refersTo: n.refersTo.length > 100 ? `${n.refersTo.slice(0, 100)}…` : n.refersTo,
      sheet: n.sheet,
      hidden: n.hidden,
    })),
    definedNamesTruncated: workbook.definedNames.length > MAX_DEFINED_NAMES,
    sheetCount: workbook.sheets.length,
    sheets,
    warnings: [...workbook.warnings, ...warnings.list()],
  };
}

function propertiesJson(workbook: Workbook): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [key, value] of Object.entries(workbook.properties)) {
    if (value instanceof Date) out[key] = value.toISOString();
    else if (typeof value === 'string') out[key] = value;
  }
  return out;
}

/** Column ranges (`"B"`, `"D:F"`) that have a default style. */
function styledColumns(sheet: Sheet, max: number): string[] {
  const out: string[] = [];
  for (const c of sheet.columns) {
    if (c.style === undefined || out.length >= max) continue;
    out.push(c.min === c.max ? columnLetters(c.min) : `${columnLetters(c.min)}:${columnLetters(c.max)}`);
  }
  return out;
}

function inspectSheet(sheet: Sheet, options: ResolvedOptions, warnings: WarningCollector): InspectSheet {
  const rows = [...sheet.rows()];
  let minRow = Number.POSITIVE_INFINITY;
  let maxRow = 0;
  let minCol = Number.POSITIVE_INFINITY;
  let maxCol = 0;
  let formulaCount = 0;
  let richTextCellCount = 0;
  const richTextCells: string[] = [];
  const formulas: { cell: string; formula: string | null }[] = [];
  for (const cells of rows) {
    for (const cell of cells) {
      minRow = Math.min(minRow, cell.row);
      maxRow = Math.max(maxRow, cell.row);
      minCol = Math.min(minCol, cell.col);
      maxCol = Math.max(maxCol, cell.col);
      if (cell.richText !== undefined) {
        richTextCellCount++;
        if (richTextCells.length < MAX_RICH_TEXT_CELLS) richTextCells.push(cell.ref);
      }
      if (cell.formula !== null) {
        formulaCount++;
        if (formulas.length < options.formulaSamples) formulas.push({ cell: cell.ref, formula: cell.formula.text });
      }
    }
  }
  const empty = maxRow === 0;
  const usedRange = empty ? null : formatRange({ start: { row: minRow, col: minCol }, end: { row: maxRow, col: maxCol } });

  checkDimension(sheet, empty ? null : { minRow, maxRow, minCol, maxCol }, warnings);
  checkMergedValues(sheet, rows, warnings);

  const base = {
    name: sheet.name,
    index: sheet.index,
    kind: sheet.kind,
    state: sheet.state,
    dimension: sheet.dimension,
    usedRange,
    cellCount: sheet.cellCount,
    formulaCount,
    mergedRanges: [...sheet.merges],
    formulas,
    richTextCellCount,
    richTextCells,
    styledColumns: styledColumns(sheet, options.maxColumns),
    hyperlinkCount: sheet.hyperlinks.length,
    hyperlinks: sheet.hyperlinks.slice(0, MAX_HYPERLINK_SAMPLES).map((l) => ({ ref: l.ref, target: l.target, location: l.location })),
    noteCount: sheet.notes.length,
    notes: sheet.notes.slice(0, MAX_HYPERLINK_SAMPLES).map((n) => n.ref),
  };
  if (empty) {
    return { ...base, rowCount: 0, columnCount: 0, headerRow: null, columns: [], columnsTruncated: false };
  }

  const width = maxCol - minCol + 1;
  const headerRow = detectHeader(rows, width, options.headerScanRows);
  const headerCells = headerRow === null ? undefined : rows.find((r) => r[0]?.row === headerRow.row);
  const firstDataRow = headerRow === null ? minRow : headerRow.row + 1;
  const dataRowCount = Math.max(0, maxRow - firstDataRow + 1);
  const shownColumns = Math.min(width, options.maxColumns);

  const stats = Array.from({ length: shownColumns }, () => ({
    counts: { number: 0, string: 0, boolean: 0, date: 0, error: 0 },
    samples: [] as InspectValue[],
  }));
  for (const cells of rows) {
    if (cells.length === 0 || (cells[0]?.row ?? 0) < firstDataRow) continue;
    for (const cell of cells) {
      const stat = stats[cell.col - minCol];
      if (stat === undefined || cell.type === 'empty' || cell.value === null) continue;
      stat.counts[cell.type]++;
      const sample = jsonValue(cell.value);
      if (stat.samples.length < 3 && !stat.samples.includes(sample)) stat.samples.push(sample);
    }
  }
  const columns = stats.map((stat, i): InspectColumn => {
    const col = minCol + i;
    const headerCell = headerCells?.find((cell) => cell.col === col);
    const nonEmpty = TYPE_ORDER.reduce((sum, t) => sum + stat.counts[t], 0);
    let dominantType: CellType = 'empty';
    let best = 0;
    for (const t of TYPE_ORDER) {
      if (stat.counts[t] > best) {
        best = stat.counts[t];
        dominantType = t;
      }
    }
    return {
      index: col,
      letter: columnLetters(col),
      header: headerCell !== undefined && typeof headerCell.value === 'string' ? headerCell.value : null,
      dominantType,
      types: { ...stat.counts, empty: dataRowCount - nonEmpty },
      samples: stat.samples,
    };
  });
  return {
    ...base,
    rowCount: maxRow - minRow + 1,
    columnCount: width,
    headerRow,
    columns,
    columnsTruncated: width > shownColumns,
  };
}

/**
 * Header heuristic: the first of the leading non-empty rows whose cells are all distinct,
 * non-blank strings covering at least half of the used width. Confidence is high when the
 * next non-empty row contains a value that is not a string.
 */
export function detectHeader(
  rows: readonly (readonly Cell[])[],
  width: number,
  scanRows: number,
): { row: number; confidence: 'high' | 'low' } | null {
  const candidates = rows.filter((r) => r.length > 0).slice(0, scanRows);
  for (let i = 0; i < candidates.length; i++) {
    const cells = candidates[i] ?? [];
    if (cells.length < Math.max(1, Math.ceil(width / 2))) continue;
    const seen = new Set<string>();
    const allText = cells.every((cell) => {
      if (cell.type !== 'string' || typeof cell.value !== 'string') return false;
      const key = cell.value.trim().toLowerCase();
      if (key.length === 0 || seen.has(key)) return false;
      seen.add(key);
      return true;
    });
    if (!allText) continue;
    const next = rows.filter((r) => r.length > 0).find((r) => (r[0]?.row ?? 0) > (cells[0]?.row ?? 0));
    const confidence = next !== undefined && next.some((cell) => cell.type !== 'string' && cell.type !== 'empty') ? 'high' : 'low';
    return { row: cells[0]?.row ?? 0, confidence };
  }
  return null;
}

/** Warns when cells lie outside the declared dimension (a dimension larger than the cells is normal). */
function checkDimension(
  sheet: Sheet,
  used: { minRow: number; maxRow: number; minCol: number; maxCol: number } | null,
  warnings: WarningCollector,
): void {
  if (sheet.dimension === null || used === null) return;
  const declared = parseRange(sheet.dimension);
  if (declared === null) return;
  const inside =
    used.minRow >= declared.start.row &&
    used.maxRow <= declared.end.row &&
    used.minCol >= declared.start.col &&
    used.maxCol <= declared.end.col;
  if (!inside) {
    warnings.add(
      'XLV_W_DIMENSION_MISMATCH',
      `The sheet declares the range ${sheet.dimension}, but it has cells outside it; programs that trust the declared range may miss data.`,
      { sheet: sheet.name },
    );
  }
}

/** Warns about values in merged ranges other than the top-left cell; applications do not show them. */
function checkMergedValues(sheet: Sheet, rows: readonly (readonly Cell[])[], warnings: WarningCollector): void {
  // Duplicate ranges are dropped first; real files have few, non-overlapping merges.
  const merges = [...new Set(sheet.merges)]
    .map((m) => parseRange(m))
    .filter((r): r is CellRange => r !== null)
    .sort((a, b) => a.start.row - b.start.row);
  if (merges.length === 0) return;
  // Overlapping merges in a crafted file could make this sweep cost merges × rows; the check
  // stops after a fixed amount of work, since it only produces a warning (D-115).
  let work = 0;
  let next = 0;
  let active: CellRange[] = [];
  for (const cells of rows) {
    const r = cells[0]?.row;
    if (r === undefined) continue;
    while (next < merges.length && (merges[next]?.start.row ?? Number.POSITIVE_INFINITY) <= r) {
      const m = merges[next++];
      if (m !== undefined) active.push(m);
    }
    work += active.length;
    active = active.filter((m) => m.end.row >= r);
    if (active.length === 0) continue;
    for (const cell of cells) {
      if (cell.value === null && cell.formula === null) continue;
      for (const m of active) {
        work++;
        const anchor = cell.row === m.start.row && cell.col === m.start.col;
        if (!anchor && cell.col >= m.start.col && cell.col <= m.end.col) {
          warnings.add(
            'XLV_W_VALUES_IN_MERGED_CELLS',
            'Some merged ranges have values in cells other than the top-left one; spreadsheet applications do not display them.',
            { sheet: sheet.name, cell: cell.ref },
          );
          break;
        }
      }
    }
    if (work > MAX_MERGE_CHECK_WORK) return;
  }
}
