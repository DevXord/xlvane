// Sheet protection flags, page setup, headers and footers, page breaks, print area and print
// titles (ECMA-376 Part 1 §18.3.1.3, .14, .46, .62–.65, .70, .74, .85, §18.2.5, §18.2.29;
// [MS-OI29500] §2.1.586–§2.1.654; docs/facts/sheet-rules.md SR-20 … SR-31; D-C03, D-C04).
// Password hashes are read and kept as data; writing them is phase C2.
import { XlvError } from './errors';
import type { NormalizedDefinedName } from './defined-names';
import { MAX_COLUMNS, MAX_ROWS, columnLetters, columnNumber, parseRange, type CellRange } from './sml/cell-ref';
import { decodeXstring } from './sml/xstring';
import { encodeXstring, escapeAttr, escapeText, firstLoneSurrogate } from './xml/writer';

// ---------------------------------------------------------------------------------------------
// Types

/**
 * The flags of `sheetProtection` (ECMA-376 Part 1 §18.3.1.85, SR-20), named as in the file. For
 * every flag except `sheet`, `true` means the action is **locked** (not allowed) while the sheet
 * is protected and `false` that it is allowed; `sheet` turns protection on, and the other flags
 * apply only when it is `true`.
 *
 * @example
 * ```ts
 * import type { SheetProtectionFlags } from 'xlvane';
 *
 * // Protected, but sorting and the autofilter stay allowed.
 * const flags: SheetProtectionFlags = { sort: false, autoFilter: false };
 * console.log(flags);
 * ```
 */
export interface SheetProtectionFlags {
  /** Editing objects locked. Default `false`. */
  readonly objects?: boolean;
  /** Editing scenarios locked. Default `false`. */
  readonly scenarios?: boolean;
  /** Formatting cells locked. Default `true`. */
  readonly formatCells?: boolean;
  /** Default `true`. */
  readonly formatColumns?: boolean;
  /** Default `true`. */
  readonly formatRows?: boolean;
  /** Default `true`. */
  readonly insertColumns?: boolean;
  /** Default `true`. */
  readonly insertRows?: boolean;
  /** Default `true`. */
  readonly insertHyperlinks?: boolean;
  /** Default `true`. */
  readonly deleteColumns?: boolean;
  /** Default `true`. */
  readonly deleteRows?: boolean;
  /** Selecting locked cells locked. Default `false`. */
  readonly selectLockedCells?: boolean;
  /** Selecting unlocked cells locked. Default `false`. */
  readonly selectUnlockedCells?: boolean;
  /** Default `true`. */
  readonly sort?: boolean;
  /** Using the autofilter locked. Default `true`. */
  readonly autoFilter?: boolean;
  /** Default `true`. */
  readonly pivotTables?: boolean;
}

/**
 * Options of {@link NewSheet.protect}: the flags of {@link SheetProtectionFlags}. A password is
 * not supported yet (it needs the hash of phase C2) and throws `XLV_E_NOT_SUPPORTED`.
 *
 * @example
 * ```ts
 * import { createWorkbook, type SheetProtectionOptions } from 'xlvane';
 *
 * const options: SheetProtectionOptions = { formatColumns: false };
 * createWorkbook().addSheet('Form').protect(options);
 * ```
 */
export interface SheetProtectionOptions extends SheetProtectionFlags {
  /** Not supported yet; throws `XLV_E_NOT_SUPPORTED`. */
  readonly password?: never;
}

/**
 * The `sheetProtection` element of a sheet as read (SR-20). The flags have the meaning given in
 * {@link SheetProtectionFlags}, with the schema defaults where the file leaves them out. The
 * password verifier is kept as stored and never interpreted: `algorithmName`, `hashValue`,
 * `saltValue` and `spinCount`, or the legacy 4-digit `password` hash of Transitional files.
 *
 * @example
 * ```ts
 * import { readWorkbook, type SheetProtection } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const protection: SheetProtection | null = readWorkbook(bytes).sheets[0]!.protection;
 * console.log(protection?.sheet, protection?.formatCells, protection?.hashValue !== null);
 * ```
 */
export interface SheetProtection extends Required<SheetProtectionFlags> {
  /** `true` when the sheet is protected. */
  readonly sheet: boolean;
  readonly algorithmName: string | null;
  readonly hashValue: string | null;
  readonly saltValue: string | null;
  readonly spinCount: number | null;
  /** Legacy (Transitional) password hash, 4 hexadecimal digits. */
  readonly password: string | null;
}

/**
 * The `workbookProtection` element as read (ECMA-376 Part 1 §18.2.29, SR-22). `lockStructure`
 * and `lockWindows` are `true` when the structure or the windows are locked. Hashes are kept as
 * stored and never interpreted.
 *
 * @example
 * ```ts
 * import { readWorkbook, type WorkbookProtection } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const protection: WorkbookProtection | null = readWorkbook(bytes).protection;
 * console.log(protection?.lockStructure);
 * ```
 */
export interface WorkbookProtection {
  readonly lockStructure: boolean;
  readonly lockWindows: boolean;
  readonly lockRevision: boolean;
  readonly workbookAlgorithmName: string | null;
  readonly workbookHashValue: string | null;
  readonly workbookSaltValue: string | null;
  readonly workbookSpinCount: number | null;
  readonly revisionsAlgorithmName: string | null;
  readonly revisionsHashValue: string | null;
  readonly revisionsSaltValue: string | null;
  readonly revisionsSpinCount: number | null;
  /** Legacy (Transitional) hashes, 4 hexadecimal digits. */
  readonly workbookPassword: string | null;
  readonly revisionsPassword: string | null;
}

/**
 * Page margins in inches (`pageMargins`, ECMA-376 Part 1 §18.3.1.62, SR-24). Each is at least 0
 * and below 49 ([MS-OI29500] §2.1.637).
 *
 * @example
 * ```ts
 * import type { PageMargins } from 'xlvane';
 *
 * const narrow: PageMargins = { left: 0.25, right: 0.25, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 };
 * console.log(narrow.left);
 * ```
 */
export interface PageMargins {
  readonly left: number;
  readonly right: number;
  readonly top: number;
  readonly bottom: number;
  readonly header: number;
  readonly footer: number;
}

/**
 * Print options (`printOptions`, ECMA-376 Part 1 §18.3.1.70, SR-25). All default `false`.
 *
 * @example
 * ```ts
 * import type { PrintOptions } from 'xlvane';
 *
 * const options: PrintOptions = { gridLines: true, horizontalCentered: true };
 * console.log(options);
 * ```
 */
export interface PrintOptions {
  /** Print the grid lines. */
  readonly gridLines?: boolean;
  /** Print row numbers and column letters. */
  readonly headings?: boolean;
  readonly horizontalCentered?: boolean;
  readonly verticalCentered?: boolean;
}

/**
 * Page orientation (ST_Orientation, §18.18.50); `"default"` leaves it to the application.
 *
 * @example
 * ```ts
 * import { createWorkbook, type PageOrientation } from 'xlvane';
 *
 * const value: PageOrientation = 'landscape';
 * createWorkbook().addSheet('Report').pageSetup({ orientation: value });
 * ```
 */
export type PageOrientation = 'default' | 'portrait' | 'landscape';
/**
 * How cell notes are printed (ST_CellComments, §18.18.5).
 *
 * @example
 * ```ts
 * import { createWorkbook, type PrintCellComments } from 'xlvane';
 *
 * const value: PrintCellComments = 'atEnd';
 * createWorkbook().addSheet('Report').pageSetup({ cellComments: value });
 * ```
 */
export type PrintCellComments = 'none' | 'asDisplayed' | 'atEnd';
/**
 * How cells with errors are printed (ST_PrintError, §18.18.60).
 *
 * @example
 * ```ts
 * import { createWorkbook, type PrintErrors } from 'xlvane';
 *
 * const value: PrintErrors = 'blank';
 * createWorkbook().addSheet('Report').pageSetup({ errors: value });
 * ```
 */
export type PrintErrors = 'displayed' | 'blank' | 'dash' | 'NA';
/**
 * Order of printed pages (ST_PageOrder, §18.18.51).
 *
 * @example
 * ```ts
 * import { createWorkbook, type PageOrder } from 'xlvane';
 *
 * const value: PageOrder = 'overThenDown';
 * createWorkbook().addSheet('Report').pageSetup({ pageOrder: value });
 * ```
 */
export type PageOrder = 'downThenOver' | 'overThenDown';

/**
 * Options of {@link NewSheet.pageSetup} (`pageSetup`, `pageMargins`, `printOptions` and
 * `sheetPr/pageSetUpPr@fitToPage`; ECMA-376 Part 1 §18.3.1.62–§18.3.1.65, §18.3.1.70; SR-23 …
 * SR-26). Limits are those of [MS-OI29500] §2.1.637 and §2.1.638.
 *
 * @example
 * ```ts
 * import { createWorkbook, type PageSetupOptions } from 'xlvane';
 *
 * const setup: PageSetupOptions = { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
 * createWorkbook().addSheet('Report').pageSetup(setup);
 * ```
 */
export interface PageSetupOptions {
  /** Margins in inches; values not given take those of the Part 1 example (0.7, 0.7, 0.75, 0.75, 0.3, 0.3). */
  readonly margins?: Partial<PageMargins>;
  readonly orientation?: PageOrientation;
  /** Paper size code of Part 1 §18.3.1.63, e.g. 1 = Letter, 9 = A4. Default 1. */
  readonly paperSize?: number;
  /** Print scale in percent: 10 to 400, or 0 (automatic, [MS-OI29500] §2.1.638 a). Default 100. */
  readonly scale?: number;
  /** Turns on "fit to page" (`pageSetUpPr@fitToPage`); then `fitToWidth`/`fitToHeight` apply. */
  readonly fitToPage?: boolean;
  /** Pages across, 0 to 32767. Default 1. */
  readonly fitToWidth?: number;
  /** Pages down, 0 to 32767. Default 1. */
  readonly fitToHeight?: number;
  readonly pageOrder?: PageOrder;
  readonly blackAndWhite?: boolean;
  /** Print without graphics. */
  readonly draft?: boolean;
  readonly cellComments?: PrintCellComments;
  readonly errors?: PrintErrors;
  /** Number of the first page, 0 to 32767; `null` numbers pages automatically. */
  readonly firstPageNumber?: number | null;
  readonly horizontalDpi?: number | null;
  readonly verticalDpi?: number | null;
  /** 1 to 32767. Default 1. */
  readonly copies?: number;
  readonly printOptions?: PrintOptions;
}

/**
 * Page setup of a sheet as read, with the schema defaults where the file leaves values out.
 * `margins` is `null` when the sheet has no `pageMargins`.
 *
 * @example
 * ```ts
 * import { readWorkbook, type PageSetup } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const setup: PageSetup = readWorkbook(bytes).sheets[0]!.pageSetup;
 * console.log(setup.orientation, setup.paperSize, setup.margins?.left);
 * ```
 */
export interface PageSetup {
  readonly margins: PageMargins | null;
  readonly orientation: PageOrientation;
  readonly paperSize: number;
  readonly scale: number;
  readonly fitToPage: boolean;
  readonly fitToWidth: number;
  readonly fitToHeight: number;
  readonly pageOrder: PageOrder;
  readonly blackAndWhite: boolean;
  readonly draft: boolean;
  readonly cellComments: PrintCellComments;
  readonly errors: PrintErrors;
  readonly firstPageNumber: number | null;
  readonly horizontalDpi: number | null;
  readonly verticalDpi: number | null;
  readonly copies: number;
  readonly printOptions: Required<PrintOptions>;
}

/**
 * Headers and footers (`headerFooter`, ECMA-376 Part 1 §18.3.1.46, SR-27). Texts use the
 * formatting codes of §18.3.1.46 (`&L`, `&C`, `&R` for the sections, `&P` page number, `&N`
 * pages, `&D` date, `&A` sheet name, `&&` an ampersand, …); xlvane stores them as given. Each
 * text has at most 255 characters ([MS-OI29500] §2.1.618 b … §2.1.634 a).
 *
 * @example
 * ```ts
 * import { createWorkbook, type HeaderFooter } from 'xlvane';
 *
 * const hf: HeaderFooter = { oddHeader: '&CQuarterly report', oddFooter: '&LPrinted &D&RPage &P of &N' };
 * createWorkbook().addSheet('Report').headerFooter(hf);
 * ```
 */
export interface HeaderFooter {
  readonly oddHeader?: string;
  readonly oddFooter?: string;
  /** Used only with `differentOddEven`. */
  readonly evenHeader?: string;
  readonly evenFooter?: string;
  /** Used only with `differentFirst`. */
  readonly firstHeader?: string;
  readonly firstFooter?: string;
  /** Default `false`. */
  readonly differentOddEven?: boolean;
  /** Default `false`. */
  readonly differentFirst?: boolean;
  /** Default `true`. */
  readonly scaleWithDoc?: boolean;
  /** Default `true`. */
  readonly alignWithMargins?: boolean;
}

/**
 * Manual page breaks (`rowBreaks`, `colBreaks`, ECMA-376 Part 1 §18.3.1.3, .14, .74; SR-28):
 * the rows and columns that start a new printed page, as row numbers (2 to 1048576) and column
 * numbers or letters ("B" to "XFD").
 *
 * @example
 * ```ts
 * import { createWorkbook, type PageBreaks } from 'xlvane';
 *
 * const breaks: PageBreaks = { rows: [25, 49], columns: ['F'] };
 * createWorkbook().addSheet('Report').pageBreaks(breaks);
 * ```
 */
export interface PageBreaks {
  readonly rows?: readonly number[];
  readonly columns?: readonly (number | string)[];
}

/**
 * Rows and columns repeated on every printed page (`_xlnm.Print_Titles`, SR-30), as `"1:2"` or
 * `"1"` for rows and `"A:B"` or `"A"` for columns; `null` when not set.
 *
 * @example
 * ```ts
 * import { createWorkbook, type PrintTitles } from 'xlvane';
 *
 * const titles: PrintTitles = { rows: '1:2', columns: 'A' };
 * createWorkbook().addSheet('Report').printTitles(titles);
 * ```
 */
export interface PrintTitles {
  readonly rows?: string | null;
  readonly columns?: string | null;
}

/** @internal Everything a sheet stores about printing and protection, normalised. */
export interface SheetPrintData {
  readonly protection: Required<SheetProtectionFlags> | null;
  readonly pageSetup: PageSetup;
  /** `true` when `pageSetup` was called (so the elements are written). */
  readonly pageSetupSet: boolean;
  readonly headerFooter: NormalizedHeaderFooter | null;
  readonly rowBreaks: readonly number[];
  readonly colBreaks: readonly number[];
  readonly printArea: readonly CellRange[] | null;
  readonly printTitles: { readonly rows: readonly [number, number] | null; readonly cols: readonly [number, number] | null } | null;
}

type NormalizedHeaderFooter = Required<HeaderFooter>;

// ---------------------------------------------------------------------------------------------
// Defaults (schema defaults of sml.xsd CT_SheetProtection, CT_PageSetup, CT_PrintOptions,
// CT_HeaderFooter; SR-20, SR-23, SR-25, SR-27)

const PROTECTION_FLAGS = [
  'objects',
  'scenarios',
  'formatCells',
  'formatColumns',
  'formatRows',
  'insertColumns',
  'insertRows',
  'insertHyperlinks',
  'deleteColumns',
  'deleteRows',
  'selectLockedCells',
  'sort',
  'autoFilter',
  'pivotTables',
  'selectUnlockedCells',
] as const;

/** @internal */
export const DEFAULT_PROTECTION: Required<SheetProtectionFlags> = Object.freeze({
  objects: false,
  scenarios: false,
  formatCells: true,
  formatColumns: true,
  formatRows: true,
  insertColumns: true,
  insertRows: true,
  insertHyperlinks: true,
  deleteColumns: true,
  deleteRows: true,
  selectLockedCells: false,
  sort: true,
  autoFilter: true,
  pivotTables: true,
  selectUnlockedCells: false,
});

/** Margins of the Part 1 §18.3.1.62 example (L78045), used for values not given (SR-24). */
const EXAMPLE_MARGINS: PageMargins = Object.freeze({ left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 });

/** @internal The page setup of a sheet without print elements. */
export const DEFAULT_PAGE_SETUP: PageSetup = Object.freeze({
  margins: null,
  orientation: 'default',
  paperSize: 1,
  scale: 100,
  fitToPage: false,
  fitToWidth: 1,
  fitToHeight: 1,
  pageOrder: 'downThenOver',
  blackAndWhite: false,
  draft: false,
  cellComments: 'none',
  errors: 'displayed',
  firstPageNumber: null,
  horizontalDpi: null,
  verticalDpi: null,
  copies: 1,
  printOptions: Object.freeze({ gridLines: false, headings: false, horizontalCentered: false, verticalCentered: false }),
});

const DEFAULT_HEADER_FOOTER: NormalizedHeaderFooter = Object.freeze({
  oddHeader: '',
  oddFooter: '',
  evenHeader: '',
  evenFooter: '',
  firstHeader: '',
  firstFooter: '',
  differentOddEven: false,
  differentFirst: false,
  scaleWithDoc: true,
  alignWithMargins: true,
});

/** @internal */
export const EMPTY_PRINT_DATA: SheetPrintData = Object.freeze({
  protection: null,
  pageSetup: DEFAULT_PAGE_SETUP,
  pageSetupSet: false,
  headerFooter: null,
  rowBreaks: [],
  colBreaks: [],
  printArea: null,
  printTitles: null,
});

/** [MS-OI29500] §2.1.618 b, §2.1.619 c, §2.1.621 a, §2.1.622 a, §2.1.633 a, §2.1.634 a. */
const MAX_HEADER_FOOTER_LENGTH = 255;
/** [MS-OI29500] §2.1.599 a (colBreaks ≤ 1023) and §2.1.645 a (rowBreaks < 1023). */
const MAX_COL_BREAKS = 1023;
const MAX_ROW_BREAKS = 1022;

// ---------------------------------------------------------------------------------------------
// Validation

function invalid(message: string, hint: string, option: string): XlvError {
  return new XlvError('XLV_E_INVALID_OPTION', message, { hint, details: { option } });
}

function readObject(value: unknown, known: readonly string[], what: string): Map<string, unknown> {
  const out = new Map<string, unknown>();
  if (value === undefined) return out;
  if (typeof value !== 'object' || value === null || Array.isArray(value)) throw invalid(`${what} must be an object.`, `Pass an object with some of these properties: ${known.join(', ')}.`, what);
  for (const key of Object.keys(value)) {
    if (!known.includes(key)) throw invalid(`Unknown option "${key.slice(0, 40)}" in ${what}.`, `Known options: ${known.join(', ')}.`, key.slice(0, 40));
    const v = (value as Record<string, unknown>)[key];
    if (v !== undefined) out.set(key, v);
  }
  return out;
}

function bool(name: string, v: unknown): boolean {
  if (typeof v !== 'boolean') throw invalid(`The option "${name}" must be true or false.`, `For example { ${name}: true }.`, name);
  return v;
}

function integer(name: string, v: unknown, min: number, max: number, example: number): number {
  if (typeof v !== 'number' || !Number.isInteger(v) || v < min || v > max) {
    throw invalid(`The option "${name}" must be a whole number from ${String(min)} to ${String(max)}.`, `For example { ${name}: ${String(example)} }.`, name);
  }
  return v;
}

function oneOf<T extends string>(name: string, v: unknown, allowed: readonly T[]): T {
  if (typeof v === 'string' && (allowed as readonly string[]).includes(v)) return v as T;
  throw invalid(`The option "${name}" has the value ${JSON.stringify(String(v).slice(0, 40))}.`, `Use ${allowed.map((a) => `"${a}"`).join(', ')}.`, name);
}

/** @internal Checks the options of `NewSheet.protect`. */
export function normalizeProtection(options: unknown, sheet: string): Required<SheetProtectionFlags> {
  if (typeof options === 'object' && options !== null && 'password' in options) {
    throw new XlvError('XLV_E_NOT_SUPPORTED', 'Protecting a sheet with a password is not supported yet.', {
      location: { sheet },
      hint: 'Protect the sheet without a password for now; password hashing comes with phase C2 of xlvane. Remove the "password" option.',
      details: { option: 'password' },
    });
  }
  const o = readObject(options, PROTECTION_FLAGS, 'protect options');
  const out: Record<string, boolean> = { ...DEFAULT_PROTECTION };
  for (const [k, v] of o) out[k] = bool(k, v);
  return out as unknown as Required<SheetProtectionFlags>;
}

const ORIENTATIONS: readonly PageOrientation[] = ['default', 'portrait', 'landscape'];
const COMMENTS: readonly PrintCellComments[] = ['none', 'asDisplayed', 'atEnd'];
const ERRORS: readonly PrintErrors[] = ['displayed', 'blank', 'dash', 'NA'];
const ORDERS: readonly PageOrder[] = ['downThenOver', 'overThenDown'];
const SETUP_KEYS = [
  'margins',
  'orientation',
  'paperSize',
  'scale',
  'fitToPage',
  'fitToWidth',
  'fitToHeight',
  'pageOrder',
  'blackAndWhite',
  'draft',
  'cellComments',
  'errors',
  'firstPageNumber',
  'horizontalDpi',
  'verticalDpi',
  'copies',
  'printOptions',
] as const;

/** @internal Applies the options of `NewSheet.pageSetup` to `base`. */
export function normalizePageSetup(options: unknown, base: PageSetup): PageSetup {
  const o = readObject(options, SETUP_KEYS, 'pageSetup options');
  const next: { -readonly [K in keyof PageSetup]: PageSetup[K] } = { ...base };
  const margins = o.get('margins');
  if (margins !== undefined) {
    const m = readObject(margins, ['left', 'right', 'top', 'bottom', 'header', 'footer'], 'margins');
    const out: Record<string, number> = { ...(base.margins ?? EXAMPLE_MARGINS) };
    for (const [k, v] of m) {
      // [MS-OI29500] §2.1.637: 0 inclusive to 49 exclusive.
      if (typeof v !== 'number' || !Number.isFinite(v) || v < 0 || v >= 49) throw invalid(`The margin "${k}" must be a number of inches from 0 to below 49.`, `For example { margins: { ${k}: 0.75 } }.`, k);
      out[k] = v;
    }
    next.margins = out as unknown as PageMargins;
  }
  if (o.has('orientation')) next.orientation = oneOf('orientation', o.get('orientation'), ORIENTATIONS);
  if (o.has('paperSize')) next.paperSize = integer('paperSize', o.get('paperSize'), 1, 2147483647, 9);
  if (o.has('scale')) {
    const s = o.get('scale');
    // 10–400 (Part 1 §18.3.1.63); 0 = automatic ([MS-OI29500] §2.1.638 a).
    if (typeof s !== 'number' || !Number.isInteger(s) || (s !== 0 && (s < 10 || s > 400))) throw invalid('The option "scale" must be a whole number from 10 to 400 (percent), or 0 for automatic.', 'For example { scale: 80 }.', 'scale');
    next.scale = s;
  }
  if (o.has('fitToPage')) next.fitToPage = bool('fitToPage', o.get('fitToPage'));
  if (o.has('fitToWidth')) next.fitToWidth = integer('fitToWidth', o.get('fitToWidth'), 0, 32767, 1);
  if (o.has('fitToHeight')) next.fitToHeight = integer('fitToHeight', o.get('fitToHeight'), 0, 32767, 1);
  if (o.has('pageOrder')) next.pageOrder = oneOf('pageOrder', o.get('pageOrder'), ORDERS);
  if (o.has('blackAndWhite')) next.blackAndWhite = bool('blackAndWhite', o.get('blackAndWhite'));
  if (o.has('draft')) next.draft = bool('draft', o.get('draft'));
  if (o.has('cellComments')) next.cellComments = oneOf('cellComments', o.get('cellComments'), COMMENTS);
  if (o.has('errors')) next.errors = oneOf('errors', o.get('errors'), ERRORS);
  if (o.has('firstPageNumber')) {
    const v = o.get('firstPageNumber');
    next.firstPageNumber = v === null ? null : integer('firstPageNumber', v, 0, 32767, 1);
  }
  for (const k of ['horizontalDpi', 'verticalDpi'] as const) {
    if (!o.has(k)) continue;
    const v = o.get(k);
    next[k] = v === null ? null : integer(k, v, 1, 4294967295, 600);
  }
  if (o.has('copies')) next.copies = integer('copies', o.get('copies'), 1, 32767, 2);
  const po = o.get('printOptions');
  if (po !== undefined) {
    const p = readObject(po, ['gridLines', 'headings', 'horizontalCentered', 'verticalCentered'], 'printOptions');
    const out: Record<string, boolean> = { ...base.printOptions };
    for (const [k, v] of p) out[k] = bool(k, v);
    next.printOptions = out as unknown as Required<PrintOptions>;
  }
  return next;
}

/** @internal Applies the options of `NewSheet.headerFooter` to `base` (or the defaults). */
export function normalizeHeaderFooter(options: unknown, base: NormalizedHeaderFooter | null): NormalizedHeaderFooter {
  const texts = ['oddHeader', 'oddFooter', 'evenHeader', 'evenFooter', 'firstHeader', 'firstFooter'] as const;
  const flags = ['differentOddEven', 'differentFirst', 'scaleWithDoc', 'alignWithMargins'] as const;
  const o = readObject(options, [...texts, ...flags], 'headerFooter options');
  const out: Record<string, string | boolean> = { ...(base ?? DEFAULT_HEADER_FOOTER) };
  for (const k of texts) {
    if (!o.has(k)) continue;
    const v = o.get(k);
    if (typeof v !== 'string' || firstLoneSurrogate(v) !== -1) throw invalid(`The option "${k}" must be text.`, `For example { ${k}: '&CPage &P' }.`, k);
    if (v.length > MAX_HEADER_FOOTER_LENGTH) throw invalid(`The option "${k}" has ${String(v.length)} characters; Office allows at most ${String(MAX_HEADER_FOOTER_LENGTH)}.`, 'Shorten the text.', k);
    out[k] = v;
  }
  for (const k of flags) if (o.has(k)) out[k] = bool(k, o.get(k));
  return out as unknown as NormalizedHeaderFooter;
}

/** @internal Checks the options of `NewSheet.pageBreaks`; returns sorted, unique row and column numbers. */
export function normalizePageBreaks(options: unknown): { rows: number[]; cols: number[] } {
  const o = readObject(options, ['rows', 'columns'], 'pageBreaks options');
  const list = (key: string): unknown[] => {
    const v = o.get(key) ?? [];
    if (!Array.isArray(v)) throw invalid(`The option "${key}" must be an array.`, key === 'rows' ? 'For example { rows: [25, 49] }.' : 'For example { columns: ["F", 12] }.', key);
    return v;
  };
  const rows = [...new Set(list('rows').map((r) => integer('rows', r, 2, MAX_ROWS, 25)))].sort((a, b) => a - b);
  const cols = [
    ...new Set(
      list('columns').map((c) => {
        const n = typeof c === 'string' ? columnNumber(c.toUpperCase()) : c;
        if (typeof n !== 'number' || !Number.isInteger(n) || n < 2 || n > MAX_COLUMNS) throw invalid(`The page break column ${JSON.stringify(String(c).slice(0, 20))} is not valid.`, 'Use a column from "B" to "XFD" (or 2 to 16384); the column starts a new page.', 'columns');
        return n;
      }),
    ),
  ].sort((a, b) => a - b);
  // [MS-OI29500] §2.1.645 a (rowBreaks count < 1023) and §2.1.599 a (colBreaks count ≤ 1023).
  if (rows.length > MAX_ROW_BREAKS) throw invalid(`A sheet can have at most ${String(MAX_ROW_BREAKS)} row page breaks; ${String(rows.length)} were given.`, 'Remove some breaks.', 'rows');
  if (cols.length > MAX_COL_BREAKS) throw invalid(`A sheet can have at most ${String(MAX_COL_BREAKS)} column page breaks; ${String(cols.length)} were given.`, 'Remove some breaks.', 'columns');
  return { rows, cols };
}

/** @internal Parses the ranges of `NewSheet.printArea`. */
export function normalizePrintArea(value: unknown): CellRange[] | null {
  if (value === null) return null;
  const list = typeof value === 'string' ? [value] : value;
  if (!Array.isArray(list) || list.length === 0) throw invalid('The print area must be a range or a non-empty array of ranges.', 'For example printArea("A1:H40") or printArea(["A1:D20", "F1:H20"]).', 'printArea');
  return list.map((r) => {
    const range = typeof r === 'string' ? parseRange(r.replace(/\$/g, '').toUpperCase()) : null;
    if (range === null) throw invalid(`The print area range ${JSON.stringify(String(r).slice(0, 40))} is not valid.`, 'Use A1-style ranges such as "A1:H40".', 'printArea');
    return range;
  });
}

function span(value: unknown, kind: 'rows' | 'columns'): [number, number] | null {
  if (value === null || value === undefined) return null;
  const parse = (s: string): number | null => {
    if (kind === 'rows') {
      if (!/^\d{1,7}$/.test(s)) return null;
      const n = Number(s);
      return n >= 1 && n <= MAX_ROWS ? n : null;
    }
    return columnNumber(s.toUpperCase());
  };
  const parts = typeof value === 'string' ? value.replace(/\$/g, '').split(':') : [];
  const a = parts.length >= 1 && parts.length <= 2 ? parse(parts[0]!) : null;
  const b = parts.length === 2 ? parse(parts[1]!) : a;
  if (a === null || b === null) {
    throw invalid(`The print titles ${kind} ${JSON.stringify(String(value).slice(0, 40))} are not valid.`, kind === 'rows' ? 'Use rows such as "1" or "1:2".' : 'Use columns such as "A" or "A:B".', kind);
  }
  return [Math.min(a, b), Math.max(a, b)];
}

/** @internal Parses the options of `NewSheet.printTitles`. */
export function normalizePrintTitles(value: unknown): SheetPrintData['printTitles'] {
  if (value === null) return null;
  const o = readObject(value, ['rows', 'columns'], 'printTitles options');
  const rows = span(o.get('rows'), 'rows');
  const cols = span(o.get('columns'), 'columns');
  return rows === null && cols === null ? null : { rows, cols };
}

// ---------------------------------------------------------------------------------------------
// Writing

function a(name: string, value: string | number): string {
  return ` ${name}="${typeof value === 'number' ? String(value) : escapeAttr(value)}"`;
}

/** @internal `sheetProtection` with `sheet="1"` and the flags that differ from the schema defaults. */
export function sheetProtectionXml(p: Required<SheetProtectionFlags> | null): string {
  if (p === null) return '';
  let xml = '<sheetProtection sheet="1"';
  for (const k of PROTECTION_FLAGS) if (p[k] !== DEFAULT_PROTECTION[k]) xml += a(k, p[k] ? 1 : 0);
  return `${xml}/>`;
}

/**
 * @internal The elements after `hyperlinks` in CT_Worksheet order (sml.xsd T2192–T2197, SR-23):
 * printOptions, pageMargins, pageSetup, headerFooter, rowBreaks, colBreaks.
 */
export function printXml(d: SheetPrintData): string {
  let xml = '';
  const s = d.pageSetup;
  if (d.pageSetupSet) {
    const po = s.printOptions;
    if (po.horizontalCentered || po.verticalCentered || po.headings || po.gridLines) {
      xml += '<printOptions';
      if (po.horizontalCentered) xml += ' horizontalCentered="1"';
      if (po.verticalCentered) xml += ' verticalCentered="1"';
      if (po.headings) xml += ' headings="1"';
      if (po.gridLines) xml += ' gridLines="1"';
      xml += '/>';
    }
    if (s.margins !== null) {
      const m = s.margins;
      xml += `<pageMargins${a('left', m.left)}${a('right', m.right)}${a('top', m.top)}${a('bottom', m.bottom)}${a('header', m.header)}${a('footer', m.footer)}/>`;
    }
    let ps = '';
    if (s.paperSize !== 1) ps += a('paperSize', s.paperSize);
    if (s.scale !== 100) ps += a('scale', s.scale);
    if (s.firstPageNumber !== null) ps += a('firstPageNumber', s.firstPageNumber);
    if (s.fitToWidth !== 1) ps += a('fitToWidth', s.fitToWidth);
    if (s.fitToHeight !== 1) ps += a('fitToHeight', s.fitToHeight);
    if (s.pageOrder !== 'downThenOver') ps += a('pageOrder', s.pageOrder);
    if (s.orientation !== 'default') ps += a('orientation', s.orientation);
    if (s.blackAndWhite) ps += ' blackAndWhite="1"';
    if (s.draft) ps += ' draft="1"';
    if (s.cellComments !== 'none') ps += a('cellComments', s.cellComments);
    if (s.firstPageNumber !== null) ps += ' useFirstPageNumber="1"';
    if (s.errors !== 'displayed') ps += a('errors', s.errors);
    if (s.horizontalDpi !== null) ps += a('horizontalDpi', s.horizontalDpi);
    if (s.verticalDpi !== null) ps += a('verticalDpi', s.verticalDpi);
    if (s.copies !== 1) ps += a('copies', s.copies);
    if (ps !== '') xml += `<pageSetup${ps}/>`;
  }
  const hf = d.headerFooter;
  if (hf !== null) {
    let attrs = '';
    if (hf.differentOddEven) attrs += ' differentOddEven="1"';
    if (hf.differentFirst) attrs += ' differentFirst="1"';
    if (!hf.scaleWithDoc) attrs += ' scaleWithDoc="0"';
    if (!hf.alignWithMargins) attrs += ' alignWithMargins="0"';
    let body = '';
    for (const k of ['oddHeader', 'oddFooter', 'evenHeader', 'evenFooter', 'firstHeader', 'firstFooter'] as const) {
      if (hf[k] !== '') body += `<${k}>${escapeText(encodeXstring(hf[k]))}</${k}>`;
    }
    if (attrs !== '' || body !== '') xml += `<headerFooter${attrs}>${body}</headerFooter>`;
  }
  // brk id is the zero-based row or column before which the break occurs; max spans the other
  // axis, as in the Part 1 examples (§18.3.1.74 L79307, §18.3.1.14 L75488; SR-28).
  const breaks = (element: string, list: readonly number[], max: number): string =>
    list.length === 0 ? '' : `<${element}${a('count', list.length)}${a('manualBreakCount', list.length)}>${list.map((n) => `<brk${a('id', n - 1)}${a('max', max)} man="1"/>`).join('')}</${element}>`;
  xml += breaks('rowBreaks', d.rowBreaks, MAX_COLUMNS - 1);
  xml += breaks('colBreaks', d.colBreaks, MAX_ROWS - 1);
  return xml;
}

function quoteSheet(name: string): string {
  return `'${name.replace(/'/g, "''")}'`;
}

function abs(row: number, col: number): string {
  return `$${columnLetters(col)}$${String(row)}`;
}

/**
 * @internal The `_xlnm.Print_Area` and `_xlnm.Print_Titles` names of one sheet (SR-30). Their
 * text follows the files written by Excel in the Apache POI corpus (no document defines it):
 * `'Sheet'!$A$1:$J$20` (several areas comma-separated) and `'Sheet'!$A:$B,'Sheet'!$1:$2`.
 */
export function printNames(sheet: string, d: SheetPrintData): NormalizedDefinedName[] {
  const out: NormalizedDefinedName[] = [];
  const q = quoteSheet(sheet);
  if (d.printArea !== null) {
    const text = d.printArea
      .map((r) => {
        const one = r.start.row === r.end.row && r.start.col === r.end.col;
        return `${q}!${abs(r.start.row, r.start.col)}${one ? '' : `:${abs(r.end.row, r.end.col)}`}`;
      })
      .join(',');
    out.push({ name: '_xlnm.Print_Area', refersTo: text, sheet, hidden: false, comment: null });
  }
  const t = d.printTitles;
  if (t !== null) {
    const parts: string[] = [];
    if (t.cols !== null) parts.push(`${q}!$${columnLetters(t.cols[0])}:$${columnLetters(t.cols[1])}`);
    if (t.rows !== null) parts.push(`${q}!$${String(t.rows[0])}:$${String(t.rows[1])}`);
    out.push({ name: '_xlnm.Print_Titles', refersTo: parts.join(','), sheet, hidden: false, comment: null });
  }
  return out;
}

// ---------------------------------------------------------------------------------------------
// Reading

/** Splits formula text at commas outside quotes. */
function splitList(text: string): string[] {
  const out: string[] = [];
  let start = 0;
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (c === "'") quoted = !quoted;
    else if (c === ',' && !quoted) {
      out.push(text.slice(start, i));
      start = i + 1;
    }
  }
  out.push(text.slice(start));
  return out;
}

/** Removes the sheet prefix of `part` if it names `sheet`; `null` if it names another sheet or none. */
function withoutSheet(part: string, sheet: string): string | null {
  const bang = part.lastIndexOf('!');
  if (bang < 0) return null;
  let prefix = part.slice(0, bang).trim();
  if (prefix.startsWith("'") && prefix.endsWith("'") && prefix.length >= 2) prefix = prefix.slice(1, -1).replace(/''/g, "'");
  return prefix.toLowerCase() === sheet.toLowerCase() ? part.slice(bang + 1).trim().replace(/\$/g, '') : null;
}

/** The ranges of a `_xlnm.Print_Area` name of `sheet`, or `null` if its text is not a list of plain ranges. */
export function parsePrintArea(text: string, sheet: string): string[] | null {
  const out: string[] = [];
  for (const part of splitList(text)) {
    const ref = withoutSheet(part, sheet);
    const range = ref === null ? null : parseRange(ref.toUpperCase());
    if (range === null || ref === null) return null;
    out.push(ref.toUpperCase());
  }
  return out;
}

/** The rows and columns of a `_xlnm.Print_Titles` name of `sheet`, or `null` if its text has another form. */
export function parsePrintTitles(text: string, sheet: string): { rows: string | null; columns: string | null } | null {
  let rows: string | null = null;
  let columns: string | null = null;
  for (const part of splitList(text)) {
    const ref = withoutSheet(part, sheet);
    if (ref === null) return null;
    if (/^\d{1,7}:\d{1,7}$/.test(ref)) rows = ref;
    else if (/^[A-Za-z]{1,3}:[A-Za-z]{1,3}$/.test(ref)) columns = ref.toUpperCase();
    else return null;
  }
  return rows === null && columns === null ? null : { rows, columns };
}

/** @internal What a worksheet's print and protection elements hold, as read. */
export interface ReadPrintData {
  readonly protection: SheetProtection | null;
  readonly pageSetup: PageSetup;
  readonly headerFooter: Required<HeaderFooter> | null;
  readonly rowBreaks: readonly number[];
  readonly colBreaks: readonly number[];
}

/** @internal */
export const EMPTY_READ_PRINT: ReadPrintData = Object.freeze({ protection: null, pageSetup: DEFAULT_PAGE_SETUP, headerFooter: null, rowBreaks: [], colBreaks: [] });

const HF_TEXTS: ReadonlySet<string> = new Set(['oddHeader', 'oddFooter', 'evenHeader', 'evenFooter', 'firstHeader', 'firstFooter']);
/** Breaks read per collection; the file's own limit is about a thousand ([MS-OI29500] §2.1.599). */
const MAX_READ_BREAKS = 65_536;

/**
 * @internal Collects `sheetProtection`, `printOptions`, `pageMargins`, `pageSetup`,
 * `sheetPr/pageSetUpPr`, `headerFooter`, `rowBreaks` and `colBreaks` of a worksheet while it is
 * parsed. `start` returns `true` for elements it handled. Invalid values warn and keep the default.
 */
export function createPrintCollector(warn: (message: string) => void): {
  start(depth: number, local: string, parent: string | undefined, attrs: ReadonlyMap<string, string>): boolean;
  end(local: string): void;
  text(value: string): boolean;
  result(): ReadPrintData;
} {
  let protection: SheetProtection | null = null;
  const setup: { -readonly [K in keyof PageSetup]: PageSetup[K] } = { ...DEFAULT_PAGE_SETUP };
  let headerFooter: Record<string, string | boolean> | null = null;
  let textTarget: string | null = null;
  let text = '';
  const rowBreaks: number[] = [];
  const colBreaks: number[] = [];
  let breaks: number[] | null = null;

  const boolAttr = (attrs: ReadonlyMap<string, string>, name: string, fallback: boolean, what: string): boolean => {
    const v = attrs.get(name);
    if (v === undefined) return fallback;
    if (v === '1' || v === 'true') return true;
    if (v === '0' || v === 'false') return false;
    warn(`The ${what} attribute ${name}="${v.slice(0, 20)}" is not a boolean; the default is used.`);
    return fallback;
  };
  const uintAttr = (attrs: ReadonlyMap<string, string>, name: string, what: string): number | null => {
    const v = attrs.get(name);
    if (v === undefined) return null;
    if (/^\d{1,10}$/.test(v.trim()) && Number(v) <= 4294967295) return Number(v);
    warn(`The ${what} attribute ${name}="${v.slice(0, 20)}" is not a whole number; the default is used.`);
    return null;
  };
  const textAttr = (attrs: ReadonlyMap<string, string>, name: string): string | null => {
    const v = attrs.get(name);
    return v === undefined ? null : v;
  };
  const enumAttr = <T extends string>(attrs: ReadonlyMap<string, string>, name: string, allowed: readonly T[], fallback: T): T => {
    const v = attrs.get(name);
    if (v === undefined) return fallback;
    if ((allowed as readonly string[]).includes(v)) return v as T;
    warn(`The page setup attribute ${name}="${v.slice(0, 30)}" is not known; the default is used.`);
    return fallback;
  };

  return {
    start(depth, local, parent, attrs) {
      if (depth === 2 && local === 'sheetProtection') {
        const flags: Record<string, boolean> = {};
        for (const k of PROTECTION_FLAGS) flags[k] = boolAttr(attrs, k, DEFAULT_PROTECTION[k], 'sheet protection');
        protection = Object.freeze({
          ...(flags as unknown as Required<SheetProtectionFlags>),
          sheet: boolAttr(attrs, 'sheet', false, 'sheet protection'),
          algorithmName: textAttr(attrs, 'algorithmName'),
          hashValue: textAttr(attrs, 'hashValue'),
          saltValue: textAttr(attrs, 'saltValue'),
          spinCount: uintAttr(attrs, 'spinCount', 'sheet protection'),
          password: textAttr(attrs, 'password'),
        });
        return true;
      }
      if (depth === 2 && local === 'printOptions') {
        setup.printOptions = Object.freeze({
          gridLines: boolAttr(attrs, 'gridLines', false, 'print options') && boolAttr(attrs, 'gridLinesSet', true, 'print options'),
          headings: boolAttr(attrs, 'headings', false, 'print options'),
          horizontalCentered: boolAttr(attrs, 'horizontalCentered', false, 'print options'),
          verticalCentered: boolAttr(attrs, 'verticalCentered', false, 'print options'),
        });
        return true;
      }
      if (depth === 2 && local === 'pageMargins') {
        const m: Record<string, number> = {};
        for (const k of ['left', 'right', 'top', 'bottom', 'header', 'footer'] as const) {
          const v = attrs.get(k);
          const n = v === undefined ? NaN : Number(v);
          if (v === undefined || v.trim() === '' || !Number.isFinite(n) || n < 0) {
            warn(`The page margin ${k}="${(v ?? '').slice(0, 20)}" is missing or not a valid number; the margins are ignored.`);
            return true;
          }
          m[k] = n;
        }
        setup.margins = Object.freeze(m as unknown as PageMargins);
        return true;
      }
      if (depth === 2 && local === 'pageSetup') {
        const what = 'page setup';
        setup.paperSize = uintAttr(attrs, 'paperSize', what) ?? 1;
        setup.scale = uintAttr(attrs, 'scale', what) ?? 100;
        setup.fitToWidth = uintAttr(attrs, 'fitToWidth', what) ?? 1;
        setup.fitToHeight = uintAttr(attrs, 'fitToHeight', what) ?? 1;
        setup.pageOrder = enumAttr(attrs, 'pageOrder', ORDERS, 'downThenOver');
        setup.orientation = enumAttr(attrs, 'orientation', ORIENTATIONS, 'default');
        setup.blackAndWhite = boolAttr(attrs, 'blackAndWhite', false, what);
        setup.draft = boolAttr(attrs, 'draft', false, what);
        setup.cellComments = enumAttr(attrs, 'cellComments', COMMENTS, 'none');
        setup.errors = enumAttr(attrs, 'errors', ERRORS, 'displayed');
        setup.firstPageNumber = boolAttr(attrs, 'useFirstPageNumber', false, what) ? (uintAttr(attrs, 'firstPageNumber', what) ?? 1) : null;
        setup.horizontalDpi = uintAttr(attrs, 'horizontalDpi', what);
        setup.verticalDpi = uintAttr(attrs, 'verticalDpi', what);
        setup.copies = uintAttr(attrs, 'copies', what) ?? 1;
        return true;
      }
      if (depth === 3 && local === 'pageSetUpPr' && parent === 'sheetPr') {
        setup.fitToPage = boolAttr(attrs, 'fitToPage', false, 'page setup');
        return true;
      }
      if (depth === 2 && local === 'headerFooter') {
        headerFooter = {
          ...DEFAULT_HEADER_FOOTER,
          differentOddEven: boolAttr(attrs, 'differentOddEven', false, 'header/footer'),
          differentFirst: boolAttr(attrs, 'differentFirst', false, 'header/footer'),
          scaleWithDoc: boolAttr(attrs, 'scaleWithDoc', true, 'header/footer'),
          alignWithMargins: boolAttr(attrs, 'alignWithMargins', true, 'header/footer'),
        };
        return true;
      }
      if (depth === 3 && parent === 'headerFooter' && HF_TEXTS.has(local)) {
        textTarget = local;
        text = '';
        return true;
      }
      if (depth === 2 && (local === 'rowBreaks' || local === 'colBreaks')) {
        breaks = local === 'rowBreaks' ? rowBreaks : colBreaks;
        return true;
      }
      if (depth === 3 && local === 'brk' && (parent === 'rowBreaks' || parent === 'colBreaks') && breaks !== null) {
        const id = uintAttr(attrs, 'id', 'page break') ?? 0;
        const limit = parent === 'rowBreaks' ? MAX_ROWS : MAX_COLUMNS;
        if (id + 1 > limit) warn(`The page break id="${String(id)}" is beyond the sheet; it is ignored.`);
        else if (breaks.length >= MAX_READ_BREAKS) warn(`The sheet has more than ${String(MAX_READ_BREAKS)} page breaks; the rest are ignored.`);
        else breaks.push(id + 1);
        return true;
      }
      return false;
    },
    end(local) {
      if (textTarget !== null && local === textTarget) {
        if (headerFooter !== null) headerFooter[textTarget] = decodeXstring(text);
        textTarget = null;
        text = '';
      } else if (local === 'rowBreaks' || local === 'colBreaks') breaks = null;
    },
    text(value) {
      if (textTarget === null) return false;
      if (text.length < 1_048_576) text += value;
      return true;
    },
    result() {
      const unique = (list: number[]): number[] => [...new Set(list)].sort((x, y) => x - y);
      return {
        protection,
        pageSetup: Object.freeze({ ...setup }),
        headerFooter: headerFooter === null ? null : (Object.freeze(headerFooter) as unknown as Required<HeaderFooter>),
        rowBreaks: unique(rowBreaks),
        colBreaks: unique(colBreaks),
      };
    },
  };
}

/** @internal Reads the `workbookProtection` attributes (SR-22). */
export function workbookProtectionOf(attrs: ReadonlyMap<string, string>, warn: (message: string) => void): WorkbookProtection {
  const b = (name: string): boolean => {
    const v = attrs.get(name);
    if (v === undefined) return false;
    if (v === '1' || v === 'true') return true;
    if (v !== '0' && v !== 'false') warn(`The workbook protection attribute ${name}="${v.slice(0, 20)}" is not a boolean; false is used.`);
    return false;
  };
  const t = (name: string): string | null => attrs.get(name) ?? null;
  const n = (name: string): number | null => {
    const v = attrs.get(name);
    return v !== undefined && /^\d{1,10}$/.test(v.trim()) && Number(v) <= 4294967295 ? Number(v) : null;
  };
  return Object.freeze({
    lockStructure: b('lockStructure'),
    lockWindows: b('lockWindows'),
    lockRevision: b('lockRevision'),
    workbookAlgorithmName: t('workbookAlgorithmName'),
    workbookHashValue: t('workbookHashValue'),
    workbookSaltValue: t('workbookSaltValue'),
    workbookSpinCount: n('workbookSpinCount'),
    revisionsAlgorithmName: t('revisionsAlgorithmName'),
    revisionsHashValue: t('revisionsHashValue'),
    revisionsSaltValue: t('revisionsSaltValue'),
    revisionsSpinCount: n('revisionsSpinCount'),
    workbookPassword: t('workbookPassword'),
    revisionsPassword: t('revisionsPassword'),
  });
}
