import type { DefinedName } from './defined-names';
import { isAppPropertiesRel, isCorePropertiesRel, parseAppProperties, parseCoreProperties, type WorkbookProperties } from './doc-properties';
import { XlvError, type XlvWarning } from './errors';
import { isUnusualTarget, type Hyperlink } from './hyperlinks';
import { DEFAULT_PAGE_SETUP, parsePrintArea, parsePrintTitles, type HeaderFooter, type PageSetup, type SheetProtection, type WorkbookProtection } from './printing';
import { resolveLimits, type ReadLimits } from './limits';
import { parseComments, THREADED_COMMENTS_REL, type Note } from './notes';
import type { ConditionalFormat, DataValidation } from './sheet-rules';
import type { Cell, ColumnInfo, Conformance, DateSystem, FrozenPanes, RichTextRun, RowInfo, Sheet, SheetKind, SheetState, SheetView, Workbook, WorkbookFormat } from './model';
import type { Color } from './style';
import { openPackage, type OpcPackage, type Relationship } from './opc/package';
import { parseCellRef, parseRange } from './sml/cell-ref';
import {
  classifyRelationship,
  INTL_MACROSHEET_REL,
  MACROSHEET_CONTENT_TYPES,
  MACROSHEET_REL,
  VBA_PROJECT_CONTENT_TYPE,
  VBA_PROJECT_REL,
} from './sml/namespaces';
import { parseSharedStrings } from './sml/shared-strings';
import { EMPTY_STYLES, parseStyles, type Styles } from './sml/styles';
import { parseWorkbookPart } from './sml/workbook';
import { DEFAULT_SHEET_VIEW, parseWorksheet, type NumberFormat, type WorksheetData } from './sml/worksheet';
import { suggest } from './util/suggest';
import { WarningCollector } from './util/warnings';
import { openZip } from './zip/reader';

/**
 * Options for {@link readWorkbook}.
 *
 * @example
 * ```ts
 * import { readWorkbook, type ReadOptions } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const options: ReadOptions = { limits: { maxCells: 1_000_000 } };
 * console.log(readWorkbook(bytes, options).sheets.length);
 * ```
 */
export interface ReadOptions {
  /** Resource limits; see {@link ReadLimits}. Unspecified limits keep their defaults. */
  readonly limits?: Partial<ReadLimits>;
}

const OPTION_NAMES = ['limits'];

/** Content types of the workbook part (ECMA-376 Part 1 §12.3.23; [MS-OFFMACRO2] §2.2.1.4). */
const FORMATS: ReadonlyMap<string, WorkbookFormat> = new Map([
  ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml', 'xlsx'],
  ['application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml', 'xltx'],
  ['application/vnd.ms-excel.sheet.macroEnabled.main+xml', 'xlsm'],
  ['application/vnd.ms-excel.template.macroEnabled.main+xml', 'xltm'],
]);

class ReadSheet implements Sheet {
  private sortedRows: (readonly Cell[])[] | null = null;

  constructor(
    readonly name: string,
    readonly index: number,
    readonly kind: SheetKind,
    readonly state: SheetState,
    readonly part: string | null,
    private readonly data: WorksheetData | null,
    readonly hyperlinks: readonly Hyperlink[] = [],
    readonly notes: readonly Note[] = [],
    private readonly names: readonly DefinedName[] = [],
  ) {}

  get protection(): SheetProtection | null {
    return this.data?.print.protection ?? null;
  }

  get pageSetup(): PageSetup {
    return this.data?.print.pageSetup ?? DEFAULT_PAGE_SETUP;
  }

  get headerFooter(): Required<HeaderFooter> | null {
    return this.data?.print.headerFooter ?? null;
  }

  get pageBreaks(): { readonly rows: readonly number[]; readonly columns: readonly number[] } {
    return { rows: this.data?.print.rowBreaks ?? [], columns: this.data?.print.colBreaks ?? [] };
  }

  get printArea(): readonly string[] | null {
    const n = this.names.find((d) => d.sheet === this.name && d.name.toLowerCase() === '_xlnm.print_area');
    return n === undefined ? null : parsePrintArea(n.refersTo, this.name);
  }

  get printTitles(): { readonly rows: string | null; readonly columns: string | null } | null {
    const n = this.names.find((d) => d.sheet === this.name && d.name.toLowerCase() === '_xlnm.print_titles');
    return n === undefined ? null : parsePrintTitles(n.refersTo, this.name);
  }

  get dimension(): string | null {
    return this.data?.dimension ?? null;
  }

  get merges(): readonly string[] {
    return this.data?.merges ?? [];
  }

  get cellCount(): number {
    return this.data?.cellCount ?? 0;
  }

  get columns(): readonly ColumnInfo[] {
    return this.data?.columns ?? [];
  }

  get freeze(): FrozenPanes | null {
    return this.data?.freeze ?? null;
  }

  get autoFilter(): string | null {
    return this.data?.autoFilter ?? null;
  }

  get tabColor(): Color | null {
    return this.data?.tabColor ?? null;
  }

  get dataValidations(): readonly DataValidation[] {
    return this.data?.dataValidations ?? [];
  }

  get conditionalFormats(): readonly ConditionalFormat[] {
    return this.data?.conditionalFormats ?? [];
  }

  get view(): SheetView {
    return this.data?.view ?? DEFAULT_SHEET_VIEW;
  }

  rowInfo(row: number): RowInfo | undefined {
    return this.data?.rowInfo.get(row);
  }

  cell(ref: string): Cell | undefined {
    const address = typeof ref === 'string' ? parseCellRef(ref.toUpperCase()) : null;
    if (address === null) {
      throw new XlvError('XLV_E_INVALID_CELL_REF', `"${String(ref).slice(0, 40)}" is not a valid cell reference.`, {
        location: { sheet: this.name },
        hint: 'Pass an A1-style reference such as "B7" (column letters A to XFD, rows 1 to 1048576).',
        details: { ref: String(ref).slice(0, 40) },
      });
    }
    return this.data?.rows.get(address.row)?.get(address.col);
  }

  *rows(): IterableIterator<readonly Cell[]> {
    if (this.sortedRows === null) {
      const rows = this.data?.rows ?? new Map<number, Map<number, Cell>>();
      this.sortedRows = [...rows.keys()]
        .sort((a, b) => a - b)
        .map((r) => {
          const cells = rows.get(r) ?? new Map<number, Cell>();
          return [...cells.keys()].sort((a, b) => a - b).map((c) => cells.get(c) as Cell);
        });
    }
    yield* this.sortedRows;
  }

  *cells(): IterableIterator<Cell> {
    for (const row of this.rows()) yield* row;
  }
}

class ReadWorkbook implements Workbook {
  constructor(
    readonly sheets: readonly Sheet[],
    readonly dateSystem: '1900' | '1904',
    readonly conformance: Conformance,
    readonly format: WorkbookFormat,
    readonly hasMacros: boolean,
    readonly externalLinkCount: number,
    readonly warnings: readonly XlvWarning[],
    readonly properties: WorkbookProperties,
    readonly definedNames: readonly DefinedName[],
    readonly protection: WorkbookProtection | null = null,
  ) {}

  getSheet(name: string): Sheet {
    const sheet = this.sheets.find((s) => s.name === name);
    if (sheet !== undefined) return sheet;
    const names = this.sheets.map((s) => s.name);
    const close = suggest(String(name), names);
    throw new XlvError('XLV_E_SHEET_NOT_FOUND', `The workbook has no sheet named "${String(name)}".`, {
      location: { sheet: String(name) },
      hint:
        (close.length > 0 ? `Did you mean ${close.map((n) => `"${n}"`).join(' or ')}? ` : '') +
        `Available sheets: ${names.map((n) => `"${n}"`).join(', ') || '(none)'}. Sheet names are case-sensitive.`,
      details: { available: names },
    });
  }
}

function checkOptions(options: unknown): ReadOptions {
  if (options === undefined) return {};
  if (options === null || typeof options !== 'object' || Array.isArray(options)) {
    throw new XlvError('XLV_E_INVALID_OPTION', 'Options must be an object.', {
      hint: 'Pass an object such as { limits: { maxCells: 1000000 } }, or leave the options out.',
    });
  }
  for (const key of Object.keys(options)) {
    if (!OPTION_NAMES.includes(key)) {
      const close = suggest(key, OPTION_NAMES);
      throw new XlvError('XLV_E_INVALID_OPTION', `Unknown option "${key}".`, {
        hint: `${close.length > 0 ? `Did you mean "${close[0] ?? ''}"? ` : ''}Known options: ${OPTION_NAMES.join(', ')}.`,
        details: { option: key, known: OPTION_NAMES },
      });
    }
  }
  return options as ReadOptions;
}

/** Converts the accepted input types to a Uint8Array. */
export function toBytes(data: unknown): Uint8Array {
  if (data instanceof Uint8Array) return data;
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  if (ArrayBuffer.isView(data)) return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  throw new XlvError('XLV_E_INVALID_INPUT', `Expected the file contents as a Uint8Array or ArrayBuffer, got ${data === null ? 'null' : typeof data}.`, {
    hint:
      typeof data === 'string'
        ? 'Pass the bytes of the file, not its path. In Node.js use readWorkbookFile() from "xlvane/node", or read the file with fs.readFile() first.'
        : 'Pass the bytes of the file, for example from fetch(...).arrayBuffer(), File.arrayBuffer() or fs.readFile().',
  });
}

function findRelationship(rels: readonly Relationship[], name: string): Relationship | undefined {
  return rels.find((r) => !r.external && classifyRelationship(r.type)?.name === name);
}

/**
 * Reads an XLSX workbook (also .xlsm, .xltx, .xltm) from memory. The input is treated as
 * untrusted: ZIP and XML limits apply, DTDs are rejected, macros are never run and external
 * links are never followed. Formulas are not calculated; formula cells hold the cached value
 * stored in the file.
 *
 * @example
 * ```ts
 * import { readWorkbook } from 'xlvane';
 *
 * const bytes = new Uint8Array(await (await fetch('/report.xlsx')).arrayBuffer());
 * const workbook = readWorkbook(bytes);
 * const sheet = workbook.getSheet('Sales');
 * for (const row of sheet.rows()) console.log(row.map((cell) => cell.value));
 * ```
 */
export function readWorkbook(data: Uint8Array | ArrayBuffer, options?: ReadOptions): Workbook {
  const bytes = toBytes(data);
  const limits = resolveLimits(checkOptions(options).limits);
  const zip = openZip(bytes, limits);
  const pkg = openPackage(zip, limits);
  const warnings = new WarningCollector();
  const book = analyzeWorkbook(pkg, limits, warnings);
  const budget = { remaining: limits.maxCells, sharedText: limits.maxSharedStringChars };
  const formatCache = new Map<number, NumberFormat>();
  const sheets = book.sheets.map((entry): Sheet => {
    let data: WorksheetData | null = null;
    if (entry.kind === 'worksheet' && entry.part !== null) {
      data = parseWorksheet(pkg.readXml(entry.part) ?? '', {
        part: entry.part,
        sheet: entry.name,
        limits,
        warnings,
        sharedStrings: book.sharedStrings,
        richStrings: book.richStrings,
        styles: book.styles,
        dateSystem: book.dateSystem,
        conformance: book.conformance,
        budget,
        formatCache,
      });
    }
    const links = data === null || entry.part === null ? [] : resolveHyperlinks(pkg, entry.part, entry.name, data, warnings);
    const notes = data === null || entry.part === null ? [] : readNotes(pkg, entry.part, entry.name, limits, warnings);
    if (data !== null) attachNotes(data.rows, notes);
    return new ReadSheet(entry.name, entry.index, entry.kind, entry.state, entry.part, data, links, notes, book.definedNames);
  });
  return new ReadWorkbook(sheets, book.dateSystem, book.conformance, book.format, book.hasMacros, book.externalLinkCount, warnings.list(), book.properties, book.definedNames, book.protection);
}

/**
 * Resolves the `r:id` of each hyperlink against the sheet's relationships (W2-42, W2-44) and
 * attaches the links to the cells they cover. Targets are only read as text, never opened.
 */
function resolveHyperlinks(pkg: OpcPackage, part: string, sheet: string, data: WorksheetData, warnings: WarningCollector): Hyperlink[] {
  if (data.hyperlinks.length === 0) return [];
  const rels = pkg.relationships(part);
  const links: Hyperlink[] = [];
  let unusual: string | null = null;
  for (const raw of data.hyperlinks) {
    let target: string | null = null;
    if (raw.relId !== null) {
      const rel = rels.find((r) => r.id === raw.relId);
      if (rel === undefined || classifyRelationship(rel.type)?.name !== 'hyperlink') {
        warnings.add('XLV_W_INVALID_VALUE', `The hyperlink of ${raw.ref} refers to relationship "${raw.relId.slice(0, 40)}", which is not a hyperlink relationship of the sheet; its target is unknown.`, { part, sheet, cell: raw.ref });
      } else {
        target = rel.target;
        if (unusual === null && isUnusualTarget(target)) unusual = target;
      }
    }
    links.push(Object.freeze({ ref: raw.ref, target, location: raw.location, tooltip: raw.tooltip, display: raw.display }));
  }
  if (unusual !== null) {
    warnings.add('XLV_W_UNUSUAL_HYPERLINK', `The sheet has hyperlinks to targets other than web pages or e-mail addresses, for example "${unusual.slice(0, 80)}". xlvane does not open them.`, { part, sheet });
  }
  // Each cell gets the first link that covers it. Per link, the smaller of the range and the
  // sheet's rows (and cells) is walked, so a single-cell link costs O(1).
  for (const link of links) {
    const range = parseRange(link.ref);
    if (range === null) continue;
    const rowSpan = range.end.row - range.start.row + 1;
    const colSpan = range.end.col - range.start.col + 1;
    const visit = (cells: Map<number, Cell>): void => {
      const mark = (c: number, cell: Cell): void => {
        if (cell.hyperlink === undefined) cells.set(c, { ...cell, hyperlink: link });
      };
      if (colSpan <= cells.size) {
        for (let c = range.start.col; c <= range.end.col; c++) {
          const cell = cells.get(c);
          if (cell) mark(c, cell);
        }
      } else for (const [c, cell] of cells) if (c >= range.start.col && c <= range.end.col) mark(c, cell);
    };
    if (rowSpan <= data.rows.size) {
      for (let r = range.start.row; r <= range.end.row; r++) {
        const cells = data.rows.get(r);
        if (cells) visit(cells);
      }
    } else for (const [r, cells] of data.rows) if (r >= range.start.row && r <= range.end.row) visit(cells);
  }
  return links;
}

/**
 * The notes of a worksheet: the Comments part named by an internal relationship of the sheet
 * (W2-46), read as text; the VML part that draws them is not read. Threaded comments
 * ([MS-XLSX] §2.3.7) are reported with `XLV_W_THREADED_COMMENTS`; their legacy placeholders are
 * ordinary notes.
 */
export function readNotes(pkg: OpcPackage, part: string, sheet: string, limits: ReadLimits, warnings: WarningCollector): Note[] {
  const rels = pkg.relationships(part);
  if (rels.some((r) => r.type === THREADED_COMMENTS_REL)) {
    warnings.add('XLV_W_THREADED_COMMENTS', 'The sheet has threaded comments (conversations); xlvane reads only the notes stored with them, whose author is "tc={id}" and whose text may not include the replies.', { part, sheet });
  }
  const rel = rels.find((r) => !r.external && classifyRelationship(r.type)?.name === 'comments');
  if (rel === undefined) return [];
  const xml = pkg.readXml(rel.target);
  if (xml === undefined) {
    warnings.add('XLV_W_PART_MISSING', 'The sheet refers to a comments part that does not exist; its notes are not read.', { part: rel.target, sheet });
    return [];
  }
  return parseComments(xml, { part: rel.target, sheet, limits, warnings });
}

/** Sets `note` on the existing cells that have one (the first note of a cell wins). */
export function attachNotes(rows: Map<number, Map<number, Cell>>, notes: readonly Note[]): void {
  for (const note of notes) {
    const address = parseCellRef(note.ref);
    if (address === null) continue;
    const cells = rows.get(address.row);
    const cell = cells?.get(address.col);
    if (cells && cell && cell.note === undefined) cells.set(address.col, { ...cell, note });
  }
}

/** A sheet as listed in the workbook part, before its cells are read. */
export interface SheetEntry {
  readonly name: string;
  readonly index: number;
  readonly kind: SheetKind;
  readonly state: SheetState;
  /** Part name, or `null` when the part is missing (reported as XLV_W_SHEET_PART_MISSING). */
  readonly part: string | null;
}

/** Everything about a workbook except the cells; shared by the batch and streaming readers. */
export interface WorkbookInfo {
  readonly workbookPart: string;
  readonly sheets: readonly SheetEntry[];
  readonly sharedStrings: readonly string[];
  /** Runs of the shared strings that are rich text, by index. */
  readonly richStrings: ReadonlyMap<number, readonly RichTextRun[]>;
  readonly styles: Styles;
  readonly dateSystem: DateSystem;
  readonly conformance: Conformance;
  readonly format: WorkbookFormat;
  readonly hasMacros: boolean;
  readonly externalLinkCount: number;
  readonly properties: WorkbookProperties;
  readonly definedNames: readonly DefinedName[];
  readonly protection: WorkbookProtection | null;
}

/** The Core Properties relationship of the package, if any (D-A02). */
function corePropertiesRel(pkg: OpcPackage): Relationship | undefined {
  return pkg.relationships(null).find((r) => !r.external && isCorePropertiesRel(r.type));
}

/** The Extended File Properties relationship of the package, if any (D-A10). */
function appPropertiesRel(pkg: OpcPackage): Relationship | undefined {
  return pkg.relationships(null).find((r) => !r.external && isAppPropertiesRel(r.type));
}

/** The workbook part's relationships that the reader loads in full (workbook, styles, shared strings). */
export function workbookPartsToLoad(pkg: OpcPackage): string[] {
  const officeDocument = findRelationship(pkg.relationships(null), 'officeDocument');
  if (officeDocument === undefined) return [];
  const rels = pkg.relationships(officeDocument.target);
  const out = [officeDocument.target];
  for (const name of ['sharedStrings', 'styles']) {
    const rel = findRelationship(rels, name);
    if (rel !== undefined) out.push(rel.target);
  }
  const core = corePropertiesRel(pkg);
  if (core !== undefined) out.push(core.target);
  const app = appPropertiesRel(pkg);
  if (app !== undefined) out.push(app.target);
  return out;
}

/** Reads the workbook structure, shared strings and styles (D-401). */
export function analyzeWorkbook(pkg: OpcPackage, limits: ReadLimits, warnings: WarningCollector): WorkbookInfo {
  const officeDocument = findRelationship(pkg.relationships(null), 'officeDocument');
  if (officeDocument === undefined || !pkg.hasPart(officeDocument.target)) {
    throw new XlvError('XLV_E_PART_MISSING', 'The package has no workbook part.', {
      location: officeDocument === undefined ? { part: '/_rels/.rels' } : { part: officeDocument.target },
      hint: 'This is not a complete spreadsheet: the package does not point to a workbook. Check that the file is an Excel workbook and re-save it in a spreadsheet application.',
    });
  }
  const workbookPart = officeDocument.target;
  const format = FORMATS.get(pkg.contentType(workbookPart) ?? '') ?? 'unknown';
  const info = parseWorkbookPart(pkg.readXml(workbookPart) ?? '', { part: workbookPart, limits, warnings });
  const rels = pkg.relationships(workbookPart);
  const relsById = new Map(rels.map((r) => [r.id, r]));
  const context = { limits, warnings };

  const readOptionalPart = (name: 'sharedStrings' | 'styles', what: string): { part: string; xml: string } | undefined => {
    const rel = findRelationship(rels, name);
    if (rel === undefined) return undefined;
    const xml = pkg.readXml(rel.target);
    if (xml === undefined) {
      warnings.add('XLV_W_PART_MISSING', `The workbook refers to a ${what} part that does not exist; reading continues without it.`, {
        part: rel.target,
      });
      return undefined;
    }
    return { part: rel.target, xml };
  };
  const sst = readOptionalPart('sharedStrings', 'shared strings');
  const richStrings = new Map<number, readonly RichTextRun[]>();
  const sharedStrings = sst === undefined ? [] : parseSharedStrings(sst.xml, { part: sst.part, ...context }, richStrings);
  const stylesPart = readOptionalPart('styles', 'styles');
  const styles: Styles = stylesPart === undefined ? EMPTY_STYLES : parseStyles(stylesPart.xml, { part: stylesPart.part, ...context });

  const hasMacros = detectMacros(pkg, rels);
  if (hasMacros) {
    warnings.add('XLV_W_MACROS_PRESENT', 'The workbook contains macros (a VBA project or Excel 4.0 macro sheets). xlvane never runs them.', {
      part: workbookPart,
    });
  }
  const externalLinkCount = rels.filter((r) => classifyRelationship(r.type)?.name === 'externalLink').length;
  if (externalLinkCount > 0) {
    warnings.add('XLV_W_EXTERNAL_LINKS', `The workbook links to ${externalLinkCount} external workbook(s). xlvane never follows the links; linked values are the cached copies.`, {
      part: workbookPart,
    });
  }

  const dateSystem = info.date1904 ? '1904' : '1900';
  const sheets = info.sheets.map((entry, index): SheetEntry => {
    if (entry.state === 'veryHidden') {
      warnings.add('XLV_W_VERY_HIDDEN_SHEET', `Sheet "${entry.name}" is very hidden: spreadsheet applications do not offer a way to unhide it.`, {
        sheet: entry.name,
      });
    }
    const rel = entry.relId === null ? undefined : relsById.get(entry.relId);
    const relName = rel === undefined ? undefined : classifyRelationship(rel.type)?.name;
    const kind: SheetKind =
      rel !== undefined && (rel.type === MACROSHEET_REL || rel.type === INTL_MACROSHEET_REL)
        ? 'macrosheet'
        : relName === 'chartsheet' || relName === 'dialogsheet'
          ? relName
          : 'worksheet';
    if (rel === undefined || rel.external || !pkg.hasPart(rel.target)) {
      warnings.add('XLV_W_SHEET_PART_MISSING', `The part of sheet "${entry.name}" is missing; the sheet is returned without cells.`, {
        sheet: entry.name,
        part: workbookPart,
      });
      return { name: entry.name, index, kind, state: entry.state, part: null };
    }
    return { name: entry.name, index, kind, state: entry.state, part: rel.target };
  });
  const definedNames = info.definedNames.map((n): DefinedName => {
    let sheet: string | null = null;
    if (n.localSheetId !== null) {
      const entry = info.sheets[n.localSheetId];
      if (entry === undefined) {
        warnings.add('XLV_W_INVALID_VALUE', `The defined name "${n.name.slice(0, 100)}" belongs to sheet number ${String(n.localSheetId)}, which does not exist; it is read as a workbook-level name.`, { part: workbookPart });
      } else sheet = entry.name;
    }
    return { name: n.name, refersTo: n.refersTo, sheet, hidden: n.hidden, comment: n.comment };
  });
  let properties: WorkbookProperties = {};
  const core = corePropertiesRel(pkg);
  if (core !== undefined) {
    const xml = pkg.readXml(core.target);
    if (xml === undefined) warnings.add('XLV_W_PART_MISSING', 'The package refers to a document properties part that does not exist; reading continues without it.', { part: core.target });
    else properties = parseCoreProperties(xml, { part: core.target, ...context });
  }
  const app = appPropertiesRel(pkg);
  if (app !== undefined) {
    const xml = pkg.readXml(app.target);
    if (xml === undefined) warnings.add('XLV_W_PART_MISSING', 'The package refers to an extended document properties part that does not exist; reading continues without it.', { part: app.target });
    else properties = { ...properties, ...parseAppProperties(xml, { part: app.target, ...context }) };
  }
  return { workbookPart, sheets, sharedStrings, richStrings, styles, dateSystem, conformance: info.conformance, format, hasMacros, externalLinkCount, properties, definedNames, protection: info.protection };
}

function detectMacros(pkg: OpcPackage, workbookRels: readonly Relationship[]): boolean {
  if (workbookRels.some((r) => r.type === VBA_PROJECT_REL || r.type === MACROSHEET_REL || r.type === INTL_MACROSHEET_REL)) {
    return true;
  }
  return pkg.partNames.some((name) => {
    const type = pkg.contentType(name);
    return type === VBA_PROJECT_CONTENT_TYPE || (type !== undefined && MACROSHEET_CONTENT_TYPES.includes(type));
  });
}
