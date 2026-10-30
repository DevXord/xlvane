import { XlvError, type XlvWarning } from './errors';
import { resolveLimits, type ReadLimits } from './limits';
import type { Cell, ColumnInfo, Conformance, DateSystem, FrozenPanes, RowInfo, Sheet, SheetKind, SheetState, Workbook, WorkbookFormat } from './model';
import { openPackage, type OpcPackage, type Relationship } from './opc/package';
import { parseCellRef } from './sml/cell-ref';
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
import { parseWorksheet, type NumberFormat, type WorksheetData } from './sml/worksheet';
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
  ) {}

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
        styles: book.styles,
        dateSystem: book.dateSystem,
        conformance: book.conformance,
        budget,
        formatCache,
      });
    }
    return new ReadSheet(entry.name, entry.index, entry.kind, entry.state, entry.part, data);
  });
  return new ReadWorkbook(sheets, book.dateSystem, book.conformance, book.format, book.hasMacros, book.externalLinkCount, warnings.list());
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
  readonly styles: Styles;
  readonly dateSystem: DateSystem;
  readonly conformance: Conformance;
  readonly format: WorkbookFormat;
  readonly hasMacros: boolean;
  readonly externalLinkCount: number;
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
  const sharedStrings = sst === undefined ? [] : parseSharedStrings(sst.xml, { part: sst.part, ...context });
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
  return { workbookPart, sheets, sharedStrings, styles, dateSystem, conformance: info.conformance, format, hasMacros, externalLinkCount };
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
