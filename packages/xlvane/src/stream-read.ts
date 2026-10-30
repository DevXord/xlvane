// Streaming reader: worksheets are decompressed and parsed incrementally, and each row is handed
// on and then dropped, so memory does not grow with the number of rows (docs/decisions.md
// D-401, D-402). The workbook part, shared strings and styles are read in full first.
import { XlvError, type XlvWarning } from './errors';
import { detectHeader } from './inspect';
import { resolveLimits, type ReadLimits } from './limits';
import type { Cell, Conformance, DateSystem, SheetKind, SheetState, WorkbookFormat } from './model';
import { openPackage } from './opc/package';
import { analyzeWorkbook, workbookPartsToLoad, type WorkbookInfo } from './read';
import { columnLetters, formatCellRef } from './sml/cell-ref';
import type { NumberFormat } from './sml/worksheet';
import { createWorksheetParser } from './sml/worksheet';
import type { StandardSchemaV1 } from './standard-schema';
import type { TableColumn, TableIssue, TableRow, TableValue } from './table';
import { oneOfOption, readOptions } from './util/options';
import { suggest } from './util/suggest';
import { WarningCollector } from './util/warnings';
import { createXmlByteStreamParser } from './xml/parser';
import type { ZipEntry, ZipReader } from './zip/reader';
import { openZipSource, type StreamZip } from './zip/stream';

/**
 * One row from {@link WorkbookStream.rows}.
 *
 * @example
 * ```ts
 * import { openWorkbookStream, type StreamRow } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const book = await openWorkbookStream(bytes);
 * const show = (row: StreamRow) => `${String(row.number)}: ${row.cells.map((c) => String(c.value)).join(', ')}`;
 * for await (const row of book.rows('Data')) console.log(show(row));
 * ```
 */
export interface StreamRow {
  /** 1-based row number. */
  readonly number: number;
  /** Cells with a value or formula, in column order. */
  readonly cells: readonly Cell[];
  /** Row height in points, or `null` if the file sets none. */
  readonly height: number | null;
  readonly hidden: boolean;
}

/**
 * A sheet listed by {@link WorkbookStream}.
 *
 * @example
 * ```ts
 * import { openWorkbookStream, type StreamSheetInfo } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const visible: StreamSheetInfo[] = (await openWorkbookStream(bytes)).sheets.filter((s) => s.state === 'visible');
 * ```
 */
export interface StreamSheetInfo {
  readonly name: string;
  readonly index: number;
  readonly kind: SheetKind;
  readonly state: SheetState;
}

/**
 * A workbook opened for streaming with {@link openWorkbookStream}.
 *
 * @example
 * ```ts
 * import { openWorkbookStream, type WorkbookStream } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const book: WorkbookStream = await openWorkbookStream(bytes);
 * let count = 0;
 * for await (const row of book.rows(book.sheets[0]!.name)) count += row.cells.length;
 * console.log(count, book.warnings.length);
 * ```
 */
export interface WorkbookStream {
  readonly sheets: readonly StreamSheetInfo[];
  readonly dateSystem: DateSystem;
  readonly conformance: Conformance;
  readonly format: WorkbookFormat;
  readonly hasMacros: boolean;
  readonly externalLinkCount: number;
  /** Warnings so far; rows read later can add more. */
  readonly warnings: readonly XlvWarning[];
  /**
   * The rows of a worksheet in file order, read lazily. Can be called again to read the sheet
   * again. Throws `XLV_E_SHEET_NOT_FOUND` for an unknown name (exact match).
   */
  rows(sheet: string): AsyncIterableIterator<StreamRow>;
  /** The same rows as a `ReadableStream`. */
  rowStream(sheet: string): ReadableStream<StreamRow>;
}

/**
 * Options for {@link openWorkbookStream}.
 *
 * @example
 * ```ts
 * import type { OpenWorkbookStreamOptions } from 'xlvane';
 *
 * const options: OpenWorkbookStreamOptions = { limits: { maxEntryBytes: 8 * 1024 ** 3 } };
 * ```
 */
export interface OpenWorkbookStreamOptions {
  /**
   * Read limits. Streaming defaults differ from `DEFAULT_READ_LIMITS`: `maxCells` does not apply
   * (memory does not grow with cells), `maxEntryBytes` is 4 GiB and `maxTotalBytes` 8 GiB.
   */
  readonly limits?: Partial<ReadLimits>;
}

const GiB = 1024 * 1024 * 1024;
/**
 * Limits that differ in streaming mode (D-401): `maxCells` does not apply, and single parts may
 * decompress to 4 GiB (8 GiB in total). All other `DEFAULT_READ_LIMITS` stay.
 *
 * @example
 * ```ts
 * import { STREAM_LIMIT_DEFAULTS } from 'xlvane';
 *
 * console.log(STREAM_LIMIT_DEFAULTS.maxEntryBytes); // 4294967296
 * ```
 */
export const STREAM_LIMIT_DEFAULTS: Partial<ReadLimits> = Object.freeze({ maxCells: Infinity, maxEntryBytes: 4 * GiB, maxTotalBytes: 8 * GiB });

/** A ZipReader over entries loaded in advance, so the OPC layer can stay synchronous. */
class PreloadedZip implements ZipReader {
  readonly entries: readonly ZipEntry[];
  bytesRead = 0;
  private readonly data = new Map<string, Uint8Array>();

  constructor(private readonly zip: StreamZip) {
    this.entries = zip.entries;
  }

  async load(names: Iterable<string>): Promise<void> {
    for (const name of names) {
      const entry = this.zip.find(name);
      if (entry === undefined || this.data.has(entry.name)) continue;
      const bytes = await this.zip.read(entry);
      this.data.set(entry.name, bytes);
      this.bytesRead += bytes.length;
    }
  }

  find(name: string): ZipEntry | undefined {
    return this.zip.find(name);
  }

  read(entry: ZipEntry | string): Uint8Array {
    const name = typeof entry === 'string' ? (this.zip.find(entry)?.name ?? entry) : entry.name;
    const bytes = this.data.get(name);
    if (bytes === undefined) throw new Error(`internal: ZIP entry "${name}" was not preloaded`);
    return bytes;
  }
}

function zipName(partName: string): string {
  return partName.slice(1);
}

function checkSource(source: unknown): Uint8Array | ArrayBuffer | Blob {
  if (source instanceof Uint8Array || source instanceof ArrayBuffer) return source;
  if (ArrayBuffer.isView(source)) return new Uint8Array(source.buffer, source.byteOffset, source.byteLength);
  if (typeof Blob !== 'undefined' && source instanceof Blob) return source;
  throw new XlvError('XLV_E_INVALID_INPUT', `Expected the file as a Uint8Array, ArrayBuffer or Blob, got ${source === null ? 'null' : typeof source}.`, {
    hint:
      typeof source === 'string'
        ? 'Pass the file, not its path. In Node.js use openWorkbookStreamFile() from "xlvane/node", or fs.openAsBlob(path).'
        : 'Pass the bytes, or a Blob/File (for example from an <input type="file"> or fs.openAsBlob()).',
  });
}

/**
 * Opens a workbook for reading row by row. Reads the workbook structure, shared strings and
 * styles; worksheet cells are decompressed and parsed only while their rows are iterated, one
 * row at a time. With a `Blob` (for example `fs.openAsBlob()` in Node.js, or a `File` in the
 * browser) the file itself is never held in memory either. Untrusted-input limits apply as in
 * `readWorkbook`, except `maxCells` (see {@link OpenWorkbookStreamOptions}).
 *
 * @example
 * ```ts
 * import { openWorkbookStream } from 'xlvane';
 *
 * declare const file: Blob;
 * const book = await openWorkbookStream(file);
 * let total = 0;
 * for await (const row of book.rows('Orders')) {
 *   const amount = row.cells.find((cell) => cell.col === 3)?.value;
 *   if (typeof amount === 'number') total += amount;
 * }
 * console.log(total);
 * ```
 */
export async function openWorkbookStream(source: Uint8Array | ArrayBuffer | Blob, options?: OpenWorkbookStreamOptions): Promise<WorkbookStream> {
  const input = checkSource(source);
  const o = readOptions(options, ['limits'], 'openWorkbookStream options');
  const userLimits = o.get('limits');
  const limits = resolveLimits({ ...STREAM_LIMIT_DEFAULTS, ...(userLimits as Partial<ReadLimits> | undefined) });
  const zip = await openZipSource(input, limits);
  const preloaded = new PreloadedZip(zip);
  await preloaded.load(zip.entries.filter((e) => e.name.toLowerCase() === '[content_types].xml' || e.name.toLowerCase().endsWith('.rels')).map((e) => e.name));
  const pkg = openPackage(preloaded, limits);
  await preloaded.load(workbookPartsToLoad(pkg).map(zipName));
  const warnings = new WarningCollector();
  const info: WorkbookInfo = analyzeWorkbook(pkg, limits, warnings);
  const formatCache = new Map<number, NumberFormat>();

  const sheetEntry = (name: string) => {
    const entry = info.sheets.find((s) => s.name === name);
    if (entry !== undefined) return entry;
    const names = info.sheets.map((s) => s.name);
    const close = suggest(String(name), names);
    throw new XlvError('XLV_E_SHEET_NOT_FOUND', `The workbook has no sheet named "${String(name).slice(0, 100)}".`, {
      hint: `${close.length > 0 ? `Did you mean "${close[0] ?? ''}"? ` : ''}Available sheets: ${names.join(', ') || '(none)'}. Sheet names are case-sensitive.`,
      details: { name: String(name).slice(0, 100), available: names },
    });
  };

  async function* rows(name: string): AsyncGenerator<StreamRow> {
    const entry = sheetEntry(name);
    if (entry.kind !== 'worksheet' || entry.part === null) return;
    const part = entry.part;
    const ready: StreamRow[] = [];
    const { handler, finish } = createWorksheetParser(
      {
        part,
        sheet: entry.name,
        limits,
        warnings,
        sharedStrings: info.sharedStrings,
        styles: info.styles,
        dateSystem: info.dateSystem,
        conformance: info.conformance,
        budget: { remaining: Infinity },
        formatCache,
      },
      (number, cells, rowInfo) => ready.push({ number, cells, height: rowInfo?.height ?? null, hidden: rowInfo?.hidden ?? false }),
    );
    const parser = createXmlByteStreamParser(handler, { part, limits });
    const reader = zip.stream(zipName(part)).getReader();
    try {
      for (;;) {
        const { done, value } = await reader.read();
        if (done) break;
        parser.write(value);
        while (ready.length > 0) yield ready.shift() as StreamRow;
      }
      parser.end();
      finish();
      while (ready.length > 0) yield ready.shift() as StreamRow;
    } finally {
      // Stops decompression when the caller breaks out of the loop early.
      await reader.cancel().catch(() => undefined);
    }
  }

  return {
    sheets: info.sheets.map((s) => ({ name: s.name, index: s.index, kind: s.kind, state: s.state })),
    dateSystem: info.dateSystem,
    conformance: info.conformance,
    format: info.format,
    hasMacros: info.hasMacros,
    externalLinkCount: info.externalLinkCount,
    get warnings() {
      return warnings.list();
    },
    rows(name: string): AsyncIterableIterator<StreamRow> {
      sheetEntry(name);
      return rows(name);
    },
    rowStream(name: string): ReadableStream<StreamRow> {
      sheetEntry(name);
      const iterator = rows(name);
      return new ReadableStream<StreamRow>({
        async pull(controller) {
          const next = await iterator.next();
          if (next.done) controller.close();
          else controller.enqueue(next.value);
        },
        async cancel() {
          await iterator.return(undefined);
        },
      });
    },
  };
}

/**
 * Options for {@link readTableStream}.
 *
 * @example
 * ```ts
 * import type { ReadTableStreamOptions } from 'xlvane';
 *
 * const options: ReadTableStreamOptions = { sheet: 'Orders', header: 1, dates: 'iso' };
 * ```
 */
export interface ReadTableStreamOptions<S extends StandardSchemaV1 | undefined = undefined> extends OpenWorkbookStreamOptions {
  readonly sheet?: string;
  /** `"auto"` (default): detected in the first 20 non-empty rows; a row number; or `false`. */
  readonly header?: 'auto' | number | false;
  readonly schema?: S;
  readonly dates?: 'date' | 'iso';
}

/**
 * One row from {@link readTableStream}: the validated value, or the issues found.
 *
 * @example
 * ```ts
 * import { readTableStream, type TableStreamItem, type TableRow } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * for await (const item of readTableStream(bytes)) {
 *   const typed: TableStreamItem<TableRow> = item;
 *   if (typed.issues.length > 0) console.log(typed.row, typed.issues[0]?.message);
 * }
 * ```
 */
export interface TableStreamItem<Row> {
  /** 1-based sheet row. */
  readonly row: number;
  /** The row object (the schema's output when valid); `undefined` when the schema rejected it. */
  readonly value: Row | undefined;
  readonly issues: readonly TableIssue[];
}

function setOwn(target: Record<string, unknown>, key: string, value: unknown): void {
  Object.defineProperty(target, key, { value, enumerable: true, writable: true, configurable: true });
}

/**
 * Reads a sheet as row objects, one at a time, with optional schema validation: the streaming
 * counterpart of `readTable` for large files. The header row is detected in the first 20
 * non-empty rows (or given), and the columns are those of the header row (without a header:
 * the columns used in those first rows); cells outside them are ignored.
 *
 * @example
 * ```ts
 * import { readTableStream } from 'xlvane';
 * import { z } from 'zod';
 *
 * declare const file: Blob;
 * const schema = z.object({ Customer: z.string(), Total: z.number() });
 * for await (const { row, value, issues } of readTableStream(file, { sheet: 'Orders', schema })) {
 *   if (value) console.log(value.Customer, value.Total);
 *   else console.log(`row ${String(row)}:`, issues.map((i) => i.message).join('; '));
 * }
 * ```
 */
export function readTableStream<S extends StandardSchemaV1>(
  source: Uint8Array | ArrayBuffer | Blob,
  options: ReadTableStreamOptions<S> & { readonly schema: S },
): AsyncGenerator<TableStreamItem<StandardSchemaV1.InferOutput<S>>>;
export function readTableStream(source: Uint8Array | ArrayBuffer | Blob, options?: ReadTableStreamOptions): AsyncGenerator<TableStreamItem<TableRow>>;
export async function* readTableStream(
  source: Uint8Array | ArrayBuffer | Blob,
  options?: ReadTableStreamOptions<StandardSchemaV1 | undefined>,
): AsyncGenerator<TableStreamItem<unknown>> {
  const o = readOptions(options, ['sheet', 'header', 'schema', 'dates', 'limits'], 'readTableStream options');
  const dates = oneOfOption('dates', o.get('dates'), ['date', 'iso'], 'date');
  const header = o.get('header') ?? 'auto';
  if (header !== 'auto' && header !== false && !(typeof header === 'number' && Number.isInteger(header) && header >= 1)) {
    throw new XlvError('XLV_E_INVALID_OPTION', 'The option "header" must be "auto", a row number or false.', { hint: 'Use header: "auto", header: 1 or header: false.', details: { option: 'header' } });
  }
  const schema = o.get('schema') as StandardSchemaV1 | undefined;
  if (schema !== undefined) {
    const props = (schema as { '~standard'?: { version?: unknown; validate?: unknown } } | null)?.['~standard'];
    if (props?.version !== 1 || typeof props.validate !== 'function') {
      throw new XlvError('XLV_E_SCHEMA_INVALID', 'The option "schema" is not a Standard Schema (version 1).', { hint: 'Pass a Zod, Valibot or ArkType schema.', details: { option: 'schema' } });
    }
  }
  const limits = o.get('limits');
  const book = await openWorkbookStream(source, limits === undefined ? undefined : { limits: limits as Partial<ReadLimits> });
  const sheetName = (o.get('sheet') as string | undefined) ?? (book.sheets.find((s) => s.kind === 'worksheet' && s.state === 'visible') ?? book.sheets[0])?.name;
  if (sheetName === undefined) return;

  const buffered: StreamRow[] = [];
  let columns: TableColumn[] | null = null;
  let headerRow: number | null = null;
  const value = (cell: Cell | undefined): TableValue => {
    const v = cell?.value ?? null;
    return v instanceof Date && dates === 'iso' ? v.toISOString() : v;
  };
  const decide = (): TableColumn[] => {
    const nonEmpty = buffered.filter((r) => r.cells.length > 0);
    if (header === 'auto') headerRow = detectHeader(nonEmpty.map((r) => r.cells), widthOf(nonEmpty), 20)?.row ?? null;
    else if (typeof header === 'number') {
      if (!nonEmpty.some((r) => r.number === header)) {
        throw new XlvError('XLV_E_HEADER_NOT_FOUND', `Row ${String(header)} of sheet "${sheetName}" has no cells to use as a header.`, {
          location: { sheet: sheetName },
          hint: 'Pass the row that holds the column names, header: "auto", or header: false.',
          details: { row: header },
        });
      }
      headerRow = header;
    }
    const source = headerRow === null ? nonEmpty.flatMap((r) => r.cells) : (nonEmpty.find((r) => r.number === headerRow)?.cells ?? []);
    const minCol = Math.min(...source.map((c) => c.col));
    const maxCol = Math.max(...source.map((c) => c.col));
    const out: TableColumn[] = [];
    const used = new Set<string>();
    for (let col = minCol; col <= maxCol; col++) {
      const letter = columnLetters(col);
      const cell = headerRow === null ? undefined : source.find((c) => c.col === col);
      const text = cell === undefined || cell.value === null ? null : cell.value instanceof Date ? cell.value.toISOString() : String(cell.value);
      const base = text === null || text.trim() === '' ? letter : text.trim();
      let key = base;
      for (let n = 2; used.has(key); n++) key = `${base}_${String(n)}`;
      used.add(key);
      out.push({ key, header: text === null || text.trim() === '' ? null : text, col, letter });
    }
    return out;
  };

  const emit = async (r: StreamRow, cols: TableColumn[]): Promise<TableStreamItem<unknown> | null> => {
    if (headerRow !== null && r.number <= headerRow) return null;
    const byCol = new Map(r.cells.map((c) => [c.col, c]));
    if (cols.every((c) => byCol.get(c.col)?.value == null)) return null;
    const row: Record<string, TableValue> = {};
    for (const c of cols) setOwn(row, c.key, value(byCol.get(c.col)));
    if (schema === undefined) return { row: r.number, value: row, issues: [] };
    const result = await schema['~standard'].validate(row);
    if (!result.issues) return { row: r.number, value: result.value, issues: [] };
    const issues = result.issues.map((issue): TableIssue => {
      const path = (issue.path ?? []).map((seg) => {
        const key = typeof seg === 'object' && seg !== null ? seg.key : seg;
        return typeof key === 'number' ? key : String(key);
      });
      const column = typeof path[0] === 'string' ? cols.find((c) => c.key === path[0]) : undefined;
      return { row: r.number, cell: column ? formatCellRef(r.number, column.col) : null, column: column?.key ?? null, message: String(issue.message), path };
    });
    return { row: r.number, value: undefined, issues };
  };

  for await (const r of book.rows(sheetName)) {
    if (columns === null) {
      buffered.push(r);
      if (buffered.filter((b) => b.cells.length > 0).length < 20) continue;
      columns = decide();
      for (const b of buffered.splice(0)) {
        const item = await emit(b, columns);
        if (item) yield item;
      }
      continue;
    }
    const item = await emit(r, columns);
    if (item) yield item;
  }
  if (columns === null && buffered.some((b) => b.cells.length > 0)) {
    columns = decide();
    for (const b of buffered.splice(0)) {
      const item = await emit(b, columns);
      if (item) yield item;
    }
  }
}

function widthOf(rows: readonly StreamRow[]): number {
  let min = Infinity;
  let max = 0;
  for (const r of rows) for (const c of r.cells) {
    if (c.col < min) min = c.col;
    if (c.col > max) max = c.col;
  }
  return max === 0 ? 0 : max - min + 1;
}
