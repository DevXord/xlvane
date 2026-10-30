// exceljs streaming API: `stream.xlsx.WorkbookWriter` (lib/stream/xlsx/workbook-writer.js) and
// `stream.xlsx.WorkbookReader` (lib/stream/xlsx/workbook-reader.js), built on xlvane's
// writeWorkbookStream and openWorkbookStream (docs/decisions.md D-507).
import { openWorkbookStream, writeWorkbookStream, XlvError, type CellInput, type CellStyle, type HyperlinkTarget, type NoteInput, type StreamRowInput, type StreamSheetSpec, type WorkbookProperties, type WriteWorkbookStreamOptions } from 'xlvane';
import { autoFilterRange, cellStyleIn, columnIn, isDateCell, DATE_NUMFMT, layoutIn, linkIn, noteIn, richTextOut, sheetRulesIn, styleIn, styleOut, valueIn } from './convert';
import type { DefinedNames } from './defined-names';
import { imageAnchorIn, imageBytes, type ImageData } from './media';
import { columnLetter, notSupported, Row, ValueType, Worksheet } from './model';
import type { CellValue, WorksheetOptions } from './types';
import { Workbook } from './workbook';

/** An unbounded async queue: producers push synchronously, the stream pulls. */
class Channel<T> implements AsyncIterable<T> {
  private readonly items: T[] = [];
  private waiting: { resolve: (r: IteratorResult<T>) => void; reject: (e: unknown) => void } | null = null;
  private closed = false;
  private error: unknown = undefined;

  push(item: T): void {
    if (this.closed) throw new Error('internal: push after close');
    if (this.waiting) {
      const w = this.waiting;
      this.waiting = null;
      w.resolve({ value: item, done: false });
    } else this.items.push(item);
  }

  close(): void {
    this.closed = true;
    if (this.waiting) {
      const w = this.waiting;
      this.waiting = null;
      w.resolve({ value: undefined as T, done: true });
    }
  }

  /** Closes the queue so that the consumer fails with `error`. */
  fail(error: unknown): void {
    this.closed = true;
    this.error = error;
    this.items.length = 0;
    if (this.waiting) {
      const w = this.waiting;
      this.waiting = null;
      w.reject(error);
    }
  }

  get isClosed(): boolean {
    return this.closed;
  }

  [Symbol.asyncIterator](): AsyncIterator<T> {
    return {
      next: (): Promise<IteratorResult<T>> => {
        const item = this.items.shift();
        if (item !== undefined) return Promise.resolve({ value: item, done: false });
        if (this.error !== undefined) return Promise.reject(this.error);
        if (this.closed) return Promise.resolve({ value: undefined as T, done: true });
        return new Promise((resolve, reject) => {
          this.waiting = { resolve, reject };
        });
      },
    };
  }
}

/** A writable Node.js stream. */
interface WritableLike {
  write(chunk: Uint8Array, callback?: (error?: unknown) => void): unknown;
  end(callback?: () => void): unknown;
}

/**
 * Options of {@link WorkbookWriter}, as exceljs `WorkbookWriter` options.
 *
 * @example
 * ```ts
 * import type { WorkbookWriterOptions } from '@xlvane/exceljs';
 *
 * const options: WorkbookWriterOptions = { filename: 'big.xlsx', useSharedStrings: false, useStyles: true };
 * ```
 */
export interface WorkbookWriterOptions {
  /** Output file (Node.js). */
  filename?: string;
  /** Output stream. */
  stream?: WritableLike;
  /** Use a shared string table (default false: strings are stored in their cells). */
  useSharedStrings?: boolean;
  /** Write cell styles (default false, as in exceljs). */
  useStyles?: boolean;
  /** Document properties written to `docProps/core.xml` (defaults as in exceljs). */
  created?: Date;
  modified?: Date;
  creator?: string;
  lastModifiedBy?: string;
  lastPrinted?: Date;
  zip?: unknown;
}

/**
 * A worksheet of a {@link WorkbookWriter}: rows are written when committed and then dropped.
 *
 * @example
 * ```ts
 * import ExcelJS from '@xlvane/exceljs';
 *
 * const writer = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: 'big.xlsx' });
 * const sheet = writer.addWorksheet('Data');
 * sheet.addRow([1, 'one']).commit();
 * sheet.commit();
 * await writer.commit();
 * ```
 */
export class WorksheetWriter extends Worksheet {
  private readonly rowsOut = new Channel<StreamRowInput>();
  private started = false;
  private committedUpTo = 0;
  /** The layout written when the sheet started, to detect later changes that would be lost. */
  private layout = '';

  /** @internal */
  constructor(
    private readonly writer: WorkbookWriter,
    id: number,
    name: string | undefined,
    options: WorksheetOptions | undefined,
  ) {
    super(writer._book, id, name, options);
  }

  /** @internal The sheet spec, made when the first row is committed (layout must be set by then). */
  _spec(): StreamSheetSpec {
    this.started = true;
    const layout = this.currentLayout();
    this.layout = JSON.stringify(layout);
    return { ...layout, rows: this.rowsOut };
  }

  private currentLayout(): Omit<StreamSheetSpec, 'rows'> {
    const layout = layoutIn(this);
    const filter = autoFilterRange(this.autoFilter);
    return {
      name: this.name,
      state: this.state,
      columns: this.columns.map((c) => columnIn(c, this.writer._useStyles)),
      ...(layout.freeze ? { freeze: layout.freeze } : {}),
      ...(Object.keys(layout.view).length > 0 ? { view: layout.view } : {}),
      ...(layout.tabColor !== undefined ? { tabColor: layout.tabColor } : {}),
      ...(filter ? { autoFilter: filter } : {}),
      merges: [...this._merges.values()].map((m) => `${columnLetter(m.left)}${String(m.top)}:${columnLetter(m.right)}${String(m.bottom)}`),
      ...sheetRulesSpec(this),
      ...this.mediaSpec(),
    };
  }

  /** The background (exceljs `addBackgroundImage`, worksheet-writer.js L474) and pictures, by image id. */
  private mediaSpec(): Pick<StreamSheetSpec, 'images' | 'background'> {
    const background = this.getBackgroundImageId();
    const images = this.getImages().flatMap((m) => {
      if (!m.range) return [];
      const { anchor, hyperlink } = imageAnchorIn(this, m.range);
      return [{ image: m.imageId, anchor, ...(hyperlink === undefined ? {} : { hyperlink }) }];
    });
    return { ...(background === undefined ? {} : { background }), ...(images.length > 0 ? { images } : {}) };
  }

  override addTable(): never {
    // exceljs's streaming WorksheetWriter has no addTable (X4-21).
    return notSupported('worksheet.addTable() on a streamed worksheet');
  }

  override mergeCells(...args: unknown[]): void {
    if (this.started) notSupported('mergeCells() after rows of a streamed worksheet were committed');
    super.mergeCells(...args);
  }

  // exceljs's streaming WorksheetWriter has none of these methods (lib/stream/xlsx/worksheet-writer.js,
  // X4-15); committed rows are already written, so rows and columns cannot move.
  override spliceRows(): void {
    notSupported('worksheet.spliceRows() on a streamed worksheet');
  }
  override insertRow(): Row {
    return notSupported('worksheet.insertRow() on a streamed worksheet');
  }
  override insertRows(): Row[] | undefined {
    return notSupported('worksheet.insertRows() on a streamed worksheet');
  }
  override duplicateRow(): void {
    notSupported('worksheet.duplicateRow() on a streamed worksheet');
  }
  override spliceColumns(): void {
    notSupported('worksheet.spliceColumns() on a streamed worksheet');
  }

  /** @internal Writes every row up to `row` and drops it from memory. */
  override _commitRow(row: Row): void {
    this.writer._start(this);
    for (let n = this.committedUpTo + 1; n <= row.number; n++) this.flush(n);
    this.committedUpTo = Math.max(this.committedUpTo, row.number);
    this._committedRows = this.committedUpTo;
  }

  private flush(n: number): void {
    const row = this._rows[n - 1];
    this._rows[n - 1] = undefined;
    if (!row) {
      this.rowsOut.push({ values: [] });
      return;
    }
    const values: CellValue[] = [];
    const styles: (CellStyle | undefined)[] = [];
    const links: (HyperlinkTarget | undefined)[] = [];
    const notes: (NoteInput | undefined)[] = [];
    let styled = false;
    row._cells.forEach((cell, i) => {
      if (!cell) return;
      // Links and notes go with their row (StreamRowInput hyperlinks and notes, D-A12).
      const link = cell.type === ValueType.Merge ? null : linkIn(cell);
      if (link) links[i] = link;
      const note = noteIn(cell);
      if (note) notes[i] = note;
      values[i] = cell.type === ValueType.Merge ? null : (valueIn(cell) as CellValue);
      // Without useStyles exceljs still gives dates format 14 (styles-xform.js Mock, L498; X4-02).
      if (this.writer._useStyles) styles[i] = cellStyleIn(cell, styleIn(cell.style));
      else if (isDateCell(cell)) styles[i] = { numFmt: DATE_NUMFMT };
      if (styles[i] !== undefined) styled = true;
    });
    for (let i = 0; i < values.length; i++) if (values[i] === undefined) values[i] = null;
    this.rowsOut.push({
      values: values as CellInput[],
      ...(styled ? { styles } : {}),
      ...(links.length > 0 ? { hyperlinks: links } : {}),
      ...(notes.length > 0 ? { notes } : {}),
      ...(row.height !== undefined ? { height: row.height } : {}),
      ...(row.hidden ? { hidden: true } : {}),
    });
  }

  /** Commits the remaining rows and closes the worksheet. */
  commit(): void {
    if (this.rowsOut.isClosed) return;
    const last = this._rows.length;
    if (last > this.committedUpTo) {
      const row = this.getRow(last);
      this._commitRow(row);
    } else this.writer._start(this);
    // The sheet layout (columns, views, autoFilter, merges, name, state) is written before the
    // rows, so a change made after the first committed row cannot be written (D-507).
    if (JSON.stringify(this.currentLayout()) !== this.layout) {
      const error = notSupportedLater(`Changing the layout of streamed worksheet "${this.name}" (columns, views, autoFilter, merges, validations, conditional formats, protection, page setup) after rows were committed`);
      this.rowsOut.fail(error);
      throw error;
    }
    this.rowsOut.close();
  }
}

/**
 * The streaming writer of exceljs (`ExcelJS.stream.xlsx.WorkbookWriter`). Committed rows are
 * written at once and dropped, so memory stays constant. Strings are stored in their cells
 * unless `useSharedStrings` is set; styles are written only with `useStyles: true`, as in exceljs.
 *
 * @example
 * ```ts
 * import ExcelJS from '@xlvane/exceljs';
 *
 * const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: 'big.xlsx', useStyles: true });
 * const sheet = workbook.addWorksheet('Data');
 * sheet.columns = [{ header: 'Id', key: 'id', width: 10 }];
 * for (let id = 1; id <= 100_000; id++) sheet.addRow({ id }).commit();
 * sheet.commit();
 * await workbook.commit();
 * ```
 */
export class WorkbookWriter {
  /** @internal */
  readonly _book: Workbook = new Workbook();
  /** @internal */
  readonly _useStyles: boolean;
  private readonly sheets = new Channel<StreamSheetSpec>();
  private readonly open: WorksheetWriter[] = [];
  private done: Promise<void> | null = null;
  private readonly options: WorkbookWriterOptions;
  private written = '';
  private nextId = 1;
  created: Date;
  modified: Date;
  creator: string;
  lastModifiedBy: string;
  lastPrinted: Date | undefined;

  constructor(options: WorkbookWriterOptions = {}) {
    this.options = options;
    this._useStyles = options.useStyles === true;
    // Defaults of exceljs WorkbookWriter (lib/stream/xlsx/workbook-writer.js L24; X4-03).
    this.created = options.created || new Date();
    this.modified = options.modified || this.created;
    this.creator = options.creator || 'ExcelJS';
    this.lastModifiedBy = options.lastModifiedBy || 'ExcelJS';
    this.lastPrinted = options.lastPrinted;
  }

  /**
   * Adds an image (PNG, JPEG or GIF; `buffer` or `base64`) and returns its id, as exceljs
   * `workbookWriter.addImage`. Images are written when the first sheet starts, so they must be
   * added before the first row is committed.
   */
  addImage(image: ImageData): number {
    if (this.done) throw notSupportedLater('workbookWriter.addImage() after rows were committed');
    return this._book.addImage(image);
  }

  /** The image with this id. */
  getImage(id: number): ImageData | undefined {
    return this._book.getImage(id);
  }

  /** Defined names, as exceljs `workbookWriter.definedNames`; set them before the first row is committed. */
  get definedNames(): DefinedNames {
    return this._book.definedNames;
  }

  /** The document properties and names written; they are fixed when writing starts. */
  private meta(): { properties: WorkbookProperties; definedNames: { name: string; refersTo: string }[] } {
    const properties: Record<string, unknown> = { creator: this.creator, lastModifiedBy: this.lastModifiedBy, created: this.created, modified: this.modified };
    if (this.lastPrinted) properties['lastPrinted'] = this.lastPrinted;
    return { properties: properties as WorkbookProperties, definedNames: this._book.definedNames.model.map((d) => ({ name: d.name, refersTo: d.ranges.join(',') })) };
  }

  /** Starts the xlvane stream (once): when the first sheet starts or at commit. */
  private begin(): Promise<void> {
    if (this.done) return this.done;
    const meta = this.meta();
    this.written = JSON.stringify(meta);
    const images = this._book._media.map((m) => imageBytes(m));
    const spec: WriteWorkbookStreamOptions = { sheets: this.sheets, sharedStrings: this.options.useSharedStrings === true, ...meta, ...(images.length > 0 ? { images } : {}) };
    const options = this.options;
    if (options.filename !== undefined) {
      const filename = options.filename;
      // exceljs replaces an existing file, so overwrite is on.
      this.done = import('xlvane/node').then(({ writeWorkbookStreamFile }) => writeWorkbookStreamFile(filename, spec, { overwrite: true }));
    } else if (options.stream !== undefined) {
      const out = options.stream;
      this.done = (async () => {
        const stream = writeWorkbookStream(spec);
        const reader = stream.getReader();
        for (;;) {
          const { done, value } = await reader.read();
          if (done) break;
          await new Promise<void>((resolve, reject) => {
            out.write(value, (error) => (error ? reject(error) : resolve()));
          });
        }
        await new Promise<void>((resolve) => {
          out.end(() => resolve());
        });
      })();
    } else {
      this.done = Promise.reject(notSupportedLater('A WorkbookWriter without filename or stream'));
    }
    // Errors are reported by commit(); avoid an unhandled rejection before that.
    this.done.catch(() => undefined);
    return this.done;
  }

  /** Adds a worksheet. Sheets are written in the order their first rows are committed. */
  addWorksheet(name?: string, options?: WorksheetOptions): WorksheetWriter {
    const sheet = new WorksheetWriter(this, this.nextId++, name, options);
    this._book._sheets[sheet.id] = sheet;
    this.open.push(sheet);
    return sheet;
  }

  getWorksheet(id?: number | string): WorksheetWriter | undefined {
    if (id === undefined) return this.open[0];
    return this.open.find((s) => (typeof id === 'number' ? s.id === id : s.name === id));
  }

  /** @internal Starts writing a sheet (once). */
  _start(sheet: WorksheetWriter): void {
    if ((sheet as unknown as { _queued?: boolean })._queued) return;
    void this.begin();
    (sheet as unknown as { _queued?: boolean })._queued = true;
    this.sheets.push(sheet._spec());
  }

  /** Commits all worksheets and finishes the file. */
  async commit(): Promise<void> {
    for (const sheet of this.open) sheet.commit();
    const done = this.begin();
    // Properties and names are written with the workbook part's data fixed at the start (D-A06).
    if (JSON.stringify(this.meta()) !== this.written) {
      const error = notSupportedLater('Changing document properties or defined names of a WorkbookWriter after rows were committed');
      this.sheets.fail(error);
      await done.catch(() => undefined);
      throw error;
    }
    this.sheets.close();
    await done;
  }
}

/** The validations, conditional formats, protection and printing of a streamed worksheet (D-C06). */
function sheetRulesSpec(ws: Worksheet): Partial<StreamSheetSpec> {
  const r = sheetRulesIn(ws);
  return {
    ...(r.dataValidations.length > 0 ? { dataValidations: r.dataValidations.map((v) => ({ ...v.rule, sqref: v.sqref })) } : {}),
    ...(r.conditionalFormats.length > 0 ? { conditionalFormats: r.conditionalFormats } : {}),
    ...(r.protect ? { protect: r.protect } : {}),
    pageSetup: r.pageSetup,
    ...(r.headerFooter ? { headerFooter: r.headerFooter } : {}),
    ...(r.pageBreaks.length > 0 ? { pageBreaks: { rows: r.pageBreaks } } : {}),
    ...(r.printArea ? { printArea: r.printArea } : {}),
    ...(r.printTitles ? { printTitles: r.printTitles } : {}),
  };
}

function notSupportedLater(feature: string): Error {
  try {
    notSupported(feature);
  } catch (error) {
    return error as Error;
  }
  return new Error(feature);
}

type Listener = (...args: unknown[]) => void;

/** The part of Node.js `EventEmitter` that exceljs's streaming reader uses; no node:events import. */
class Emitter {
  private readonly listeners = new Map<string, { fn: Listener; once: boolean }[]>();

  on(event: string, fn: Listener): this {
    const list = this.listeners.get(event) ?? [];
    list.push({ fn, once: false });
    this.listeners.set(event, list);
    return this;
  }

  addListener(event: string, fn: Listener): this {
    return this.on(event, fn);
  }

  once(event: string, fn: Listener): this {
    this.on(event, fn);
    this.listeners.get(event)!.at(-1)!.once = true;
    return this;
  }

  off(event: string, fn: Listener): this {
    const list = this.listeners.get(event);
    if (list) {
      const i = list.findIndex((l) => l.fn === fn);
      if (i >= 0) list.splice(i, 1);
    }
    return this;
  }

  removeListener(event: string, fn: Listener): this {
    return this.off(event, fn);
  }

  removeAllListeners(event?: string): this {
    if (event === undefined) this.listeners.clear();
    else this.listeners.delete(event);
    return this;
  }

  listenerCount(event: string): number {
    return this.listeners.get(event)?.length ?? 0;
  }

  emit(event: string, ...args: unknown[]): boolean {
    const list = this.listeners.get(event);
    if (!list || list.length === 0) {
      // Node.js throws an unhandled 'error' event; so does this emitter.
      if (event === 'error') throw args[0] instanceof Error ? args[0] : new Error(String(args[0]));
      return false;
    }
    for (const l of [...list]) {
      if (l.once) this.off(event, l.fn);
      l.fn(...args);
    }
    return true;
  }
}

/** exceljs `WorkbookReader.Options` (workbook-reader.js L326–332). */
const READER_OPTIONS: Record<string, readonly string[]> = {
  worksheets: ['emit', 'ignore'],
  sharedStrings: ['cache', 'emit', 'ignore'],
  hyperlinks: ['cache', 'emit', 'ignore'],
  styles: ['cache', 'ignore'],
  entries: ['emit', 'ignore'],
};

/**
 * Options of {@link WorkbookReader}, as exceljs `WorkbookStreamReaderOptions` (index.d.ts L1988).
 *
 * @example
 * ```ts
 * import type { WorkbookReaderOptions } from '@xlvane/exceljs';
 *
 * const options: WorkbookReaderOptions = { styles: 'cache', entries: 'emit' };
 * ```
 */
export interface WorkbookReaderOptions {
  worksheets?: 'emit' | 'ignore';
  /** Only `cache` (the default): `emit` and `ignore` throw `XLV_E_NOT_SUPPORTED`. */
  sharedStrings?: 'cache' | 'emit' | 'ignore';
  /** `ignore` (default) or `cache`, which changes nothing for files in schema order (links follow the rows); `emit` throws `XLV_E_NOT_SUPPORTED`. */
  hyperlinks?: 'cache' | 'emit' | 'ignore';
  /** `cache` gives the cells their styles; `ignore` (default) does not. */
  styles?: 'cache' | 'ignore';
  /** `emit` sends `entry` events for the workbook and each worksheet. */
  entries?: 'emit' | 'ignore';
}

function readerOptions(raw: Record<string, unknown>): Required<WorkbookReaderOptions> {
  const out: Record<string, string> = { worksheets: 'emit', sharedStrings: 'cache', hyperlinks: 'ignore', styles: 'ignore', entries: 'ignore' };
  for (const [key, allowed] of Object.entries(READER_OPTIONS)) {
    const v = raw[key];
    if (v === undefined) continue;
    if (typeof v !== 'string' || !allowed.includes(v)) {
      throw new XlvError('XLV_E_INVALID_OPTION', `The WorkbookReader option "${key}" must be one of ${allowed.map((a) => `"${a}"`).join(', ')}.`, {
        hint: `For example { ${key}: "${allowed[0]!}" }.`,
        details: { option: key },
      });
    }
    out[key] = v;
  }
  if (out['sharedStrings'] !== 'cache') notSupported(`WorkbookReader option sharedStrings: "${out['sharedStrings']!}" (only "cache")`);
  if (out['hyperlinks'] === 'emit') notSupported('WorkbookReader option hyperlinks: "emit"');
  return out as Required<WorkbookReaderOptions>;
}

/** The range of the rows read so far, as exceljs's `Dimensions` (lib/doc/range.js). */
class ReaderDimensions {
  top = 0;
  left = 0;
  bottom = 0;
  right = 0;

  /** @internal */
  expandRow(row: Row): void {
    const d = row.dimensions;
    if (!d) return;
    if (!this.top || row.number < this.top) this.top = row.number;
    if (!this.left || d.min < this.left) this.left = d.min;
    if (!this.bottom || row.number > this.bottom) this.bottom = row.number;
    if (!this.right || d.max > this.right) this.right = d.max;
  }

  get model(): { top: number; left: number; bottom: number; right: number } {
    return { top: this.top, left: this.left, bottom: this.bottom, right: this.right };
  }

  get tl(): string {
    return `${columnLetter(this.left || 1)}${String(this.top || 1)}`;
  }

  get br(): string {
    return `${columnLetter(this.right || 1)}${String(this.bottom || 1)}`;
  }

  get range(): string {
    return `${this.tl}:${this.br}`;
  }
}

/**
 * A worksheet read by {@link WorkbookReader}, as exceljs `WorksheetReader`: iterate its rows with
 * `for await`, or call `read()` and listen to `row` and `finished` events.
 *
 * @example
 * ```ts
 * import ExcelJS from '@xlvane/exceljs';
 *
 * const reader = new ExcelJS.stream.xlsx.WorkbookReader('big.xlsx', {});
 * for await (const worksheet of reader) {
 *   console.log(worksheet.name, worksheet.state);
 *   for await (const row of worksheet) console.log(row.number);
 *   console.log(worksheet.dimensions.range);
 * }
 * ```
 */
export class WorksheetReader extends Emitter implements AsyncIterable<Row> {
  readonly dimensions: ReaderDimensions = new ReaderDimensions();

  /** @internal */
  constructor(
    readonly workbook: WorkbookReader,
    readonly id: number,
    readonly name: string,
    readonly state: 'visible' | 'hidden' | 'veryHidden',
    private readonly sheet: Worksheet,
    private readonly rows: () => AsyncIterable<Row>,
  ) {
    super();
  }

  /** Column definitions read so far, or `null` (as exceljs before any `getColumn`). */
  get columns(): ReturnType<Worksheet['getColumn']>[] | null {
    return this.sheet.columns.length > 0 ? this.sheet.columns : null;
  }

  getColumn(c: number | string): ReturnType<Worksheet['getColumn']> {
    return this.sheet.getColumn(c);
  }

  getColumnKey(key: string): ReturnType<Worksheet['getColumnKey']> {
    return this.sheet.getColumnKey(key);
  }

  setColumnKey(key: string, value: ReturnType<Worksheet['getColumn']>): void {
    this.sheet.setColumnKey(key, value);
  }

  deleteColumnKey(key: string): void {
    this.sheet.deleteColumnKey(key);
  }

  eachColumnKey(f: Parameters<Worksheet['eachColumnKey']>[0]): void {
    this.sheet.eachColumnKey(f);
  }

  /** Not valid on a streamed worksheet (it throws in exceljs too). */
  destroy(): never {
    throw new Error('Invalid Operation: destroy');
  }

  async *[Symbol.asyncIterator](): AsyncIterator<Row> {
    for await (const row of this.rows()) {
      this.dimensions.expandRow(row);
      yield row;
    }
  }

  /** Events of the rows: `row` (a Row), then `finished`; `error` on failure. */
  async *parse(): AsyncGenerator<{ eventType: string; value: unknown }[]> {
    for await (const row of this) yield [{ eventType: 'row', value: row }];
  }

  /** Reads the rows and emits `row` events, then `finished` (or `error`). */
  async read(): Promise<void> {
    try {
      for await (const row of this) this.emit('row', row);
      this.emit('finished');
    } catch (error) {
      this.emit('error', error);
    }
  }
}

/**
 * The streaming reader of exceljs (`ExcelJS.stream.xlsx.WorkbookReader`): iterate worksheets
 * and their rows with `for await`, or call `read()` and listen to `worksheet`, `end`,
 * `finished` and `error` events (and `entry` with `entries: 'emit'`). Rows are not kept.
 *
 * @example
 * ```ts
 * import ExcelJS from '@xlvane/exceljs';
 *
 * const reader = new ExcelJS.stream.xlsx.WorkbookReader('big.xlsx', { styles: 'cache' });
 * for await (const worksheet of reader) {
 *   for await (const row of worksheet) console.log(row.number, row.values);
 * }
 * ```
 */
export class WorkbookReader extends Emitter implements AsyncIterable<WorksheetReader> {
  private options: Required<WorkbookReaderOptions>;

  constructor(
    private readonly input: string | Uint8Array | ArrayBuffer | AsyncIterable<Uint8Array>,
    options: WorkbookReaderOptions | Record<string, unknown> = {},
  ) {
    super();
    this.options = readerOptions(options as Record<string, unknown>);
  }

  private async open(input: string | Uint8Array | ArrayBuffer | AsyncIterable<Uint8Array>): Promise<Awaited<ReturnType<typeof openWorkbookStream>>> {
    if (typeof input === 'string') {
      const { openWorkbookStreamFile } = await import('xlvane/node');
      return openWorkbookStreamFile(input);
    }
    if (input instanceof Uint8Array || input instanceof ArrayBuffer) return openWorkbookStream(input);
    const chunks: Uint8Array[] = [];
    for await (const c of input) chunks.push(c);
    return openWorkbookStream(new Blob(chunks as BlobPart[]));
  }

  /** Worksheet events in workbook order, as exceljs `parse()` yields them (`{ eventType: 'worksheet', value }`). */
  async *parse(input?: string | Uint8Array | ArrayBuffer | AsyncIterable<Uint8Array>, options?: WorkbookReaderOptions): AsyncGenerator<{ eventType: string; value: unknown }> {
    if (options) this.options = readerOptions(options as Record<string, unknown>);
    const book = await this.open(input ?? this.input);
    if (this.options.entries === 'emit') this.emit('entry', { type: 'workbook' });
    if (this.options.worksheets === 'ignore') return;
    const host = new Workbook();
    const styles = this.options.styles === 'cache';
    let id = 0;
    for (const info of book.sheets) {
      if (info.kind !== 'worksheet') continue;
      id++;
      if (this.options.entries === 'emit') this.emit('entry', { type: 'worksheet', id: String(id) });
      const sheet = host.addWorksheet(info.name, { state: info.state });
      const name = info.name;
      const rows = async function* (): AsyncGenerator<Row> {
        for await (const r of book.rows(name)) {
          const row = sheet.getRow(r.number);
          for (const cell of r.cells) {
            const plain = cell.type === 'error' ? { error: cell.value as '#N/A' } : cell.value;
            const target = row.getCell(cell.col);
            target.value =
              cell.formula && cell.formula.text !== null ? (plain === null ? { formula: cell.formula.text } : { formula: cell.formula.text, result: plain as number }) : cell.richText ? richTextOut(cell.richText) : plain;
            // styles: 'cache' gives cells their style (worksheet-reader.js L282–286).
            if (styles) target.style = styleOut(cell.style, cell.numFmt);
          }
          if (r.height !== null) row.height = r.height;
          if (r.hidden) row.hidden = true;
          sheet._rows[r.number - 1] = undefined;
          yield row;
        }
      };
      yield { eventType: 'worksheet', value: new WorksheetReader(this, id, info.name, info.state, sheet, rows) };
    }
  }

  async *[Symbol.asyncIterator](): AsyncIterator<WorksheetReader> {
    for await (const { eventType, value } of this.parse()) if (eventType === 'worksheet') yield value as WorksheetReader;
  }

  /**
   * Reads the workbook and emits events, as exceljs `read()`: `worksheet` for each worksheet
   * (its rows are read after the listeners run; listen to the worksheet's `row` events), then
   * `end` and `finished`; `error` on failure.
   */
  async read(input?: string | Uint8Array | ArrayBuffer | AsyncIterable<Uint8Array>, options?: WorkbookReaderOptions): Promise<void> {
    try {
      for await (const { eventType, value } of this.parse(input, options)) {
        if (eventType === 'worksheet') {
          this.emit('worksheet', value);
          await (value as WorksheetReader).read();
        }
      }
      this.emit('end');
      this.emit('finished');
    } catch (error) {
      this.emit('error', error);
    }
  }
}
