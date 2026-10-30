// exceljs streaming API: `stream.xlsx.WorkbookWriter` (lib/stream/xlsx/workbook-writer.js) and
// `stream.xlsx.WorkbookReader` (lib/stream/xlsx/workbook-reader.js), built on xlvane's
// writeWorkbookStream and openWorkbookStream (docs/decisions.md D-507).
import { openWorkbookStream, writeWorkbookStream, type CellInput, type CellStyle, type StreamRowInput, type StreamSheetSpec, type WorkbookProperties, type WriteWorkbookStreamOptions } from 'xlvane';
import { autoFilterRange, cellStyleIn, columnIn, isDateCell, DATE_NUMFMT, layoutIn, richTextOut, styleIn, valueIn } from './convert';
import type { DefinedNames } from './defined-names';
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
    };
  }

  override mergeCells(...args: unknown[]): void {
    if (this.started) notSupported('mergeCells() after rows of a streamed worksheet were committed');
    super.mergeCells(...args);
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
    let styled = false;
    row._cells.forEach((cell, i) => {
      if (!cell) return;
      // TODO(phase A2, notes fork): links are set per cell while rows are committed, but the core
      // streaming writer takes a sheet's links up front (sheets[].hyperlinks). Wire them through
      // per-row links of StreamRowInput once the core has them; until then a link would be lost,
      // so it is refused rather than dropped.
      if (cell.type === ValueType.Hyperlink) notSupported(`Hyperlinks in streamed worksheets (cell ${cell.address})`);
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
      const error = notSupportedLater(`Changing the layout of streamed worksheet "${this.name}" (columns, views, autoFilter, merges) after rows were committed`);
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
    const spec: WriteWorkbookStreamOptions = { sheets: this.sheets, sharedStrings: this.options.useSharedStrings === true, ...meta };
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

function notSupportedLater(feature: string): Error {
  try {
    notSupported(feature);
  } catch (error) {
    return error as Error;
  }
  return new Error(feature);
}

/** A worksheet read by {@link WorkbookReader}. */
interface ReaderWorksheet extends AsyncIterable<Row> {
  readonly name: string;
  readonly id: number;
}

/**
 * The streaming reader of exceljs (`ExcelJS.stream.xlsx.WorkbookReader`): iterate worksheets
 * and their rows with `for await`. Each row is a `Row` with cells; rows are not kept.
 *
 * @example
 * ```ts
 * import ExcelJS from '@xlvane/exceljs';
 *
 * const reader = new ExcelJS.stream.xlsx.WorkbookReader('big.xlsx', {});
 * for await (const worksheet of reader) {
 *   for await (const row of worksheet) console.log(row.number, row.values);
 * }
 * ```
 */
export class WorkbookReader implements AsyncIterable<ReaderWorksheet> {
  constructor(
    private readonly input: string | Uint8Array | ArrayBuffer | AsyncIterable<Uint8Array>,
    private readonly options: Record<string, unknown> = {},
  ) {}

  private async open(): Promise<Awaited<ReturnType<typeof openWorkbookStream>>> {
    const input = this.input;
    if (typeof input === 'string') {
      const { openWorkbookStreamFile } = await import('xlvane/node');
      return openWorkbookStreamFile(input);
    }
    if (input instanceof Uint8Array || input instanceof ArrayBuffer) return openWorkbookStream(input);
    const chunks: Uint8Array[] = [];
    for await (const c of input) chunks.push(c);
    return openWorkbookStream(new Blob(chunks as BlobPart[]));
  }

  async *[Symbol.asyncIterator](): AsyncIterator<ReaderWorksheet> {
    if (this.options['worksheets'] === 'ignore') return;
    const book = await this.open();
    const host = new Workbook();
    let id = 0;
    for (const info of book.sheets) {
      if (info.kind !== 'worksheet') continue;
      id++;
      const sheet = host.addWorksheet(info.name, { state: info.state });
      const name = info.name;
      yield {
        name,
        id,
        async *[Symbol.asyncIterator](): AsyncIterator<Row> {
          for await (const r of book.rows(name)) {
            const row = sheet.getRow(r.number);
            for (const cell of r.cells) {
              const plain = cell.type === 'error' ? { error: cell.value as '#N/A' } : cell.value;
              row.getCell(cell.col).value =
                cell.formula && cell.formula.text !== null ? (plain === null ? { formula: cell.formula.text } : { formula: cell.formula.text, result: plain as number }) : cell.richText ? richTextOut(cell.richText) : plain;
            }
            if (r.height !== null) row.height = r.height;
            if (r.hidden) row.hidden = true;
            sheet._rows[r.number - 1] = undefined;
            yield row;
          }
        },
      };
    }
  }

  /** exceljs event API (`reader.read()` with `on('worksheet')`) is not supported. */
  read(): never {
    return notSupported("WorkbookReader.read() and its events (use 'for await' instead)");
  }
}
