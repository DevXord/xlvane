// exceljs `Workbook` (lib/doc/workbook.js) and its `xlsx` reader/writer (lib/xlsx/xlsx.js), built
// on xlvane (docs/decisions.md D-503 … D-506).
import { XlvError, writeWorkbook, type Workbook as XlvWorkbook } from 'xlvane';
import { readBytes, richTextOut, styleOut, toNewWorkbook, type WorkbookMeta } from './convert';
import { DefinedNames, extractRanges } from './defined-names';
import { notSupported, Worksheet } from './model';
import type { CellValue, WorksheetOptions } from './types';

/** Node.js Buffer, when available, so `writeBuffer()` returns what exceljs returns there. */
interface BufferCtor {
  from(data: ArrayBuffer, offset: number, length: number): Uint8Array;
}
const NodeBuffer: BufferCtor | undefined = (globalThis as { Buffer?: BufferCtor }).Buffer;

function asBuffer(bytes: Uint8Array): Uint8Array {
  return NodeBuffer ? NodeBuffer.from(bytes.buffer as ArrayBuffer, bytes.byteOffset, bytes.byteLength) : bytes;
}

function toBytes(data: unknown): Uint8Array {
  if (data instanceof Uint8Array) return data;
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  if (ArrayBuffer.isView(data)) return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  throw new XlvError('XLV_E_INVALID_INPUT', 'workbook.xlsx.load() expects a Buffer, Uint8Array or ArrayBuffer.', {
    hint: 'Pass the file contents, e.g. fs.readFileSync(path), or use workbook.xlsx.readFile(path).',
  });
}

/** A writable Node.js stream, as accepted by `xlsx.write()`. */
interface WritableLike {
  write(chunk: Uint8Array): unknown;
  end(callback?: () => void): unknown;
  once?(event: string, listener: (error?: unknown) => void): unknown;
}

/**
 * The `xlsx` property of a {@link Workbook}: reading and writing XLSX files.
 *
 * @example
 * ```ts
 * import ExcelJS from '@xlvane/exceljs';
 *
 * const workbook = new ExcelJS.Workbook();
 * workbook.addWorksheet('Data').getCell('A1').value = 'Hello';
 * const buffer = await workbook.xlsx.writeBuffer();
 * const copy = await new ExcelJS.Workbook().xlsx.load(buffer);
 * console.log(copy.getWorksheet('Data')?.getCell('A1').value); // Hello
 * ```
 */
export class XlsxIO {
  /** @internal */
  constructor(private readonly workbook: Workbook) {}

  /** Reads XLSX bytes into the workbook (its sheets are replaced) and returns the workbook. */
  async load(data: Uint8Array | ArrayBuffer, _options?: unknown): Promise<Workbook> {
    this.workbook._fill(readBytes(toBytes(data)));
    return this.workbook;
  }

  /** Reads an XLSX file (Node.js). */
  async readFile(path: string, _options?: unknown): Promise<Workbook> {
    const { readWorkbookFile } = await import('xlvane/node');
    this.workbook._fill(await readWorkbookFile(path));
    return this.workbook;
  }

  /** Reads XLSX data from a readable stream (Node.js stream or any async iterable of bytes). */
  async read(stream: AsyncIterable<Uint8Array>, _options?: unknown): Promise<Workbook> {
    const chunks: Uint8Array[] = [];
    let size = 0;
    for await (const chunk of stream) {
      chunks.push(chunk);
      size += chunk.length;
    }
    const all = new Uint8Array(size);
    let at = 0;
    for (const c of chunks) {
      all.set(c, at);
      at += c.length;
    }
    return this.load(all);
  }

  /** Writes the workbook and returns the file contents (a Buffer in Node.js). */
  async writeBuffer(_options?: unknown): Promise<Uint8Array> {
    return asBuffer(this.bytes());
  }

  /** Writes the workbook to a file, replacing an existing file as exceljs does (Node.js). */
  async writeFile(path: string, _options?: unknown): Promise<void> {
    const { writeWorkbookFile } = await import('xlvane/node');
    await writeWorkbookFile(toNewWorkbook(this.workbook.worksheets, this.workbook._date1904, this.workbook._meta()), path, { overwrite: true });
  }

  /** Writes the workbook to a writable stream and ends it. */
  async write(stream: WritableLike, _options?: unknown): Promise<void> {
    const bytes = asBuffer(this.bytes());
    stream.write(bytes);
    await new Promise<void>((resolve) => {
      stream.end(() => resolve());
    });
  }

  private bytes(): Uint8Array {
    return writeWorkbook(toNewWorkbook(this.workbook.worksheets, this.workbook._date1904, this.workbook._meta()));
  }
}

/**
 * A workbook, as exceljs `Workbook`. Use `xlsx` to read and write files.
 *
 * @example
 * ```ts
 * import ExcelJS from '@xlvane/exceljs';
 *
 * const workbook = new ExcelJS.Workbook();
 * workbook.creator = 'Me';
 * const sheet = workbook.addWorksheet('My Sheet');
 * sheet.addRow(['Name', 'Total']);
 * sheet.addRow(['Tea', 3.5]);
 * await workbook.xlsx.writeFile('report.xlsx');
 * ```
 */
export class Workbook {
  /** Written to `docProps/core.xml`; `"Unknown"` when not set, as exceljs (lib/doc/workbook.js L155). */
  creator: string | undefined;
  lastModifiedBy: string | undefined;
  created: Date | undefined = new Date();
  modified: Date | undefined = this.created;
  lastPrinted: Date | undefined;
  title: string | undefined = '';
  subject: string | undefined = '';
  keywords: string | undefined = '';
  category: string | undefined = '';
  description: string | undefined = '';
  language: string | undefined;
  revision: string | number | undefined;
  contentStatus: string | undefined;
  /** Kept on the object; not written yet (it belongs to `docProps/app.xml`). */
  company: string | undefined = '';
  /** Kept on the object; not written yet (it belongs to `docProps/app.xml`). */
  manager: string | undefined = '';
  /** Workbook properties; `date1904: true` selects the 1904 date system. */
  properties: { date1904?: boolean; [key: string]: unknown } = {};
  calcProperties: Record<string, unknown> = {};
  views: Record<string, unknown>[] = [];
  /** @internal */
  _sheets: (Worksheet | undefined)[] = [];
  private _xlsx: XlsxIO | undefined;
  private order = 0;
  private readonly orders = new Map<Worksheet, number>();

  get xlsx(): XlsxIO {
    this._xlsx ??= new XlsxIO(this);
    return this._xlsx;
  }

  get csv(): never {
    return notSupported('CSV (workbook.csv); use toCSV/fromCSV from xlvane');
  }

  private readonly _definedNames = new DefinedNames();

  /** Defined names, as exceljs `workbook.definedNames`. */
  get definedNames(): DefinedNames {
    return this._definedNames;
  }

  /** @internal The document properties and defined names to write, as exceljs `workbook.model`. */
  _meta(): WorkbookMeta {
    const text = (v: unknown): string | undefined => (v === undefined || v === null ? undefined : String(v));
    const entries: [string, unknown][] = [
      ['creator', this.creator || 'Unknown'],
      ['lastModifiedBy', this.lastModifiedBy || 'Unknown'],
      ['title', text(this.title)],
      ['subject', text(this.subject)],
      ['description', text(this.description)],
      ['keywords', text(this.keywords)],
      ['category', text(this.category)],
      ['language', text(this.language)],
      ['revision', text(this.revision)],
      ['contentStatus', text(this.contentStatus)],
      // xlsx.js prepareModel (L597): created and modified default to now.
      ['created', this.created ?? new Date()],
      ['modified', this.modified ?? new Date()],
      ['lastPrinted', this.lastPrinted],
    ];
    const properties = Object.fromEntries(entries.filter(([, v]) => v !== undefined));
    return { properties, definedNames: this._definedNames.model };
  }

  /** @internal */
  get _date1904(): boolean {
    return this.properties.date1904 === true;
  }

  /** @internal Names of the other sheets, for the duplicate check of Worksheet.name. */
  readonly _names = (exclude: Worksheet): string[] => this.worksheets.filter((w) => w !== exclude).map((w) => w.name);

  private get nextId(): number {
    for (let i = 1; i < this._sheets.length; i++) if (!this._sheets[i]) return i;
    return this._sheets.length || 1;
  }

  /** Adds a worksheet. Names follow exceljs rules (31 characters at most; longer names are cut). */
  addWorksheet(name?: string, options?: WorksheetOptions): Worksheet {
    const id = this.nextId;
    const sheet = new Worksheet(this, id, name, options);
    this._sheets[id] = sheet;
    this.orders.set(sheet, ++this.order);
    return sheet;
  }

  removeWorksheet(id: number | string): void {
    const sheet = this.getWorksheet(id);
    if (sheet) {
      this._sheets[sheet.id] = undefined;
      this.orders.delete(sheet);
    }
  }

  /** A worksheet by id or name; the first worksheet without an argument; `undefined` if none. */
  getWorksheet(id?: number | string): Worksheet | undefined {
    if (id === undefined) return this.worksheets[0];
    if (typeof id === 'number') return this._sheets[id];
    if (typeof id === 'string') return this.worksheets.find((w) => w.name === id);
    return undefined;
  }

  /** Worksheets in workbook order. */
  get worksheets(): Worksheet[] {
    return this._sheets.filter((w): w is Worksheet => w !== undefined).sort((a, b) => (this.orders.get(a) ?? 0) - (this.orders.get(b) ?? 0));
  }

  eachSheet(iteratee: (sheet: Worksheet, id: number) => void): void {
    for (const sheet of this.worksheets) iteratee(sheet, sheet.id);
  }

  addImage(): never {
    return notSupported('Images (workbook.addImage)');
  }

  /** @internal Replaces the sheets with those of a workbook read by xlvane. */
  _fill(read: XlvWorkbook): void {
    this._sheets = [];
    this.orders.clear();
    this.order = 0;
    if (read.dateSystem === '1904') this.properties.date1904 = true;
    else delete this.properties.date1904;
    // exceljs assigns every core property from the file, so missing ones become undefined (workbook.js L184).
    const p = read.properties;
    this.creator = p.creator;
    this.lastModifiedBy = p.lastModifiedBy;
    this.lastPrinted = p.lastPrinted;
    this.created = p.created;
    this.modified = p.modified;
    this.title = p.title;
    this.subject = p.subject;
    this.keywords = p.keywords;
    this.category = p.category;
    this.description = p.description;
    this.language = p.language;
    // cp:revision is read as an integer (core-xform.js L22, IntegerXform).
    this.revision = p.revision === undefined ? undefined : Number.parseInt(p.revision, 10);
    this.contentStatus = p.contentStatus;
    // company and manager come from docProps/app.xml, which xlvane does not read yet.
    this.company = undefined;
    this.manager = undefined;
    // Print areas and titles become pageSetup in exceljs (workbook-xform.js L187), which is not
    // supported; other built-in names are written by xlvane itself (autofilter). The scope
    // (localSheetId) is dropped, as in exceljs.
    this._definedNames.model = read.definedNames.filter((d) => !d.name.startsWith('_xlnm.')).map((d) => ({ name: d.name, ranges: extractRanges(d.refersTo) }));
    for (const s of read.sheets) {
      if (s.kind !== 'worksheet') continue;
      const ws = this.addWorksheet(s.name, { state: s.state });
      for (const c of s.columns) {
        const last = Math.min(c.max, 16_384);
        for (let n = c.min; n <= last; n++) {
          const column = ws.getColumn(n);
          if (c.width !== null) column.width = c.width;
          if (c.hidden) column.hidden = true;
          if (c.style) column.style = styleOut(c.style, typeof c.style.numFmt === 'string' ? c.style.numFmt : null);
        }
      }
      for (const cell of s.cells()) {
        const target = ws.getCell(cell.row, cell.col);
        let value: CellValue;
        const plain = cell.type === 'error' ? { error: cell.value as '#N/A' } : cell.value;
        if (cell.formula && cell.formula.text !== null) {
          value = plain === null ? { formula: cell.formula.text } : { formula: cell.formula.text, result: plain as number };
        } else value = cell.richText ? richTextOut(cell.richText) : plain;
        target.value = value;
        target.style = styleOut(cell.style, cell.numFmt);
        const info = s.rowInfo(cell.row);
        if (info) {
          const row = ws.getRow(cell.row);
          if (info.height !== null) row.height = info.height;
          if (info.hidden) row.hidden = true;
        }
      }
      for (const m of s.merges) if (m.includes(':')) ws.mergeCells(m);
      if (s.freeze) {
        const view: Record<string, unknown> = { state: 'frozen' };
        if (s.freeze.cols > 0) view['xSplit'] = s.freeze.cols;
        if (s.freeze.rows > 0) view['ySplit'] = s.freeze.rows;
        ws.views = [view];
      }
      if (s.autoFilter) ws.autoFilter = s.autoFilter;
    }
  }
}
