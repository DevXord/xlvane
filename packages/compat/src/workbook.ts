// exceljs `Workbook` (lib/doc/workbook.js) and its `xlsx` reader/writer (lib/xlsx/xlsx.js), built
// on xlvane (docs/decisions.md D-503 … D-506).
import { DEFAULT_READ_LIMITS, XlvError, writeWorkbook, type CellStyle, type FontStyle, type Workbook as XlvWorkbook, type WorkbookImage } from 'xlvane';
import { colorOut, sameStyleData, writableDefaultFont, formulaOut, hasFormulaText, noteOut, readBytes, readValueOut, richTextOut, styleOut, toNewWorkbook, viewOut, type ReadDates, type WorkbookMeta } from './convert';
import { CsvIO } from './csv';
import { DefinedNames, extractRanges, rangeCellCount } from './defined-names';
import { imageBytes, imageOut, loadImageFiles, mediaOut, tableOut, type ImageData } from './media';
import { checkSheetName, columnLetter, Worksheet } from './model';
import { readFileBytes, xlvaneNode } from './node-api';
import { conditionalFormatsOut, dataValidationsOut, headerFooterOut, pageSetupOut, protectionOut, rowBreaksOut, validationBudget } from './sheet-rules';
import type { CellValue, HeaderFooter, Style, WorksheetOptions } from './types';
import type { Image } from './compat-types';

/** Node.js Buffer, when available, so `writeBuffer()` returns what exceljs returns there. */
interface BufferCtor {
  from(data: ArrayBuffer, offset: number, length: number): Uint8Array;
}
const NodeBuffer: BufferCtor | undefined = (globalThis as { Buffer?: BufferCtor }).Buffer;

function asBuffer(bytes: Uint8Array): Uint8Array {
  return NodeBuffer ? NodeBuffer.from(bytes.buffer as ArrayBuffer, bytes.byteOffset, bytes.byteLength) : bytes;
}

/**
 * Options of `xlsx.load`, `read` and `readFile`, as exceljs `XlsxReadOptions` (index.d.ts L1460;
 * lib/xlsx/xlsx.js L257, lib/xlsx/xform/sheet/worksheet-xform.js L95, L372; X4-25).
 *
 * @example
 * ```ts
 * import type { XlsxReadOptions } from '@xlvane/compat';
 *
 * const options: XlsxReadOptions = { ignoreNodes: ['dataValidations'], maxRows: 100_000 };
 * ```
 */
export interface XlsxReadOptions {
  /** Worksheet elements to skip, by element name (`sheetData`, `mergeCells`, `hyperlinks`, `dataValidations`, …). */
  ignoreNodes?: string[];
  /** Refuse a worksheet with more rows than this (`XLV_E_LIMIT_EXCEEDED`). */
  maxRows?: number;
  /** Refuse a row with more cells than this (`XLV_E_LIMIT_EXCEEDED`). */
  maxCols?: number;
  /** `load` only: the data is a base64 string. */
  base64?: boolean;
}

/**
 * Options of `xlsx.writeBuffer`, `write` and `writeFile`, as exceljs `XlsxWriteOptions`
 * (index.d.ts L1447–1472; lib/xlsx/xlsx.js L602–603). `useSharedStrings: false` writes inline
 * strings; `zip: { compression: 'STORE' }` or `zip: { compressionOptions: { level } }` sets the
 * Deflate level (0–9).
 *
 * @example
 * ```ts
 * import type { XlsxWriteOptions } from '@xlvane/compat';
 *
 * const options: XlsxWriteOptions = { useStyles: false };
 * ```
 */
export interface XlsxWriteOptions {
  /** `false` writes no cell or column styles; dates keep format 14, as in exceljs. */
  useStyles?: boolean;
  useSharedStrings?: boolean;
  zip?: Record<string, unknown>;
}

/**
 * The core writer options for exceljs `useSharedStrings` (default true) and
 * `zip.compressionOptions.level` (xlsx.js L602–603, JSZip options; X4-25).
 */
export function writeOptions(options: XlsxWriteOptions | undefined): { sharedStrings: boolean; compressionLevel?: number } {
  const zip = (options?.zip ?? {}) as { compression?: unknown; compressionOptions?: { level?: unknown } };
  const level = zip.compressionOptions?.level;
  const compressionLevel = zip.compression === 'STORE' ? 0 : typeof level === 'number' && Number.isInteger(level) && level >= 0 && level <= 9 ? level : undefined;
  return { sharedStrings: options?.useSharedStrings !== false, ...(compressionLevel === undefined ? {} : { compressionLevel }) };
}

function readOptions(options: unknown): XlsxReadOptions {
  if (options === undefined || options === null) return {};
  if (typeof options !== 'object') throw new XlvError('XLV_E_INVALID_OPTION', 'The read options must be an object.', { hint: 'For example { maxRows: 100000 }.', details: { option: 'options' } });
  const o = options as Record<string, unknown>;
  const limit = (name: 'maxRows' | 'maxCols'): number | undefined => {
    const v = o[name];
    if (v === undefined || v === null || v === 0) return undefined;
    if (typeof v !== 'number' || !Number.isInteger(v) || v < 1) {
      throw new XlvError('XLV_E_INVALID_OPTION', `The option "${name}" must be a whole number of at least 1.`, { hint: `For example { ${name}: 100000 }.`, details: { option: name } });
    }
    return v;
  };
  const ignore = o['ignoreNodes'];
  if (ignore !== undefined && (!Array.isArray(ignore) || !ignore.every((n) => typeof n === 'string'))) {
    throw new XlvError('XLV_E_INVALID_OPTION', 'The option "ignoreNodes" must be an array of element names.', { hint: "For example { ignoreNodes: ['dataValidations'] }.", details: { option: 'ignoreNodes' } });
  }
  const maxRows = limit('maxRows');
  const maxCols = limit('maxCols');
  return { ...(ignore ? { ignoreNodes: ignore as string[] } : {}), ...(maxRows ? { maxRows } : {}), ...(maxCols ? { maxCols } : {}), ...(o['base64'] === true ? { base64: true } : {}) };
}

function fromBase64(data: unknown): Uint8Array {
  const text = typeof data === 'string' ? data : data instanceof Uint8Array ? new TextDecoder().decode(data) : String(data);
  let binary: string;
  try {
    binary = atob(text.replace(/\s+/g, ''));
  } catch (cause) {
    throw new XlvError('XLV_E_INVALID_INPUT', 'workbook.xlsx.load(data, { base64: true }) got data that is not base64.', { hint: 'Pass the base64 text of the file, or drop the base64 option for bytes.', cause });
  }
  const out = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
  return out;
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
 * import ExcelJS from '@xlvane/compat';
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
  async load(data: Uint8Array | ArrayBuffer | string, options?: XlsxReadOptions): Promise<Workbook> {
    const o = readOptions(options);
    this.workbook._fill(readBytes(o.base64 ? fromBase64(data) : toBytes(data)), o);
    return this.workbook;
  }

  /** Reads an XLSX file (Node.js). */
  async readFile(path: string, options?: XlsxReadOptions): Promise<Workbook> {
    const o = readOptions(options);
    const { readWorkbookFile } = await xlvaneNode('xlsx.readFile()');
    this.workbook._fill(await readWorkbookFile(path), o);
    return this.workbook;
  }

  /** Reads XLSX data from a readable stream (Node.js stream or any async iterable of bytes). */
  async read(stream: AsyncIterable<Uint8Array>, options?: XlsxReadOptions): Promise<Workbook> {
    const { chunks, size } = await collectInput(stream, 'xlsx.read()');
    const all = new Uint8Array(size);
    let at = 0;
    for (const c of chunks) {
      all.set(c, at);
      at += c.length;
    }
    return this.load(all, { ...readOptions(options), base64: false });
  }

  /** Writes the workbook and returns the file contents (a Buffer in Node.js). */
  async writeBuffer(options?: XlsxWriteOptions): Promise<Uint8Array> {
    await this.loadFiles();
    return asBuffer(this.bytes(options));
  }

  /** Writes the workbook to a file, replacing an existing file as exceljs does (Node.js). */
  async writeFile(path: string, options?: XlsxWriteOptions): Promise<void> {
    const { writeWorkbookFile } = await xlvaneNode('xlsx.writeFile()');
    await this.loadFiles();
    await writeWorkbookFile(this.book(options), path, { overwrite: true, ...writeOptions(options) });
  }

  /** Writes the workbook to a writable stream and ends it. */
  async write(stream: WritableLike, options?: XlsxWriteOptions): Promise<void> {
    await this.loadFiles();
    const bytes = asBuffer(this.bytes(options));
    stream.write(bytes);
    await new Promise<void>((resolve) => {
      stream.end(() => resolve());
    });
  }

  /** Removed in exceljs 4 (it throws there too); use `read`. */
  createInputStream(): never {
    throw new Error(
      '`XLSX#createInputStream` is deprecated. You should use `XLSX#read` instead. This method will be removed in version 5.0. Please follow upgrade instruction: https://github.com/exceljs/exceljs/blob/master/UPGRADE-4.0.md',
    );
  }

  /** Reads image files given by `filename`, as exceljs does when it writes (xlsx.js L426–429). */
  private loadFiles(): Promise<void> {
    return loadImageFiles(this.workbook._media, (path) => readFileBytes(path, 'workbook.addImage({ filename })'));
  }

  private book(options: XlsxWriteOptions | undefined) {
    const useStyles = options?.useStyles !== false;
    const font = this.workbook._defaultFont;
    return toNewWorkbook(this.workbook.worksheets, this.workbook._date1904, { ...this.workbook._meta(), useStyles, ...(font ? { defaultFont: font } : {}) });
  }

  private bytes(options: XlsxWriteOptions | undefined): Uint8Array {
    return writeWorkbook(this.book(options), writeOptions(options));
  }
}

/**
 * A workbook, as exceljs `Workbook`. Use `xlsx` to read and write files.
 *
 * @example
 * ```ts
 * import ExcelJS from '@xlvane/compat';
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
  /** Written to `docProps/app.xml`, always (`company || ''`), as exceljs (X4-13). */
  company: string | undefined = '';
  /** Written to `docProps/app.xml` when defined, as exceljs (X4-13). */
  manager: string | undefined = '';
  /** Workbook properties; `date1904: true` selects the 1904 date system. */
  properties: { date1904?: boolean; [key: string]: unknown } = {};
  calcProperties: Record<string, unknown> = {};
  views: Record<string, unknown>[] = [];
  /** @internal */
  _sheets: (Worksheet | undefined)[] = [];
  private _xlsx: XlsxIO | undefined;
  private _csv: CsvIO | undefined;

  get xlsx(): XlsxIO {
    this._xlsx ??= new XlsxIO(this);
    return this._xlsx;
  }

  /** CSV reading and writing, as exceljs `workbook.csv` (X4-23, X4-24). */
  get csv(): CsvIO {
    this._csv ??= new CsvIO(this);
    return this._csv;
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
      // app-xform.js L31–32: Company is always written (`company || ''`), Manager when defined (X4-13).
      ['company', this.company || ''],
      ['manager', text(this.manager)],
    ];
    const properties = Object.fromEntries(entries.filter(([, v]) => v !== undefined));
    return { properties, definedNames: this._definedNames.model, media: this._media, views: this.views, calcProperties: this.calcProperties };
  }

  /** @internal The default font (font 0) of the file this workbook was loaded from, written again (F4); `null` for a workbook made in code. */
  _defaultFont: FontStyle | null = null;

  /** @internal */
  get _date1904(): boolean {
    return this.properties.date1904 === true;
  }

  // Indexes of `_sheets`, so that adding, naming and finding sheets takes constant time and a
  // workbook of many sheets loads in linear time (security review J, C3): the sheets by lower-case
  // name, a position below which every id is taken, and the largest orderNo (recomputed as exceljs
  // computes it when a sheet holding it was removed or lowered).
  private _byName = new Map<string, Worksheet>();
  private _firstFree = 1;
  private _maxOrderNo = 0;
  private _maxOrderNoStale = false;

  /** @internal Whether a sheet other than `exclude` has this name, ignoring case (the duplicate check of Worksheet.name). */
  _nameTaken(name: string, exclude?: Worksheet): boolean {
    const sheet = this._byName.get(name.toLowerCase());
    return sheet !== undefined && sheet !== exclude;
  }

  /** @internal Keeps the name index in step with a renamed sheet. */
  _renamed(sheet: Worksheet, oldName: string): void {
    if (this._sheets[sheet.id] !== sheet) return;
    const old = oldName.toLowerCase();
    if (this._byName.get(old) === sheet) this._byName.delete(old);
    this._byName.set(sheet.name.toLowerCase(), sheet);
  }

  /** @internal Keeps the largest orderNo in step with a changed one. */
  _orderNoChanged(sheet: Worksheet, oldOrderNo: number | undefined): void {
    if (this._sheets[sheet.id] !== sheet || this._maxOrderNoStale) return;
    const value = sheet.orderNo;
    const plain = (v: unknown): v is number | undefined => v === undefined || (typeof v === 'number' && !Number.isNaN(v));
    if (!plain(value) || !plain(oldOrderNo)) this._maxOrderNoStale = true;
    else if (value !== undefined && value > this._maxOrderNo) this._maxOrderNo = value;
    else if (oldOrderNo !== undefined && oldOrderNo >= this._maxOrderNo && (value === undefined || value < oldOrderNo)) this._maxOrderNoStale = true;
  }

  /** @internal Puts a worksheet into `_sheets` at its id and indexes it. */
  _putSheet(sheet: Worksheet): void {
    const previous = this._sheets[sheet.id];
    if (previous && previous !== sheet) this._dropSheet(previous);
    this._sheets[sheet.id] = sheet;
    this._byName.set(sheet.name.toLowerCase(), sheet);
    this._orderNoChanged(sheet, undefined);
  }

  /** Takes a worksheet out of `_sheets` and its indexes. */
  private _dropSheet(sheet: Worksheet): void {
    this._sheets[sheet.id] = undefined;
    const key = sheet.name.toLowerCase();
    if (this._byName.get(key) === sheet) this._byName.delete(key);
    if (sheet.id < this._firstFree) this._firstFree = Math.max(1, sheet.id);
    const order = sheet.orderNo;
    if (order !== undefined && !(typeof order === 'number' && order < this._maxOrderNo)) this._maxOrderNoStale = true;
  }

  /** Empties `_sheets` and its indexes. */
  private _clearSheets(): void {
    this._sheets = [];
    this._byName = new Map();
    this._firstFree = 1;
    this._maxOrderNo = 0;
    this._maxOrderNoStale = false;
  }

  /** The largest orderNo of the sheets, as exceljs computes it (lib/doc/workbook.js L81-85). */
  private _lastOrderNo(): number {
    if (this._maxOrderNoStale) {
      this._maxOrderNo = this._sheets.reduce((acc, ws) => (ws?.orderNo !== undefined && ws.orderNo > acc ? ws.orderNo : acc), 0);
      this._maxOrderNoStale = false;
    }
    return this._maxOrderNo;
  }

  /** The id the next worksheet gets (lib/doc/workbook.js `nextId`): the lowest free id. */
  get nextId(): number {
    const sheets = this._sheets;
    let i = Math.max(1, this._firstFree);
    while (i < sheets.length && sheets[i]) i++;
    this._firstFree = i;
    return i < sheets.length ? i : sheets.length || 1;
  }

  /** Adds a worksheet. Names follow exceljs rules (31 characters at most; longer names are cut). */
  addWorksheet(name?: string, options?: WorksheetOptions): Worksheet {
    return this._addWorksheetWithId(this.nextId, name, options);
  }

  /** @internal Adds a worksheet with a given id (a loaded file's sheetId, a model's id; D-F03). */
  _addWorksheetWithId(id: number, name?: string, options?: WorksheetOptions): Worksheet {
    const sheet = new Worksheet(this, id, name, options);
    // exceljs: one more than the largest orderNo (lib/doc/workbook.js L81-85).
    const last = this._lastOrderNo();
    this._putSheet(sheet);
    sheet.orderNo = last + 1;
    return sheet;
  }

  /**
   * @internal Adds a worksheet read from a file (`xlsx.load`, `stream.xlsx.WorkbookReader`). A
   * name exceljs refuses gives an XlvError with exceljs's message, the sheet and a hint, instead
   * of the plain Error of `addWorksheet`; an empty name is refused too.
   */
  _addReadWorksheet(id: number, name: string, options?: WorksheetOptions): Worksheet {
    const checked = checkSheetName(name, (n) => this._nameTaken(n));
    if (checked.message !== undefined) {
      const existing = checked.duplicate ? this._byName.get(checked.name.toLowerCase())?.name : undefined;
      throw new XlvError(checked.duplicate ? 'XLV_E_DUPLICATE_SHEET_NAME' : 'XLV_E_INVALID_SHEET_NAME', checked.message, {
        location: { sheet: name },
        hint: checked.duplicate
          ? 'Two worksheets of the file have the same name, ignoring case and after the first 31 characters, which exceljs does not accept. Rename one in a spreadsheet application, or read the file with xlvane\'s readWorkbook.'
          : 'The file names a worksheet in a way exceljs does not accept. Rename the sheet in a spreadsheet application, or read the file with xlvane\'s readWorkbook.',
        details: existing === undefined ? { name } : { name, existing },
      });
    }
    return this._addWorksheetWithId(id, name, options);
  }

  removeWorksheet(id: number | string): void {
    const sheet = this.getWorksheet(id);
    if (sheet) this._dropSheet(sheet);
  }

  /** Removes a worksheet by id or name, as `removeWorksheet` (lib/doc/workbook.js L77). */
  removeWorksheetEx(worksheet: Worksheet): void {
    this.removeWorksheet(worksheet.id);
  }

  /**
   * exceljs keeps the theme of a loaded file and `clearThemes()` drops it (lib/doc/workbook.js
   * L137). xlvane never keeps or writes a theme part, so there is nothing to clear.
   */
  clearThemes(): void {}

  /**
   * A plain-data snapshot of the workbook, as exceljs `workbook.model` (lib/doc/workbook.js L153):
   * the document properties, `views`, `calcProperties`, the defined names, the media and the
   * worksheets' `model`. The setter rebuilds the workbook from such a snapshot (X4-30).
   */
  get model(): Record<string, unknown> {
    return {
      creator: this.creator || 'Unknown',
      lastModifiedBy: this.lastModifiedBy || 'Unknown',
      lastPrinted: this.lastPrinted,
      created: this.created,
      modified: this.modified,
      properties: { ...this.properties },
      worksheets: this.worksheets.map((ws) => ws.model),
      definedNames: this._definedNames.model,
      views: this.views,
      company: this.company,
      manager: this.manager,
      title: this.title,
      subject: this.subject,
      keywords: this.keywords,
      category: this.category,
      description: this.description,
      language: this.language,
      revision: this.revision,
      contentStatus: this.contentStatus,
      calcProperties: { ...this.calcProperties },
      media: [...this._media],
    };
  }

  set model(value: Record<string, unknown>) {
    for (const key of ['creator', 'lastModifiedBy', 'lastPrinted', 'created', 'modified', 'company', 'manager', 'title', 'subject', 'keywords', 'category', 'description', 'language', 'revision', 'contentStatus'] as const) {
      (this as Record<string, unknown>)[key] = value[key];
    }
    this.properties = { ...((value['properties'] as Record<string, unknown> | undefined) ?? {}) };
    this.calcProperties = { ...((value['calcProperties'] as Record<string, unknown> | undefined) ?? {}) };
    this.views = Array.isArray(value['views']) ? [...(value['views'] as Record<string, unknown>[])] : [];
    this._media = Array.isArray(value['media']) ? [...(value['media'] as ImageData[])] : [];
    for (const ws of this.worksheets) this.removeWorksheet(ws.id);
    const models = Array.isArray(value['worksheets']) ? (value['worksheets'] as Record<string, unknown>[]) : [];
    // The model's ids are kept, as exceljs keeps them (lib/doc/workbook.js L201-212); a missing or
    // repeated one takes the lowest free id (D-F03).
    const ids = freeIds(models.map((m) => m['id']));
    for (const [i, m] of models.entries()) {
      const ws = this._addWorksheetWithId(ids[i] ?? this.nextId, String(m['name'] ?? ''), { state: (m['state'] as 'visible') ?? 'visible' });
      ws.model = m;
    }
    this._definedNames.model = Array.isArray(value['definedNames']) ? (value['definedNames'] as { name: string; ranges: string[] }[]) : [];
  }

  /** A worksheet by id or name; the first worksheet without an argument; `undefined` if none. */
  getWorksheet(id?: number | string): Worksheet | undefined {
    if (id === undefined) return this.worksheets[0];
    if (typeof id === 'number') return this._sheets[id];
    if (typeof id === 'string') {
      const sheet = this._byName.get(id.toLowerCase());
      if (sheet?.name === id) return sheet;
      // Only sheets without a name (a model without names) can share an index entry.
      return id === '' ? this.worksheets.find((w) => w.name === id) : undefined;
    }
    return undefined;
  }

  /** Worksheets in workbook order. */
  get worksheets(): Worksheet[] {
    // Sorted by orderNo, as exceljs (lib/doc/workbook.js L119-124); ties keep id order.
    return this._sheets.filter((w): w is Worksheet => w !== undefined).sort((a, b) => (a.orderNo ?? 0) - (b.orderNo ?? 0));
  }

  eachSheet(iteratee: (sheet: Worksheet, id: number) => void): void {
    for (const sheet of this.worksheets) iteratee(sheet, sheet.id);
  }

  /** @internal Images of the workbook, as exceljs `workbook.media`. */
  _media: ImageData[] = [];

  /**
   * Adds an image (PNG, JPEG or GIF, or EMF or WMF as Excel stores them; recognised from its bytes) and returns its id for
   * `worksheet.addImage` and `addBackgroundImage`, as exceljs `workbook.addImage`. Give `buffer`,
   * `base64` or `filename`; a file is read when the workbook is written (Node.js, Bun, Deno), as
   * in exceljs.
   */
  addImage(image: ImageData): number {
    if (!(image && typeof image === 'object' && image.buffer === undefined && typeof image.base64 !== 'string' && typeof image.filename === 'string')) imageBytes(image);
    this._media.push({ ...image, type: 'image' });
    return this._media.length - 1;
  }

  /**
   * The image with this id, as exceljs `workbook.getImage`. Typed `Image` as in exceljs
   * (index.d.ts L1780); at run time an unknown id gives `undefined`, as in exceljs
   * (lib/doc/workbook.js L149-151).
   */
  getImage(id: number): Image {
    return this._media[id] as Image;
  }

  /**
   * @internal Replaces the sheets with those of a workbook read by xlvane. `options.ignoreNodes`
   * skips what exceljs skips for those worksheet elements, and `maxRows`/`maxCols` refuse
   * bigger worksheets (worksheet-xform.js L95, L372; list-xform.js L75; row-xform.js L119; X4-25).
   */
  _fill(read: XlvWorkbook, options: XlsxReadOptions = {}): void {
    const ignore = new Set(options.ignoreNodes ?? []);
    const use = (node: string): boolean => !ignore.has(node);
    // Merged ranges, table ranges and <col> ranges make one object per position. They count
    // against max(maxCells, 16 x cells read), the limit of the core's table and CSV output, so a
    // few bytes such as <mergeCell ref="A1:XFD1048576"/> cannot exhaust memory (D-H01).
    let cellsRead = 0;
    for (const s of read.sheets) cellsRead += s.cellCount;
    const dates: ReadDates = { date1904: read.dateSystem === '1904', strict: read.conformance === 'strict' };
    const budget = Math.max(DEFAULT_READ_LIMITS.maxCells, 16 * cellsRead);
    let spent = 0;
    // `sheet` is undefined for what belongs to the workbook (defined names).
    const spend = (positions: number, what: string, sheet: string | undefined): void => {
      spent += positions;
      if (spent <= budget) return;
      const of = sheet === undefined ? '' : ` of worksheet "${sheet}"`;
      throw new XlvError('XLV_E_LIMIT_EXCEEDED', `${what}${of} would create more than ${String(budget)} cells or columns for ${String(cellsRead)} cells in the file.`, {
        ...(sheet === undefined ? {} : { location: { sheet } }),
        hint:
          sheet === undefined
            ? 'The file declares names far larger than its data. Read it with xlvane\'s readWorkbook, which does not create these objects.'
            : 'The file declares ranges far larger than its data. Load it with ignoreNodes for that element (for example ignoreNodes: ["mergeCells"]), or read it with xlvane\'s readWorkbook, which does not create these objects.',
        details: { limit: 'maxCells', value: budget },
      });
    };
    // Merged ranges also have a smaller budget of their own: each merged cell carries the master's
    // style, and writing a styled cell costs a few hundred bytes, so 2 000 000 merged positions from
    // a 2.4 KB file took 1.4 GB to write and 4 000 000 ran out of memory (K2). A whole column
    // (1 048 576 cells) fits; a file with more cells may merge as many positions as it has cells.
    const mergeBudget = Math.max(MERGED_CELLS_BUDGET, cellsRead);
    let merged = 0;
    const spendMerge = (positions: number, ref: string, sheet: string): void => {
      merged += positions;
      if (merged <= mergeBudget) return;
      throw new XlvError('XLV_E_LIMIT_EXCEEDED', `The merged range ${ref} of worksheet "${sheet}" would bring the merged cells past ${String(mergeBudget)} for ${String(cellsRead)} cells in the file.`, {
        location: { sheet, cell: ref },
        hint: 'The file merges far more cells than it holds. Load it with ignoreNodes: ["mergeCells"] to skip merged ranges, or read it with xlvane\'s readWorkbook, which does not create these cells.',
        details: { limit: 'mergedCells', value: mergeBudget },
      });
    };
    // Rows looked up for their settings (height, visibility, outline) up to a sheet's dimension or
    // last cell: no object per row, but a lookup each, so a dimension of a million rows on many
    // sheets is limited too (D-H01): 64 Mi lookups or 16 per cell read, whichever is larger.
    const rowBudget = Math.max(64 * 1024 * 1024, 16 * cellsRead);
    // Data validations given to each cell they cover: 100 000 cells for the whole workbook (C2).
    const validations = validationBudget();
    let rowsLooked = 0;
    for (const s of read.sheets) {
      if (s.kind !== 'worksheet' || (options.maxRows === undefined && options.maxCols === undefined) || !use('sheetData')) continue;
      const perRow = new Map<number, number>();
      for (const cell of s.cells()) perRow.set(cell.row, (perRow.get(cell.row) ?? 0) + 1);
      const limit = (name: 'maxRows' | 'maxCols', max: number, value: number): never => {
        throw new XlvError('XLV_E_LIMIT_EXCEEDED', `Worksheet "${s.name}" has more ${name === 'maxRows' ? 'rows' : 'cells in a row'} than ${name} (${String(max)}).`, {
          location: { sheet: s.name },
          hint: `Raise the option, for example { ${name}: ${String(max * 2)} }, or leave it out.`,
          details: { limit: name, value: max, found: value },
        });
      };
      if (options.maxRows !== undefined && perRow.size > options.maxRows) limit('maxRows', options.maxRows, perRow.size);
      if (options.maxCols !== undefined) for (const count of perRow.values()) if (count > options.maxCols) limit('maxCols', options.maxCols, count);
    }
    this._clearSheets();
    // Font 0 of the file, written again as the workbook's default font (F4).
    // Only the parts the writer accepts, so that the workbook can be written again (review of 1.3.1).
    const font0 = writableDefaultFont(read.defaultFont);
    this._defaultFont = font0;
    // exceljs gives a style read from a file (a cell's s >= 1, a row or column style) the font its
    // fontId names, also font 0 (styles-xform.js L296-345), which the core leaves out (F4). A row
    // or column style counts as given when it has any part.
    const fileStyle = (style: CellStyle, numFmt: string | null, styled = Object.keys(style).length > 0): Style => styleOut(styled && font0 && !style.font ? { ...style, font: font0 } : style, numFmt);
    const lineFormat = (style: CellStyle): string | null => (typeof style.numFmt === 'string' ? style.numFmt : null);
    // exceljs ignores style id 0 of a column or row (col-xform.js L80, row-xform.js L130), which
    // LibreOffice writes on every column. The core gives no style index for them, so a column or
    // row style equal to xf 0 counts as id 0: xf 0 is the style of a cell without `s`, looked up
    // once, when a column or row has a style (review of 1.3.1).
    let xf0: CellStyle | null | undefined;
    const defaultXf = (): CellStyle | null => {
      if (xf0 !== undefined) return xf0;
      xf0 = null;
      for (const sheet of read.sheets) {
        if (sheet.kind !== 'worksheet') continue;
        for (const cell of sheet.cells()) {
          if (cell.styleIndex === 0) return (xf0 = cell.style);
        }
      }
      return xf0;
    };
    const lineStyle = (style: CellStyle): Style => {
      const zero = defaultXf();
      return zero && sameStyleData(style, zero) ? {} : fileStyle(style, lineFormat(style));
    };
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
    // xlsx.js L348–349: company and manager come from docProps/app.xml (X4-13).
    this.company = p.company;
    this.manager = p.manager;
    // Print areas and titles become pageSetup in exceljs (workbook-xform.js L187; read from
    // Sheet.printArea / printTitles below); other built-in names are written by xlvane itself
    // (autofilter). The scope (localSheetId) is dropped, as in exceljs.
    // Images: the media list as exceljs reads it, one entry per image part (X4-22).
    this._media = mediaOut(read.images);
    // bookViews and calcPr as exceljs reads them (workbook-view-xform.js L19–43; calcPr gives {}; X4-28).
    // The indexes count every sheet of the file; chart sheets are not loaded (D-R01), so an index
    // is mapped to the worksheet it falls on, or to the worksheet before a chart sheet.
    // worksheetsUpTo[k]: the worksheets among the first k + 1 sheets, counted once (C3).
    const worksheetsUpTo: number[] = [];
    for (const s of read.sheets) worksheetsUpTo.push((worksheetsUpTo[worksheetsUpTo.length - 1] ?? 0) + (s.kind === 'worksheet' ? 1 : 0));
    const sheetIndex = (i: number): number => (i < 0 || worksheetsUpTo.length === 0 ? 0 : Math.max((worksheetsUpTo[Math.min(i, worksheetsUpTo.length - 1)] ?? 0) - 1, 0));
    this.views = read.views.map((v) => {
      const view: Record<string, unknown> = { x: v.x ?? 0, y: v.y ?? 0, width: v.width ?? 25000, height: v.height ?? 10000, visibility: v.visibility };
      if (v.activeTab !== null) view['activeTab'] = sheetIndex(v.activeTab);
      if (v.firstSheet !== null) view['firstSheet'] = sheetIndex(v.firstSheet);
      return view;
    });
    this.calcProperties = read.calculation?.fullCalcOnLoad ? { fullCalcOnLoad: true } : {};
    const mediaIds = new Map<WorkbookImage, number>(read.images.map((img, i) => [img, i]));
    // A defined name is one object per cell, as in exceljs (cell-matrix.js), so its cells count
    // against the budget too; a range reaching past the sheet is dropped, as extractRanges drops
    // other invalid ranges (security review J, C1).
    this._definedNames.model = read.definedNames
      .filter((d) => !d.name.startsWith('_xlnm.'))
      .map((d) => {
        const ranges = extractRanges(d.refersTo).filter((range) => {
          const cells = rangeCellCount(range);
          if (cells === null) return false;
          spend(cells, `The defined name "${d.name.slice(0, 255)}"`, undefined);
          return true;
        });
        return { name: d.name, ranges };
      });
    // Worksheets keep the file's sheetId as id, as exceljs reads it (sheet-xform.js L18,
    // workbook-xform.js L161-186); a missing or repeated one takes the lowest free id (D-F03).
    const ids = freeIds(read.sheets.map((s) => (s.kind === 'worksheet' ? s.sheetId : undefined)));
    for (const [position, s] of read.sheets.entries()) {
      if (s.kind !== 'worksheet') continue;
      const ws = this._addReadWorksheet(ids[position] ?? this.nextId, s.name, { state: s.state });
      // exceljs: the sheet's 0-based position in the file (lib/doc/workbook.js L203).
      ws.orderNo = position;
      for (const c of use('cols') ? s.columns : []) {
        const last = Math.min(c.max, 16_384);
        spend(Math.max(0, last - c.min + 1), 'The column definitions', s.name);
        for (let n = c.min; n <= last; n++) {
          const column = ws.getColumn(n);
          if (c.width !== null) column.width = c.width;
          if (c.hidden) column.hidden = true;
          if (c.outlineLevel) column.outlineLevel = c.outlineLevel;
          if (c.style) column.style = lineStyle(c.style);
        }
      }
      // sheetFormatPr and outlinePr as exceljs reads them (sheet-format-properties-xform.js; X4-28).
      if (use('sheetFormatPr') || use('sheetPr')) {
        const props = s.properties;
        ws.properties['defaultRowHeight'] = props.defaultRowHeight;
        ws.properties['outlineLevelRow'] = props.outlineLevelRow;
        ws.properties['outlineLevelCol'] = props.outlineLevelCol;
        if (!props.summaryBelow || !props.summaryRight) ws.properties['outlineProperties'] = { summaryBelow: props.summaryBelow, summaryRight: props.summaryRight };
      }
      // Masters of shared formulas by group index, so dependents name their master (X4-29).
      const masters = new Map<number, string>();
      for (const cell of use('sheetData') ? s.cells() : []) {
        const f = cell.formula;
        if (f && f.kind === 'shared' && f.ref !== null && f.sharedIndex !== null) masters.set(f.sharedIndex, cell.ref);
      }
      // Cells that hold only a style keep it (D-E05), as exceljs reads them (cell-xform.js).
      let lastRow = 0;
      for (const cell of use('sheetData') ? s.styledCells() : []) {
        ws.getCell(cell.row, cell.col).style = fileStyle(cell.style, cell.numFmt, cell.styleIndex > 0);
        if (cell.row > lastRow) lastRow = cell.row;
      }
      for (const cell of use('sheetData') ? s.cells() : []) {
        const target = ws.getCell(cell.row, cell.col);
        let value: CellValue;
        // Time-only dates and exceljs's date formats (F6, F7; cell-xform.js L449-459).
        const plain = cell.type === 'error' ? { error: cell.value as '#N/A' } : (readValueOut(cell.value, cell.numFmt, dates) as CellValue);
        // hasFormulaText: a dependent's text is not built; formulaOut needs only its master (K2).
        if (cell.formula && hasFormulaText(cell.formula)) value = formulaOut(cell.formula, plain, cell.ref, masters);
        else value = cell.richText ? richTextOut(cell.richText) : plain;
        const h = cell.hyperlink;
        // A link with neither a target nor a location leads nowhere; exceljs, which reads only
        // links with a relationship, drops it as well (X4-10).
        if (h && use('hyperlinks') && (h.target !== null || (h.location !== null && h.location !== ''))) {
          // cell-xform.js L477–489: a linked cell becomes { text, hyperlink }, its text being the
          // value or the formula result (X4-10). xlvane also reads links stored only as a
          // location, and the tooltip, which exceljs drops. A target that is itself a place in
          // the workbook ("#Sheet2!A1", as some producers write it) is used as it is.
          const hyperlink = h.target !== null ? (h.location !== null && !h.target.startsWith('#') ? `${h.target}#${h.location}` : h.target) : `#${h.location ?? ''}`;
          const text = cell.richText ? richTextOut(cell.richText) : plain;
          value = h.tooltip ? { text, hyperlink, tooltip: h.tooltip } : { text, hyperlink };
        }
        target.value = value;
        target.style = fileStyle(cell.style, cell.numFmt, cell.styleIndex > 0);
        if (cell.row > lastRow) lastRow = cell.row;
      }
      // Row height, visibility and outline level, also of rows without cells (exceljs keeps every
      // row element, row-xform.js). The core gives row settings by number, so the rows up to the
      // sheet's dimension or its last cell are looked at (D-R01).
      const dim = /(\d+)$/.exec(s.dimension ?? '');
      lastRow = use('sheetData') ? Math.min(Math.max(lastRow, dim ? Number(dim[1]) : 0), 1_048_576) : 0;
      rowsLooked += lastRow;
      if (rowsLooked > rowBudget) {
        throw new XlvError('XLV_E_LIMIT_EXCEEDED', `The rows of worksheet "${s.name}" up to its dimension or last cell would bring the rows looked at for their settings past ${String(rowBudget)} for ${String(cellsRead)} cells in the file.`, {
          location: { sheet: s.name },
          hint: 'The sheets declare far more rows than they hold. Load the file with ignoreNodes: ["sheetData"] to skip cells and rows, or read it with xlvane\'s readWorkbook.',
          details: { limit: 'maxCells', value: rowBudget },
        });
      }
      for (let r = 1; r <= lastRow; r++) {
        const info = s.rowInfo(r);
        if (!info) continue;
        const row = ws.getRow(r);
        // The row's style, read only when the file marks it with customFormat="1" (F3; exceljs
        // reads row@s without that check, row-xform.js L80-82). Cells created later take it.
        if (info.style && use('sheetData')) row.style = lineStyle(info.style);
        // Excel shows rows up to 409.5 points (W-04, W-04a); a taller height in a file is read as
        // 409.5, so the workbook can be written again (D-R01).
        if (info.height !== null) row.height = Math.min(info.height, 409.5);
        if (info.hidden) row.hidden = true;
        if (info.outlineLevel) row.outlineLevel = info.outlineLevel;
      }
      // Notes are read for every cell, also one without a value; exceljs loses those (X4-14).
      for (const note of s.notes) ws.getCell(note.ref).note = noteOut(note);
      for (const m of use('mergeCells') ? s.merges : []) {
        if (!m.includes(':')) continue;
        spend(areaOf(m), 'A merged range', s.name);
        spendMerge(areaOf(m), m, s.name);
        try {
          // exceljs merges a file's ranges without styles (lib/doc/worksheet.js L895-901
          // `_parseMergeCells`), so each cell keeps its own style from the file.
          ws.mergeCellsWithoutStyle(m);
        } catch (error) {
          // exceljs's plain Error for overlapping merges becomes an XlvError on load.
          if (error instanceof XlvError || !(error instanceof Error) || error.message !== 'Cannot merge already merged cells') throw error;
          const other = mergeOverlapping(ws, m);
          throw new XlvError('XLV_E_OVERLAPPING_MERGE', `Cannot merge already merged cells: the merged range ${m} of worksheet "${s.name}" shares cells with ${other ?? 'an earlier merged range'}.`, {
            location: { sheet: s.name, cell: m },
            hint: 'The file has merged ranges that share cells, which exceljs does not accept. Load it with ignoreNodes: ["mergeCells"], or read it with xlvane\'s readWorkbook.',
            details: other === undefined ? {} : { other },
            cause: error,
          });
        }
      }
      // exceljs reads one view per sheetView, with every setting (sheet-view-xform.js; X4-11).
      if (use('sheetViews')) ws.views = [viewOut(s.view, s.freeze)];
      const tab = colorOut(s.tabColor ?? undefined);
      if (tab && use('sheetPr')) ws.properties['tabColor'] = tab;
      if (s.autoFilter && use('autoFilter')) ws.autoFilter = s.autoFilter;
      // Validations, conditional formats, protection and printing (X4-16 … X4-19; D-C06).
      if (use('dataValidations')) ws.dataValidations.model = dataValidationsOut(s.dataValidations, read.dateSystem === '1904', validations);
      if (use('conditionalFormatting')) ws.conditionalFormattings = conditionalFormatsOut(s.conditionalFormats);
      if (use('sheetProtection')) ws.sheetProtection = s.protection ? protectionOut(s.protection) : null;
      if (use('pageSetup')) ws.pageSetup = pageSetupOut(s.pageSetup, s.printArea, s.printTitles);
      if (use('headerFooter')) ws.headerFooter = headerFooterOut(s.headerFooter) as HeaderFooter;
      if (use('rowBreaks')) ws.rowBreaks = rowBreaksOut(s.pageBreaks.rows);
      // Tables with their ref and rows (X4-21); pictures and the background (X4-22).
      for (const def of use('tableParts') ? s.tables : []) {
        spend(areaOf(def.ref), 'A table', s.name);
        ws.tables[def.name] = tableOut(ws, def);
      }
      for (const img of use('drawing') ? s.images : []) {
        const id = img.image === null ? undefined : mediaIds.get(img.image);
        const media = id === undefined ? null : imageOut(ws, img, id);
        if (media) ws._media.push(media);
      }
      const background = s.backgroundImage === null ? undefined : mediaIds.get(s.backgroundImage);
      if (background !== undefined && use('picture')) ws._media.push({ type: 'background', imageId: background });
    }
  }
}

/** Merged cells `xlsx.load` creates at most for a file with fewer cells (one whole column; K2). */
const MERGED_CELLS_BUDGET = 1_048_576;

/**
 * @internal The chunks of an input stream, read until it ends. Reading stops with
 * XLV_E_LIMIT_EXCEEDED as soon as the chunks pass maxInputBytes, the size the core reads at most,
 * so an endless stream is not collected first (security review J, C8).
 */
export async function collectInput(stream: AsyncIterable<Uint8Array>, what: string): Promise<{ chunks: Uint8Array[]; size: number }> {
  const limit = DEFAULT_READ_LIMITS.maxInputBytes;
  const chunks: Uint8Array[] = [];
  let size = 0;
  for await (const chunk of stream) {
    chunks.push(chunk);
    size += chunk.length;
    if (size > limit) {
      throw new XlvError('XLV_E_LIMIT_EXCEEDED', `The stream given to ${what} holds more than ${String(limit)} bytes (maxInputBytes).`, {
        hint: 'Check that the stream is the XLSX file you expect. A larger file from a trusted source can be read with xlvane\'s readWorkbook and a higher limit, for example { limits: { maxInputBytes: 1073741824 } }.',
        details: { limit: 'maxInputBytes', value: limit },
      });
    }
  }
  return { chunks, size };
}

/**
 * Worksheet ids for a list of wanted ids: each whole number from 1 to 65534 (the sheet ids Office
 * uses, W2-20) that is not wanted by an earlier entry is kept; other entries (missing, invalid,
 * larger or repeated ids) get the lowest ids no entry wants, in order. The limit keeps the
 * `_sheets` array small: an id from an untrusted file such as 4294967295 is not an array index and
 * 20000000 makes every scan of the array slow (found by review, D-F03). `undefined` entries
 * (sheets that are not worksheets) get no id.
 */
function freeIds(wanted: readonly unknown[]): (number | undefined)[] {
  const taken = new Set<number>();
  const kept = wanted.map((w) => {
    if (typeof w !== 'number' || !Number.isInteger(w) || w < 1 || w > 65_534 || taken.has(w)) return null;
    taken.add(w);
    return w;
  });
  let next = 1;
  return kept.map((k, i) => {
    if (wanted[i] === undefined) return undefined;
    if (k !== null) return k;
    while (taken.has(next)) next++;
    taken.add(next);
    return next;
  });
}

/** The rows and columns of a range such as "A1:C3" or a cell, or `null` for anything else. */
function rangeOf(ref: string): { top: number; left: number; bottom: number; right: number } | null {
  const m = /^([A-Z]{1,3})(\d{1,7})(?::([A-Z]{1,3})(\d{1,7}))?$/.exec(ref.replace(/\$/g, ''));
  if (!m) return null;
  const col = (letters: string): number => {
    let n = 0;
    for (const ch of letters) n = n * 26 + ch.charCodeAt(0) - 64;
    return n;
  };
  const [, c1 = 'A', r1 = '1', c2 = c1, r2 = r1] = m;
  return { top: Math.min(Number(r1), Number(r2)), left: Math.min(col(c1), col(c2)), bottom: Math.max(Number(r1), Number(r2)), right: Math.max(col(c1), col(c2)) };
}

/** Cells in a range such as "A1:C3" (1 for a cell, 0 for anything else). */
function areaOf(ref: string): number {
  const r = rangeOf(ref);
  return r ? (r.bottom - r.top + 1) * (r.right - r.left + 1) : 0;
}

/** A merged range of the worksheet that shares cells with `ref`, as "A1:B2". */
function mergeOverlapping(ws: Worksheet, ref: string): string | undefined {
  const r = rangeOf(ref);
  if (!r) return undefined;
  for (const o of ws._merges.values()) {
    if (o.top <= r.bottom && r.top <= o.bottom && o.left <= r.right && r.left <= o.right) return `${columnLetter(o.left)}${String(o.top)}:${columnLetter(o.right)}${String(o.bottom)}`;
  }
  return undefined;
}
