// exceljs `Workbook` (lib/doc/workbook.js) and its `xlsx` reader/writer (lib/xlsx/xlsx.js), built
// on xlvane (docs/decisions.md D-503 … D-506).
import { XlvError, writeWorkbook, type Workbook as XlvWorkbook, type WorkbookImage } from 'xlvane';
import { colorOut, formulaOut, noteOut, readBytes, richTextOut, styleOut, toNewWorkbook, viewOut, type WorkbookMeta } from './convert';
import { CsvIO } from './csv';
import { DefinedNames, extractRanges } from './defined-names';
import { imageBytes, imageOut, mediaOut, tableOut, type ImageData } from './media';
import { Worksheet } from './model';
import { conditionalFormatsOut, dataValidationsOut, headerFooterOut, pageSetupOut, protectionOut, rowBreaksOut } from './sheet-rules';
import type { CellValue, HeaderFooter, WorksheetOptions } from './types';

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
 * import type { XlsxReadOptions } from '@xlvane/exceljs';
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
 * import type { XlsxWriteOptions } from '@xlvane/exceljs';
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
  async load(data: Uint8Array | ArrayBuffer | string, options?: XlsxReadOptions): Promise<Workbook> {
    const o = readOptions(options);
    this.workbook._fill(readBytes(o.base64 ? fromBase64(data) : toBytes(data)), o);
    return this.workbook;
  }

  /** Reads an XLSX file (Node.js). */
  async readFile(path: string, options?: XlsxReadOptions): Promise<Workbook> {
    const o = readOptions(options);
    const { readWorkbookFile } = await import('xlvane/node');
    this.workbook._fill(await readWorkbookFile(path), o);
    return this.workbook;
  }

  /** Reads XLSX data from a readable stream (Node.js stream or any async iterable of bytes). */
  async read(stream: AsyncIterable<Uint8Array>, options?: XlsxReadOptions): Promise<Workbook> {
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
    return this.load(all, { ...readOptions(options), base64: false });
  }

  /** Writes the workbook and returns the file contents (a Buffer in Node.js). */
  async writeBuffer(options?: XlsxWriteOptions): Promise<Uint8Array> {
    return asBuffer(this.bytes(options));
  }

  /** Writes the workbook to a file, replacing an existing file as exceljs does (Node.js). */
  async writeFile(path: string, options?: XlsxWriteOptions): Promise<void> {
    const { writeWorkbookFile } = await import('xlvane/node');
    await writeWorkbookFile(this.book(options), path, { overwrite: true, ...writeOptions(options) });
  }

  /** Writes the workbook to a writable stream and ends it. */
  async write(stream: WritableLike, options?: XlsxWriteOptions): Promise<void> {
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

  private book(options: XlsxWriteOptions | undefined) {
    const useStyles = options?.useStyles !== false;
    return toNewWorkbook(this.workbook.worksheets, this.workbook._date1904, { ...this.workbook._meta(), useStyles });
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
  private order = 0;
  private readonly orders = new Map<Worksheet, number>();

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

  /** @internal */
  get _date1904(): boolean {
    return this.properties.date1904 === true;
  }

  /** @internal Names of the other sheets, for the duplicate check of Worksheet.name. */
  readonly _names = (exclude: Worksheet): string[] => this.worksheets.filter((w) => w !== exclude).map((w) => w.name);

  /** The id the next worksheet gets (lib/doc/workbook.js `nextId`). */
  get nextId(): number {
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
    for (const m of Array.isArray(value['worksheets']) ? (value['worksheets'] as Record<string, unknown>[]) : []) {
      const ws = this.addWorksheet(String(m['name'] ?? ''), { state: (m['state'] as 'visible') ?? 'visible' });
      ws.model = m;
    }
    this._definedNames.model = Array.isArray(value['definedNames']) ? (value['definedNames'] as { name: string; ranges: string[] }[]) : [];
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

  /** @internal Images of the workbook, as exceljs `workbook.media`. */
  _media: ImageData[] = [];

  /**
   * Adds an image (PNG, JPEG or GIF, recognised from its bytes) and returns its id for
   * `worksheet.addImage` and `addBackgroundImage`, as exceljs `workbook.addImage`. Give `buffer`
   * or `base64`; `filename` throws `XLV_E_NOT_SUPPORTED` (read the file and pass `buffer`).
   */
  addImage(image: ImageData): number {
    imageBytes(image);
    this._media.push({ ...image, type: 'image' });
    return this._media.length - 1;
  }

  /** The image with this id, as exceljs `workbook.getImage`. */
  getImage(id: number): ImageData | undefined {
    return this._media[id];
  }

  /**
   * @internal Replaces the sheets with those of a workbook read by xlvane. `options.ignoreNodes`
   * skips what exceljs skips for those worksheet elements, and `maxRows`/`maxCols` refuse
   * bigger worksheets (worksheet-xform.js L95, L372; list-xform.js L75; row-xform.js L119; X4-25).
   */
  _fill(read: XlvWorkbook, options: XlsxReadOptions = {}): void {
    const ignore = new Set(options.ignoreNodes ?? []);
    const use = (node: string): boolean => !ignore.has(node);
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
    // xlsx.js L348–349: company and manager come from docProps/app.xml (X4-13).
    this.company = p.company;
    this.manager = p.manager;
    // Print areas and titles become pageSetup in exceljs (workbook-xform.js L187; read from
    // Sheet.printArea / printTitles below); other built-in names are written by xlvane itself
    // (autofilter). The scope (localSheetId) is dropped, as in exceljs.
    // Images: the media list as exceljs reads it, one entry per image part (X4-22).
    this._media = mediaOut(read.images);
    // bookViews and calcPr as exceljs reads them (workbook-view-xform.js L19–43; calcPr gives {}; X4-28).
    this.views = read.views.map((v) => {
      const view: Record<string, unknown> = { x: v.x ?? 0, y: v.y ?? 0, width: v.width ?? 25000, height: v.height ?? 10000, visibility: v.visibility };
      if (v.activeTab !== null) view['activeTab'] = v.activeTab;
      if (v.firstSheet !== null) view['firstSheet'] = v.firstSheet;
      return view;
    });
    this.calcProperties = read.calculation?.fullCalcOnLoad ? { fullCalcOnLoad: true } : {};
    const mediaIds = new Map<WorkbookImage, number>(read.images.map((img, i) => [img, i]));
    this._definedNames.model = read.definedNames.filter((d) => !d.name.startsWith('_xlnm.')).map((d) => ({ name: d.name, ranges: extractRanges(d.refersTo) }));
    for (const s of read.sheets) {
      if (s.kind !== 'worksheet') continue;
      const ws = this.addWorksheet(s.name, { state: s.state });
      for (const c of use('cols') ? s.columns : []) {
        const last = Math.min(c.max, 16_384);
        for (let n = c.min; n <= last; n++) {
          const column = ws.getColumn(n);
          if (c.width !== null) column.width = c.width;
          if (c.hidden) column.hidden = true;
          if (c.outlineLevel) column.outlineLevel = c.outlineLevel;
          if (c.style) column.style = styleOut(c.style, typeof c.style.numFmt === 'string' ? c.style.numFmt : null);
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
      for (const cell of use('sheetData') ? s.cells() : []) {
        const target = ws.getCell(cell.row, cell.col);
        let value: CellValue;
        const plain = cell.type === 'error' ? { error: cell.value as '#N/A' } : cell.value;
        if (cell.formula && cell.formula.text !== null) value = formulaOut(cell.formula, plain, cell.ref, masters);
        else value = cell.richText ? richTextOut(cell.richText) : plain;
        if (cell.hyperlink && use('hyperlinks')) {
          // cell-xform.js L477–489: a linked cell becomes { text, hyperlink }, its text being the
          // value or the formula result (X4-10). xlvane also reads links stored only as a
          // location, and the tooltip, which exceljs drops.
          const h = cell.hyperlink;
          const hyperlink = h.target !== null ? (h.location !== null ? `${h.target}#${h.location}` : h.target) : `#${h.location ?? ''}`;
          const text = cell.richText ? richTextOut(cell.richText) : plain;
          value = h.tooltip ? { text, hyperlink, tooltip: h.tooltip } : { text, hyperlink };
        }
        target.value = value;
        target.style = styleOut(cell.style, cell.numFmt);
        const info = s.rowInfo(cell.row);
        if (info) {
          const row = ws.getRow(cell.row);
          if (info.height !== null) row.height = info.height;
          if (info.hidden) row.hidden = true;
          if (info.outlineLevel) row.outlineLevel = info.outlineLevel;
        }
      }
      // Notes are read for every cell, also one without a value; exceljs loses those (X4-14).
      for (const note of s.notes) ws.getCell(note.ref).note = noteOut(note);
      for (const m of use('mergeCells') ? s.merges : []) if (m.includes(':')) ws.mergeCells(m);
      // exceljs reads one view per sheetView, with every setting (sheet-view-xform.js; X4-11).
      if (use('sheetViews')) ws.views = [viewOut(s.view, s.freeze)];
      const tab = colorOut(s.tabColor ?? undefined);
      if (tab && use('sheetPr')) ws.properties['tabColor'] = tab;
      if (s.autoFilter && use('autoFilter')) ws.autoFilter = s.autoFilter;
      // Validations, conditional formats, protection and printing (X4-16 … X4-19; D-C06).
      if (use('dataValidations')) ws.dataValidations.model = dataValidationsOut(s.dataValidations, read.dateSystem === '1904');
      if (use('conditionalFormatting')) ws.conditionalFormattings = conditionalFormatsOut(s.conditionalFormats);
      if (use('sheetProtection')) ws.sheetProtection = s.protection ? protectionOut(s.protection) : null;
      if (use('pageSetup')) ws.pageSetup = pageSetupOut(s.pageSetup, s.printArea, s.printTitles);
      if (use('headerFooter')) ws.headerFooter = headerFooterOut(s.headerFooter) as HeaderFooter;
      if (use('rowBreaks')) ws.rowBreaks = rowBreaksOut(s.pageBreaks.rows);
      // Tables with their ref and rows (X4-21); pictures and the background (X4-22).
      for (const def of use('tableParts') ? s.tables : []) ws.tables[def.name] = tableOut(ws, def);
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
