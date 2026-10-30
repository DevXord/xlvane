// The exceljs-compatible document model: Workbook, Worksheet, Row, Column and Cell, following
// the behaviour of exceljs 4.4.0 (lib/doc/*.js; docs/decisions.md D-501 … D-506). Reading and
// writing go through xlvane (convert.ts).
import { XlvError } from 'xlvane';
import type { DefinedNames } from './defined-names';
import { adjustFormulaText, checkInteger, copyFormula, copyPlain, formulaNotAdjustable, invalidArgument, moveSpan, movePosition, type SpliceChange } from './splice';
import type { Alignment, Borders, CellValue, ColumnDefn, Comment, Fill, Font, Style, WorksheetOptions, WorksheetView } from './types';

/**
 * Value types, as in exceljs `lib/doc/enums.js` (ValueType).
 *
 * @example
 * ```ts
 * import ExcelJS from '@xlvane/exceljs';
 *
 * const cell = new ExcelJS.Workbook().addWorksheet('S').getCell('A1');
 * cell.value = 42;
 * console.log(cell.type === ExcelJS.ValueType.Number); // true
 * ```
 */
export const ValueType: {
  readonly Null: 0;
  readonly Merge: 1;
  readonly Number: 2;
  readonly String: 3;
  readonly Date: 4;
  readonly Hyperlink: 5;
  readonly Formula: 6;
  readonly SharedString: 7;
  readonly RichText: 8;
  readonly Boolean: 9;
  readonly Error: 10;
} = Object.freeze({
  Null: 0,
  Merge: 1,
  Number: 2,
  String: 3,
  Date: 4,
  Hyperlink: 5,
  Formula: 6,
  SharedString: 7,
  RichText: 8,
  Boolean: 9,
  Error: 10,
} as const);

/**
 * Formula types, as in exceljs `lib/doc/enums.js` (FormulaType).
 *
 * @example
 * ```ts
 * import { FormulaType } from '@xlvane/exceljs';
 *
 * console.log(FormulaType.Master); // 1
 * ```
 */
export const FormulaType: { readonly None: 0; readonly Master: 1; readonly Shared: 2 } = Object.freeze({ None: 0, Master: 1, Shared: 2 } as const);

/**
 * Error values, as in exceljs `lib/doc/enums.js` (ErrorValue).
 *
 * @example
 * ```ts
 * import ExcelJS from '@xlvane/exceljs';
 *
 * const cell = new ExcelJS.Workbook().addWorksheet('S').getCell('A1');
 * cell.value = { error: ExcelJS.ErrorValue.NotApplicable };
 * ```
 */
export const ErrorValue: {
  readonly NotApplicable: '#N/A';
  readonly Ref: '#REF!';
  readonly Name: '#NAME?';
  readonly DivZero: '#DIV/0!';
  readonly Null: '#NULL!';
  readonly Value: '#VALUE!';
  readonly Num: '#NUM!';
} = Object.freeze({
  NotApplicable: '#N/A',
  Ref: '#REF!',
  Name: '#NAME?',
  DivZero: '#DIV/0!',
  Null: '#NULL!',
  Value: '#VALUE!',
  Num: '#NUM!',
} as const);

const COMPAT_DOCS = 'https://github.com/DevXord/xlvane/blob/main/docs/exceljs-compat.md';

/** Throws `XLV_E_NOT_SUPPORTED` for exceljs features that @xlvane/exceljs does not implement (D-502). */
export function notSupported(feature: string): never {
  throw new XlvError('XLV_E_NOT_SUPPORTED', `${feature} is not supported by @xlvane/exceljs.`, {
    hint: `See ${COMPAT_DOCS} for what is supported. Remove this call or keep using exceljs for this file.`,
    details: { feature },
  });
}

/** Column letters for a 1-based number (A … XFD). */
export function columnLetter(n: number): string {
  let s = '';
  let x = n;
  while (x > 0) {
    const r = (x - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    x = Math.floor((x - 1) / 26);
  }
  return s;
}

/** 1-based column number for letters, or NaN. */
export function columnNumber(letters: string): number {
  if (!/^[A-Za-z]{1,3}$/.test(letters)) return Number.NaN;
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

/** Parses "B7" (also "$B$7") into row and column. */
export function decodeAddress(address: string): { row: number; col: number } {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(address.trim());
  const col = m ? columnNumber(m[1] ?? '') : Number.NaN;
  const row = m ? Number(m[2]) : Number.NaN;
  if (!m || !(row >= 1 && row <= 1_048_576) || !(col >= 1 && col <= 16_384)) {
    throw new XlvError('XLV_E_INVALID_CELL_REF', `"${address.slice(0, 40)}" is not a valid cell address.`, {
      hint: 'Use an A1-style address such as "B7".',
      details: { ref: address.slice(0, 40) },
    });
  }
  return { row, col };
}

function shallowStyle(style: Style | undefined): Style {
  return style ? { ...style } : {};
}

function isRecord(v: unknown): v is Record<string, unknown> {
  return typeof v === 'object' && v !== null && !(v instanceof Date) && !Array.isArray(v);
}

/** The ValueType of a value, as exceljs `Value.getType` (lib/doc/cell.js L1062). */
export function typeOf(value: unknown): number {
  if (value === null || value === undefined) return ValueType.Null;
  if (typeof value === 'string' || value instanceof String) return ValueType.String;
  if (typeof value === 'number') return ValueType.Number;
  if (typeof value === 'boolean') return ValueType.Boolean;
  if (value instanceof Date) return ValueType.Date;
  if (isRecord(value)) {
    if (value['text'] && value['hyperlink']) return ValueType.Hyperlink;
    if (value['formula'] || value['sharedFormula']) return ValueType.Formula;
    if (value['richText']) return ValueType.RichText;
    if (value['sharedString']) return ValueType.SharedString;
    if (value['error']) return ValueType.Error;
  }
  return -1;
}

/** Checks a value for writing support (D-503); unsupported kinds throw XLV_E_NOT_SUPPORTED. */
function checkValue(value: unknown): CellValue {
  const type = typeOf(value);
  if (type === ValueType.RichText) {
    const runs = (value as { richText: unknown }).richText;
    if (!Array.isArray(runs)) return notSupported('A { richText } cell value whose richText is not an array');
  }
  if (type === ValueType.SharedString) return notSupported('A { sharedString } cell value');
  if (type === -1) return notSupported(`A cell value of this kind (${Array.isArray(value) ? 'array' : typeof value})`);
  if (value instanceof String) return value.valueOf();
  return value as CellValue;
}

/**
 * A cell, as exceljs `Cell` (lib/doc/cell.js).
 *
 * @example
 * ```ts
 * import ExcelJS from '@xlvane/exceljs';
 *
 * const sheet = new ExcelJS.Workbook().addWorksheet('Data');
 * const cell = sheet.getCell('B2');
 * cell.value = 42;
 * cell.font = { bold: true };
 * console.log(cell.address, cell.type === ExcelJS.ValueType.Number, cell.text); // B2 true 42
 * ```
 */
export class Cell {
  /** The cell's style; `font`, `fill`, … are shortcuts to its properties. */
  style: Style;
  private _value: CellValue = null;
  private _note: string | Comment | undefined = undefined;
  /** @internal Master cell when this cell is part of a merged range but not its top-left cell. */
  _master: Cell | null = null;
  /** @internal Number of merged ranges this cell is the master of. */
  _mergeCount = 0;

  /** @internal Use `worksheet.getCell()`. */
  constructor(
    private readonly _row: Row,
    /** @internal The column number; changed when cells are spliced. */
    public _col: number,
  ) {
    const column = _row.worksheet.findColumn(_col);
    // Row style wins over column style, part by part (lib/doc/cell.js _mergeStyle).
    this.style = { ...column?.style, ...withoutEmpty(_row.style) };
  }

  get worksheet(): Worksheet {
    return this._row.worksheet;
  }

  get row(): number {
    return this._row.number;
  }

  get col(): number {
    return this._col;
  }

  get address(): string {
    return `${columnLetter(this._col)}${String(this._row.number)}`;
  }

  get $col$row(): string {
    return `$${columnLetter(this._col)}$${String(this._row.number)}`;
  }

  get value(): CellValue {
    return this._master ? this._master.value : this._value;
  }

  set value(v: CellValue) {
    if (this._master) {
      this._master.value = v;
      return;
    }
    this._value = checkValue(v);
  }

  /** The ValueType (Merge for the non-master cells of a merged range). */
  get type(): number {
    if (this._master) return ValueType.Merge;
    return typeOf(this._value);
  }

  /** The type of the value shown: for a formula, the type of its result. */
  get effectiveType(): number {
    const v = this.value;
    if (isRecord(v) && ('formula' in v || 'sharedFormula' in v)) return typeOf(v['result']);
    return this._master ? this._master.effectiveType : this.type;
  }

  /** The value as text, as exceljs `cell.text`. */
  get text(): string {
    const v = this.value;
    if (v === null || v === undefined) return '';
    if (isRecord(v)) {
      if ('formula' in v || 'sharedFormula' in v) {
        // exceljs FormulaValue.toString: `result ? result.toString() : ''` (lib/doc/cell.js L877).
        const r = v['result'];
        if (!r) return '';
        return isRecord(r) ? String(r['error'] ?? '') : String(r);
      }
      if ('error' in v) return String(v['error']);
      // HyperlinkValue.toString returns the link text (lib/doc/cell.js L664; X4-10).
      if (typeOf(v) === ValueType.Hyperlink) {
        const t = v['text'];
        if (isRecord(t) && Array.isArray(t['richText'])) return (t['richText'] as { text?: unknown }[]).map((r) => String(r.text ?? '')).join('');
        return String(t);
      }
      // RichTextValue.text joins the texts of the runs (lib/doc/cell.js L511).
      if (Array.isArray(v['richText'])) return (v['richText'] as { text?: unknown }[]).map((r) => String(r.text ?? '')).join('');
    }
    return String(v);
  }

  toString(): string {
    return this.text;
  }

  get isMerged(): boolean {
    return this._mergeCount > 0 || this._master !== null;
  }

  /** The top-left cell of the merged range this cell belongs to (the cell itself otherwise). */
  get master(): Cell {
    return this._master ?? this;
  }

  get formula(): string | undefined {
    const v = this.value;
    return isRecord(v) && typeof v['formula'] === 'string' ? v['formula'] : undefined;
  }

  get result(): unknown {
    const v = this.value;
    return isRecord(v) ? v['result'] : undefined;
  }

  get numFmt(): string | undefined {
    return this.style.numFmt;
  }
  set numFmt(v: string | undefined) {
    this.style.numFmt = v;
  }
  get font(): Partial<Font> | undefined {
    return this.style.font;
  }
  set font(v: Partial<Font> | undefined) {
    this.style.font = v;
  }
  get fill(): Fill | undefined {
    return this.style.fill;
  }
  set fill(v: Fill | undefined) {
    this.style.fill = v;
  }
  get border(): Partial<Borders> | undefined {
    return this.style.border;
  }
  set border(v: Partial<Borders> | undefined) {
    this.style.border = v;
  }
  get alignment(): Partial<Alignment> | undefined {
    return this.style.alignment;
  }
  set alignment(v: Partial<Alignment> | undefined) {
    this.style.alignment = v;
  }
  get protection(): unknown {
    return this.style.protection;
  }
  set protection(v: unknown) {
    if (v !== undefined && v !== null) notSupported('Cell protection');
  }
  /**
   * The cell's note: a string or `{ texts: [{ text, font? }] }`, as in exceljs (lib/doc/cell.js
   * L224–230). The note belongs to this cell, also inside a merged range. `undefined` removes it.
   */
  get note(): string | Comment | undefined {
    return this._note;
  }
  set note(v: string | Comment | undefined) {
    if (v === undefined || v === null) {
      this._note = undefined;
      return;
    }
    if (typeof v !== 'string' && (typeof v !== 'object' || Array.isArray(v) || (v.texts !== undefined && !Array.isArray(v.texts)))) {
      notSupported(`A note of cell ${this.address} that is not a string or { texts: [...] }`);
    }
    this._note = v;
  }
  get dataValidation(): undefined {
    return undefined;
  }
  set dataValidation(_v: unknown) {
    notSupported('Data validation');
  }
  /** The link target of a `{ text, hyperlink }` value, as exceljs `cell.hyperlink` (lib/doc/cell.js L201). */
  get hyperlink(): string | undefined {
    const v = this.value;
    return isRecord(v) && typeOf(v) === ValueType.Hyperlink ? String(v['hyperlink']) : undefined;
  }

  /** Whether the value is a `{ text, hyperlink }` value (lib/doc/cell.js L197). */
  get isHyperlink(): boolean {
    return this.type === ValueType.Hyperlink;
  }

  /** The cell's location for defined names: `{ sheetName, address, row, col }` (lib/doc/cell.js L270). */
  get fullAddress(): { sheetName: string; address: string; row: number; col: number } {
    return { sheetName: this.worksheet.name, address: this.address, row: this.row, col: this.col };
  }

  private get definedNames(): DefinedNames {
    const names = this.worksheet.workbook.definedNames;
    if (!names) return notSupported('Defined names of this workbook');
    return names;
  }

  /** The first defined name that includes this cell. */
  get name(): string | undefined {
    return this.names[0];
  }
  set name(value: string) {
    this.names = [value];
  }

  /** The defined names that include this cell. */
  get names(): string[] {
    return this.definedNames.getNamesEx(this.fullAddress);
  }
  set names(value: string[]) {
    const names = this.definedNames;
    names.removeAllNames(this.fullAddress);
    for (const name of value) names.addEx(this.fullAddress, name);
  }

  addName(name: string): void {
    this.definedNames.addEx(this.fullAddress, name);
  }

  removeName(name: string): void {
    this.definedNames.removeEx(this.fullAddress, name);
  }

  removeAllNames(): void {
    this.definedNames.removeAllNames(this.fullAddress);
  }

  /** @internal */
  _merge(master: Cell): void {
    this._master = master;
    this._value = null;
  }

  /** @internal */
  _unmerge(): void {
    this._master = null;
  }
}

function withoutEmpty(style: Style): Style {
  const out: Style = {};
  for (const [k, v] of Object.entries(style)) if (v !== undefined) (out as Record<string, unknown>)[k] = v;
  return out;
}

/**
 * A row, as exceljs `Row` (lib/doc/row.js).
 *
 * @example
 * ```ts
 * import ExcelJS from '@xlvane/exceljs';
 *
 * const sheet = new ExcelJS.Workbook().addWorksheet('Data');
 * const row = sheet.addRow(['Tea', 3.5]);
 * row.font = { italic: true };
 * console.log(row.number, row.values); // 1 [ <1 empty item>, 'Tea', 3.5 ]
 * ```
 */
export class Row {
  /** Style applied to cells created in this row, and set on existing cells by the setters. */
  style: Style = {};
  height: number | undefined;
  private _hidden = false;
  /** @internal */
  _cells: (Cell | undefined)[] = [];

  /** @internal The row number; changed when rows are spliced. */
  _number: number;

  /** @internal Use `worksheet.getRow()`. */
  constructor(
    readonly worksheet: Worksheet,
    number: number,
  ) {
    this._number = number;
  }

  /** The row number. It follows the row when rows are inserted or removed above it. */
  get number(): number {
    return this._number;
  }

  get hidden(): boolean {
    return this._hidden;
  }
  set hidden(v: boolean) {
    this._hidden = !!v;
  }

  get outlineLevel(): number {
    return 0;
  }
  set outlineLevel(v: number) {
    if (v) notSupported('Row outline levels');
  }

  findCell(col: number): Cell | undefined {
    return this._cells[col - 1];
  }

  /** The cell at a column number, letters or column key; created if missing. */
  getCell(col: number | string): Cell {
    let n: number;
    if (typeof col === 'string') {
      const keyed = this.worksheet.getColumnKey(col);
      n = keyed ? keyed.number : columnNumber(col);
    } else n = col;
    if (!(Number.isInteger(n) && n >= 1 && n <= 16_384)) {
      throw new XlvError('XLV_E_INVALID_CELL_REF', `Column ${String(col).slice(0, 20)} is not valid.`, {
        hint: 'Use a column number from 1 to 16384, column letters such as "C", or a column key.',
        details: { ref: String(col).slice(0, 20) },
      });
    }
    let cell = this._cells[n - 1];
    if (!cell) {
      cell = new Cell(this, n);
      this._cells[n - 1] = cell;
    }
    return cell;
  }

  /** Visits cells with a value (all cells up to the last one with `includeEmpty`). */
  eachCell(options: { includeEmpty?: boolean } | ((cell: Cell, col: number) => void), iteratee?: (cell: Cell, col: number) => void): void {
    const cb = typeof options === 'function' ? options : iteratee;
    const includeEmpty = typeof options === 'object' && options !== null && options.includeEmpty === true;
    if (!cb) return;
    if (includeEmpty) {
      for (let i = 1; i <= this._cells.length; i++) cb(this.getCell(i), i);
      return;
    }
    this._cells.forEach((cell, i) => {
      if (cell && cell.type !== ValueType.Null) cb(cell, i + 1);
    });
  }

  /** Sparse array of values, index 0 unused (lib/doc/row.js `values`). */
  get values(): CellValue[] {
    const out: CellValue[] = [];
    this._cells.forEach((cell) => {
      if (cell && cell.type !== ValueType.Null) out[cell.col] = cell.value;
    });
    return out;
  }

  /**
   * Replaces the row's values: a contiguous array starts at column A, a sparse array (index 0
   * unset) uses its indexes as column numbers, an object is matched to column keys.
   */
  set values(value: CellValue[] | Record<string, CellValue> | null | undefined) {
    this._cells = [];
    if (!value) return;
    if (Array.isArray(value)) {
      const offset = Object.prototype.hasOwnProperty.call(value, '0') ? 1 : 0;
      value.forEach((item, index) => {
        if (item !== undefined) this.getCell(index + offset).value = item;
      });
      return;
    }
    this.worksheet.eachColumnKey((column, key) => {
      const v = Object.prototype.hasOwnProperty.call(value, key) ? value[key] : undefined;
      if (v !== undefined) this.getCell(column.number).value = v;
    });
  }

  get hasValues(): boolean {
    return this._cells.some((c) => c && c.type !== ValueType.Null);
  }

  get cellCount(): number {
    return this._cells.length;
  }

  get actualCellCount(): number {
    let n = 0;
    this.eachCell(() => n++);
    return n;
  }

  get dimensions(): { min: number; max: number } | null {
    let min = 0;
    let max = 0;
    for (const c of this._cells) {
      if (c && c.type !== ValueType.Null) {
        if (!min || c.col < min) min = c.col;
        if (c.col > max) max = c.col;
      }
    }
    return min > 0 ? { min, max } : null;
  }

  private applyStyle<K extends keyof Style>(name: K, value: Style[K]): void {
    this.style[name] = value;
    for (const c of this._cells) if (c) c.style[name] = value;
  }

  get numFmt(): string | undefined {
    return this.style.numFmt;
  }
  set numFmt(v: string | undefined) {
    this.applyStyle('numFmt', v);
  }
  get font(): Partial<Font> | undefined {
    return this.style.font;
  }
  set font(v: Partial<Font> | undefined) {
    this.applyStyle('font', v);
  }
  get fill(): Fill | undefined {
    return this.style.fill;
  }
  set fill(v: Fill | undefined) {
    this.applyStyle('fill', v);
  }
  get border(): Partial<Borders> | undefined {
    return this.style.border;
  }
  set border(v: Partial<Borders> | undefined) {
    this.applyStyle('border', v);
  }
  get alignment(): Partial<Alignment> | undefined {
    return this.style.alignment;
  }
  set alignment(v: Partial<Alignment> | undefined) {
    this.applyStyle('alignment', v);
  }

  /** Streaming writers write committed rows; in a normal workbook this does nothing (as in exceljs). */
  commit(): void {
    this.worksheet._commitRow(this);
  }

  /**
   * Removes `count` cells from column `start` and inserts `inserts` there; the cells to the right
   * move with their values, styles and notes (as exceljs `row.splice`, lib/doc/row.js L77). A row
   * crossing a merged range at or right of `start` is refused, because the range cannot follow
   * cells of one row; use `worksheet.spliceColumns` or unmerge first. Formulas are not rewritten,
   * as in exceljs.
   */
  splice(start: number, count: number, ...inserts: CellValue[]): void {
    checkInteger('row.splice()', 'start', start, 1, 16_384);
    checkInteger('row.splice()', 'count', count, 0, 16_384);
    for (const m of this.worksheet._merges.values()) {
      if (m.top <= this._number && this._number <= m.bottom && m.right >= start) {
        notSupported(`row.splice() of row ${String(this._number)}, which crosses the merged range ${columnLetter(m.left)}${String(m.top)}:${columnLetter(m.right)}${String(m.bottom)}`);
      }
    }
    this._spliceCells({ sheet: this.worksheet.name, axis: 'col', at: start, count, inserts: inserts.length }, 'row.splice()');
    inserts.forEach((value, i) => {
      const cell = this.getCell(start + i);
      cell.value = value;
      cell.style = {};
    });
  }

  /** @internal Moves the cell objects of this row for a column splice. */
  _spliceCells(change: SpliceChange, method: string): void {
    const last = this._cells.length;
    if (last >= change.at + change.count && last - change.count + change.inserts > 16_384) {
      invalidArgument(method, 'inserts', change.inserts, 'The cells would move beyond column XFD (16384). Remove cells at the end of the row first.');
    }
    const next: (Cell | undefined)[] = [];
    this._cells.forEach((cell, i) => {
      if (!cell) return;
      const col = movePosition(i + 1, change);
      if (col === null) return;
      cell._col = col;
      next[col - 1] = cell;
    });
    this._cells = next;
  }

  addPageBreak(): void {
    notSupported('Page breaks');
  }
}

const DEFAULT_COLUMN_WIDTH = 9;

/**
 * A column, as exceljs `Column` (lib/doc/column.js).
 *
 * @example
 * ```ts
 * import ExcelJS from '@xlvane/exceljs';
 *
 * const sheet = new ExcelJS.Workbook().addWorksheet('Data');
 * sheet.columns = [{ header: 'Name', key: 'name', width: 20 }, { header: 'Total', key: 'total' }];
 * sheet.getColumn('total').numFmt = '#,##0.00';
 * console.log(sheet.getColumn('name').letter); // A
 * ```
 */
export class Column {
  width: number | undefined;
  style: Style = {};
  private _hidden = false;
  private _key: string | undefined;
  private _header: string | string[] | undefined;

  /** @internal The column number; changed when columns are spliced. */
  _number: number;

  /** @internal Use `worksheet.getColumn()`. */
  constructor(
    readonly worksheet: Worksheet,
    number: number,
  ) {
    this._number = number;
  }

  /** The column number. It follows the column when columns are inserted or removed before it. */
  get number(): number {
    return this._number;
  }

  get letter(): string {
    return columnLetter(this.number);
  }

  get key(): string | undefined {
    return this._key;
  }
  set key(value: string | undefined) {
    if (this._key !== undefined && this.worksheet.getColumnKey(this._key) === this) this.worksheet.deleteColumnKey(this._key);
    this._key = value;
    if (value) this.worksheet.setColumnKey(value, this);
  }

  get header(): string | string[] | undefined {
    return this._header;
  }
  set header(value: string | string[] | undefined) {
    this._header = value;
    if (value === undefined) return;
    const headers = Array.isArray(value) ? value : [value];
    headers.forEach((text, i) => {
      this.worksheet.getCell(i + 1, this.number).value = text;
    });
  }

  get headers(): (string | undefined)[] {
    return Array.isArray(this._header) ? this._header : [this._header];
  }

  get hidden(): boolean {
    return this._hidden;
  }
  set hidden(v: boolean) {
    this._hidden = !!v;
  }

  get isCustomWidth(): boolean {
    return this.width !== undefined && this.width !== DEFAULT_COLUMN_WIDTH;
  }

  get defn(): ColumnDefn {
    return { header: this._header, key: this._key, width: this.width, style: this.style, hidden: this._hidden };
  }

  set defn(value: ColumnDefn | undefined) {
    if (!value) {
      this._header = undefined;
      this.key = undefined;
      this.width = undefined;
      this.style = {};
      return;
    }
    this.key = value.key;
    this.width = value.width !== undefined ? value.width : DEFAULT_COLUMN_WIDTH;
    this.style = value.style ? { ...value.style } : {};
    this.header = value.header;
    this._hidden = !!value.hidden;
  }

  /** Cell values of this column by row number (index 0 unused). */
  get values(): CellValue[] {
    const out: CellValue[] = [];
    this.eachCell((cell, r) => {
      out[r] = cell.value;
    });
    return out;
  }

  set values(values: CellValue[]) {
    const offset = Object.prototype.hasOwnProperty.call(values, '0') ? 1 : 0;
    values.forEach((v, i) => {
      if (v !== undefined) this.worksheet.getCell(i + offset, this.number).value = v;
    });
  }

  eachCell(options: { includeEmpty?: boolean } | ((cell: Cell, row: number) => void), iteratee?: (cell: Cell, row: number) => void): void {
    const cb = typeof options === 'function' ? options : iteratee;
    const includeEmpty = typeof options === 'object' && options !== null && options.includeEmpty === true;
    if (!cb) return;
    this.worksheet.eachRow({ includeEmpty }, (row, r) => {
      const cell = includeEmpty ? row.getCell(this.number) : row.findCell(this.number);
      if (cell && (includeEmpty || cell.type !== ValueType.Null)) cb(cell, r);
    });
  }

  private applyStyle<K extends keyof Style>(name: K, value: Style[K]): void {
    this.style[name] = value;
    this.eachCell((cell) => {
      cell.style[name] = value;
    });
  }

  get numFmt(): string | undefined {
    return this.style.numFmt;
  }
  set numFmt(v: string | undefined) {
    this.applyStyle('numFmt', v);
  }
  get font(): Partial<Font> | undefined {
    return this.style.font;
  }
  set font(v: Partial<Font> | undefined) {
    this.applyStyle('font', v);
  }
  get fill(): Fill | undefined {
    return this.style.fill;
  }
  set fill(v: Fill | undefined) {
    this.applyStyle('fill', v);
  }
  get border(): Partial<Borders> | undefined {
    return this.style.border;
  }
  set border(v: Partial<Borders> | undefined) {
    this.applyStyle('border', v);
  }
  get alignment(): Partial<Alignment> | undefined {
    return this.style.alignment;
  }
  set alignment(v: Partial<Alignment> | undefined) {
    this.applyStyle('alignment', v);
  }
}

/** Range bounds parsed from "A1:C3", "A1" or (top, left, bottom, right). */
export function parseRangeArgs(args: readonly unknown[]): { top: number; left: number; bottom: number; right: number } {
  if (args.length >= 4 && args.slice(0, 4).every((a) => typeof a === 'number')) {
    const [t, l, b, r] = args as number[];
    return { top: Math.min(t ?? 0, b ?? 0), left: Math.min(l ?? 0, r ?? 0), bottom: Math.max(t ?? 0, b ?? 0), right: Math.max(l ?? 0, r ?? 0) };
  }
  const text = args.length === 2 && typeof args[0] === 'string' && typeof args[1] === 'string' ? `${args[0]}:${args[1]}` : String(args[0] ?? '');
  const [a, b = a] = text.split(':');
  const s = decodeAddress(a ?? '');
  const e = decodeAddress(b ?? '');
  return { top: Math.min(s.row, e.row), left: Math.min(s.col, e.col), bottom: Math.max(s.row, e.row), right: Math.max(s.col, e.col) };
}

/**
 * A worksheet, as exceljs `Worksheet` (lib/doc/worksheet.js).
 *
 * @example
 * ```ts
 * import ExcelJS from '@xlvane/exceljs';
 *
 * const sheet = new ExcelJS.Workbook().addWorksheet('Sales', { views: [{ state: 'frozen', ySplit: 1 }] });
 * sheet.columns = [{ header: 'Region', key: 'region', width: 15 }, { header: 'Revenue', key: 'revenue' }];
 * sheet.addRow({ region: 'North', revenue: 1250 });
 * sheet.mergeCells('D1:E1');
 * sheet.eachRow((row, n) => console.log(n, row.values));
 * ```
 */
export class Worksheet {
  readonly id: number;
  state: 'visible' | 'hidden' | 'veryHidden';
  views: WorksheetView[];
  autoFilter: string | { from: string | { row: number; column: number }; to: string | { row: number; column: number } } | null = null;
  properties: Record<string, unknown>;
  /** @internal */
  _rows: (Row | undefined)[] = [];
  /** @internal */
  _columns: Column[] = [];
  /** @internal */
  _merges: Map<string, { top: number; left: number; bottom: number; right: number }> = new Map();
  private readonly _keys = new Map<string, Column>();
  private _name = '';

  /** @internal Use `workbook.addWorksheet()`. */
  constructor(
    readonly workbook: { readonly _names: (exclude: Worksheet) => string[]; readonly definedNames?: DefinedNames; readonly worksheets?: readonly Worksheet[] },
    id: number,
    name: string | undefined,
    options: WorksheetOptions | undefined,
  ) {
    this.id = id;
    this.state = options?.state ?? 'visible';
    this.views = options?.views ? [...options.views] : [];
    this.properties = { ...(options?.properties as Record<string, unknown> | undefined) };
    if (options?.pageSetup !== undefined) notSupported('Worksheet page setup');
    if (options?.headerFooter !== undefined) notSupported('Worksheet header and footer');
    this.name = name ?? `sheet${String(id)}`;
  }

  get name(): string {
    return this._name;
  }

  /** Sheet name rules of exceljs (lib/doc/worksheet.js L134–173), including truncation to 31 characters. */
  set name(value: string) {
    let name = value;
    if (name === this._name) return;
    if (typeof name !== 'string') throw new Error('The name has to be a string.');
    if (name === '') throw new Error("The name can't be empty.");
    if (name === 'History') throw new Error('The name "History" is protected. Please use a different name.');
    if (/[*?:/\\[\]]/.test(name)) throw new Error(`Worksheet name ${name} cannot include any of the following characters: * ? : \\ / [ ]`);
    if (/(^')|('$)/.test(name)) throw new Error(`The first or last character of worksheet name cannot be a single quotation mark: ${name}`);
    if (name.length > 31) name = name.substring(0, 31);
    if (this.workbook._names(this).some((n) => n.toLowerCase() === name.toLowerCase())) throw new Error(`Worksheet name already exists: ${name}`);
    this._name = name;
  }

  /** Deprecated in exceljs in favour of `properties.tabColor`, which it reads and sets (lib/doc/worksheet.js L827). */
  get tabColor(): unknown {
    return this.properties['tabColor'];
  }
  set tabColor(value: unknown) {
    this.properties['tabColor'] = value;
  }

  // ---- columns ------------------------------------------------------------------------------

  get columns(): Column[] {
    return this._columns;
  }

  set columns(defs: ColumnDefn[]) {
    this._columns = [];
    this._keys.clear();
    defs.forEach((defn, i) => {
      const column = new Column(this, i + 1);
      this._columns.push(column);
      column.defn = defn;
    });
  }

  /** @internal */
  findColumn(n: number): Column | undefined {
    return this._columns[n - 1];
  }

  getColumn(c: number | string): Column {
    let n: number;
    if (typeof c === 'string') {
      const keyed = this._keys.get(c);
      if (keyed) return keyed;
      n = columnNumber(c);
    } else n = c;
    if (!(Number.isInteger(n) && n >= 1 && n <= 16_384)) {
      throw new XlvError('XLV_E_INVALID_CELL_REF', `Column ${String(c).slice(0, 20)} is not valid.`, {
        hint: 'Use a column number from 1 to 16384, column letters such as "C", or a column key.',
        details: { ref: String(c).slice(0, 20) },
      });
    }
    while (this._columns.length < n) this._columns.push(new Column(this, this._columns.length + 1));
    return this._columns[n - 1] as Column;
  }

  getColumnKey(key: string): Column | undefined {
    return this._keys.get(key);
  }
  setColumnKey(key: string, column: Column): void {
    this._keys.set(key, column);
  }
  deleteColumnKey(key: string): void {
    this._keys.delete(key);
  }
  eachColumnKey(f: (column: Column, key: string) => void): void {
    for (const [key, column] of this._keys) f(column, key);
  }

  get columnCount(): number {
    let max = 0;
    this.eachRow((row) => {
      max = Math.max(max, row.cellCount);
    });
    return max;
  }

  get actualColumnCount(): number {
    const seen = new Set<number>();
    this.eachRow((row) => row.eachCell((cell) => seen.add(cell.col)));
    return seen.size;
  }

  /**
   * Removes `count` columns from `start` and inserts one column per item of `inserts` there; each
   * insert is an array of values by row (index 0 is row 1). Columns to the right move with their
   * cells, widths, styles, keys and notes, and merged ranges, the autofilter, defined names and
   * formulas follow, as in a spreadsheet application (exceljs 4.4.0 moves values and styles only,
   * docs/exceljs-compat.md). Falsy insert values such as `0` are kept.
   */
  spliceColumns(start: number, count: number, ...inserts: CellValue[][]): void {
    const method = 'worksheet.spliceColumns()';
    checkInteger(method, 'start', start, 1, 16_384);
    checkInteger(method, 'count', count, 0, 16_384);
    inserts.forEach((insert, i) => {
      if (!Array.isArray(insert)) invalidArgument(method, `inserts[${String(i)}]`, insert, 'Give each inserted column as an array of values by row (index 0 is row 1).');
    });
    const change: SpliceChange = { sheet: this.name, axis: 'col', at: start, count, inserts: inserts.length };
    if (this._columns.length >= start + count && this._columns.length - count + inserts.length > 16_384) {
      invalidArgument(method, 'inserts', inserts.length, 'The columns would move beyond column XFD (16384). Remove columns at the end of the sheet first.');
    }
    this._checkFormulas(change, method);
    const merges = this._takeMerges();
    for (const row of this._rows) row?._spliceCells(change, method);
    // Column definitions move as objects, so keys keep pointing at their column.
    const columns: Column[] = [];
    this._columns.forEach((column, i) => {
      const n = movePosition(i + 1, change);
      if (n === null) {
        if (column.key !== undefined && this._keys.get(column.key) === column) this._keys.delete(column.key);
        return;
      }
      column._number = n;
      columns[n - 1] = column;
    });
    for (let i = 0; i < columns.length; i++) if (!columns[i]) columns[i] = new Column(this, i + 1);
    this._columns = columns;
    inserts.forEach((insert, j) => {
      insert.forEach((value, i) => {
        if (value === undefined) return;
        const cell = this.getCell(i + 1, start + j);
        cell.value = value;
        cell.style = {};
      });
    });
    this._restoreMerges(merges, change);
    this._afterSplice(change);
    this.workbook.definedNames?.spliceColumns(this.name, start, count, inserts.length);
    this.workbook.definedNames?.fillInserted(this.name, 'col', start, inserts.length);
  }

  // ---- rows ---------------------------------------------------------------------------------

  /** @internal Rows already written and dropped by a streaming writer still count. */
  _committedRows = 0;

  private get lastRowNumber(): number {
    let n = this._rows.length;
    while (n > 0 && this._rows[n - 1] === undefined) n--;
    return Math.max(n, this._committedRows);
  }

  get rowCount(): number {
    return this.lastRowNumber;
  }

  get actualRowCount(): number {
    let n = 0;
    this.eachRow(() => n++);
    return n;
  }

  get lastRow(): Row | undefined {
    return this._rows.length > 0 ? this._rows[this._rows.length - 1] : undefined;
  }

  findRow(r: number): Row | undefined {
    return this._rows[r - 1];
  }

  getRow(r: number): Row {
    if (!(Number.isInteger(r) && r >= 1 && r <= 1_048_576)) {
      throw new XlvError('XLV_E_INVALID_CELL_REF', `Row ${String(r)} is not valid.`, { hint: 'Use a row number from 1 to 1048576.', details: { ref: String(r) } });
    }
    let row = this._rows[r - 1];
    if (!row) {
      row = new Row(this, r);
      this._rows[r - 1] = row;
    }
    return row;
  }

  getRows(start: number, length: number): Row[] | undefined {
    if (length < 1) return undefined;
    const out: Row[] = [];
    for (let i = start; i < start + length; i++) out.push(this.getRow(i));
    return out;
  }

  /** Adds a row after the last one; style `'i'` / `'i+'` copies the style of the row above (exceljs `addRow`). */
  addRow(values: CellValue[] | Record<string, CellValue>, style = 'n'): Row {
    const rowNo = this.lastRowNumber + 1;
    const row = this.getRow(rowNo);
    row.values = values;
    this._setStyleOption(rowNo, typeof style === 'string' && style[0] === 'i' ? style : 'n');
    return row;
  }

  addRows(rows: (CellValue[] | Record<string, CellValue>)[], style = 'n'): Row[] {
    return rows.map((r) => this.addRow(r, style));
  }

  /**
   * Inserts a row at `pos`; the rows from `pos` down move one row down. `style` copies a style as
   * exceljs does: `'i'` from the row above, `'o'` from the row below (the one that moved down),
   * with `+` also for empty cells, `'n'` none.
   */
  insertRow(pos: number, value?: CellValue[] | Record<string, CellValue> | null, style = 'n'): Row {
    checkInteger('worksheet.insertRow()', 'pos', pos, 1, 1_048_576);
    this.spliceRows(pos, 0, value ?? null);
    this._setStyleOption(pos, style);
    return this.getRow(pos);
  }

  /** Inserts one row per item of `values` at `pos`, with the style options of `insertRow`. */
  insertRows(pos: number, values: (CellValue[] | Record<string, CellValue> | null)[], style = 'n'): Row[] | undefined {
    const method = 'worksheet.insertRows()';
    checkInteger(method, 'pos', pos, 1, 1_048_576);
    if (!Array.isArray(values)) invalidArgument(method, 'values', values, 'Pass an array with one item (an array or object of values) per row to insert.');
    this.spliceRows(pos, 0, ...values);
    if (style !== 'n') {
      // exceljs worksheet.js L432–439.
      for (let i = 0; i < values.length; i++) {
        if (style[0] === 'o' && this.findRow(values.length + pos + i) !== undefined) this._copyStyle(values.length + pos + i, pos + i, style[1] === '+');
        else if (style[0] === 'i' && this.findRow(pos - 1) !== undefined) this._copyStyle(pos - 1, pos + i, style[1] === '+');
      }
    }
    return this.getRows(pos, values.length);
  }

  /**
   * Makes `count` copies of row `rowNum` below it, inserting new rows with `insert = true` or
   * replacing the rows below with `insert = false` (the default of exceljs's code; its README says
   * `true`). Copies get the values, styles and height, the merged ranges inside the row, and
   * formulas moved with the row.
   */
  duplicateRow(rowNum: number, count: number, insert = false): void {
    const method = 'worksheet.duplicateRow()';
    checkInteger(method, 'rowNum', rowNum, 1, 1_048_576);
    checkInteger(method, 'count', count, 0, 1_048_575);
    const src = this.findRow(rowNum);
    if (!src) invalidArgument(method, 'rowNum', rowNum, `Row ${String(rowNum)} does not exist; duplicate a row that has cells.`);
    if (count === 0) return;
    const values: CellValue[] = [];
    src._cells.forEach((cell) => {
      if (cell && !cell._master) values[cell.col] = cell.value;
    });
    src._cells.forEach((cell) => {
      const v = cell && !cell._master ? cell.value : null;
      const formula = v !== null && typeof v === 'object' && !(v instanceof Date) ? (v as { formula?: unknown }).formula : undefined;
      if (cell && typeof formula === 'string' && (copyFormula(formula, 1) === null || copyFormula(formula, count) === null)) {
        formulaNotAdjustable(method, this.name, cell.address, formula);
      }
    });
    const rowMerges = [...this._merges.values()].filter((m) => m.top === rowNum && m.bottom === rowNum);
    this.spliceRows(rowNum + 1, insert ? 0 : count, ...new Array<CellValue[]>(count).fill(values));
    for (let i = 1; i <= count; i++) {
      const dst = this.getRow(rowNum + i);
      dst.style = copyPlain(src.style);
      dst.height = src.height;
      src.eachCell({ includeEmpty: true }, (cell, col) => {
        const copy = dst.getCell(col);
        copy.style = copyPlain(cell.style);
        const v = copy.value;
        if (v !== null && typeof v === 'object' && !(v instanceof Date) && typeof (v as { formula?: unknown }).formula === 'string') {
          copy.value = { ...(v as Record<string, unknown>), formula: copyFormula((v as { formula: string }).formula, i) ?? (v as { formula: string }).formula } as CellValue;
        }
      });
      for (const m of rowMerges) this._mergeRange({ top: rowNum + i, bottom: rowNum + i, left: m.left, right: m.right });
    }
  }

  /**
   * Removes `count` rows from `start` and inserts one row per item of `inserts` (an array or object
   * of values) there. Rows below move with their cells, styles, heights, hidden flag and notes, and
   * merged ranges, the autofilter, defined names and formulas follow, as in a spreadsheet
   * application. exceljs 4.4.0 moves values, styles and heights only (docs/exceljs-compat.md).
   */
  spliceRows(start: number, count: number, ...inserts: (CellValue[] | Record<string, CellValue> | null | undefined)[]): void {
    const method = 'worksheet.spliceRows()';
    checkInteger(method, 'start', start, 1, 1_048_576);
    checkInteger(method, 'count', count, 0, 1_048_576);
    const change: SpliceChange = { sheet: this.name, axis: 'row', at: start, count, inserts: inserts.length };
    const last = this.lastRowNumber;
    if (last >= start + count && last - count + inserts.length > 1_048_576) {
      invalidArgument(method, 'inserts', inserts.length, 'The rows would move beyond row 1048576. Remove rows at the end of the sheet first.');
    }
    this._checkFormulas(change, method);
    const merges = this._takeMerges();
    const rows: (Row | undefined)[] = [];
    this._rows.forEach((row, i) => {
      if (!row) return;
      const n = movePosition(i + 1, change);
      if (n === null) return;
      row._number = n;
      rows[n - 1] = row;
    });
    this._rows = rows;
    inserts.forEach((value, i) => {
      const row = this.getRow(start + i);
      row.style = {};
      row.values = value ?? null;
    });
    this._restoreMerges(merges, change);
    this._afterSplice(change);
    this.workbook.definedNames?.spliceRows(this.name, start, count, inserts.length);
    this.workbook.definedNames?.fillInserted(this.name, 'row', start, inserts.length);
  }

  /** exceljs `_setStyleOption` (worksheet.js L444–450). */
  private _setStyleOption(pos: number, style = 'n'): void {
    if (style[0] === 'o' && this.findRow(pos + 1) !== undefined) this._copyStyle(pos + 1, pos, style[1] === '+');
    else if (style[0] === 'i' && this.findRow(pos - 1) !== undefined) this._copyStyle(pos - 1, pos, style[1] === '+');
  }

  /** exceljs `_copyStyle` (worksheet.js L452–461), with copies instead of shared style objects. */
  private _copyStyle(src: number, dest: number, styleEmpty: boolean): void {
    const rSrc = this.getRow(src);
    const rDst = this.getRow(dest);
    rDst.style = copyPlain(rSrc.style);
    rSrc.eachCell({ includeEmpty: styleEmpty }, (cell, col) => {
      rDst.getCell(col).style = copyPlain(cell.style);
    });
    rDst.height = rSrc.height;
  }

  /** @internal Removes every merged range and returns them, for a splice. */
  _takeMerges(): { top: number; left: number; bottom: number; right: number }[] {
    const merges = [...this._merges.values()];
    for (const m of merges) {
      for (let r = m.top; r <= m.bottom; r++) for (let c = m.left; c <= m.right; c++) this.findCell(r, c)?._unmerge();
      const master = this.findCell(m.top, m.left);
      if (master) master._mergeCount = 0;
    }
    this._merges.clear();
    return merges;
  }

  /**
   * @internal Merges the ranges again at their new place: a range loses the removed rows or
   * columns, grows when rows or columns are inserted inside it, and disappears when one cell or
   * nothing is left. When the top-left cell was removed, the new top-left cell becomes the master
   * with its own (empty) value.
   */
  _restoreMerges(merges: readonly { top: number; left: number; bottom: number; right: number }[], change: SpliceChange): void {
    for (const m of merges) {
      const span = change.axis === 'row' ? moveSpan(m.top, m.bottom, change) : moveSpan(m.left, m.right, change);
      if (!span) continue;
      const range = change.axis === 'row' ? { top: span[0], bottom: span[1], left: m.left, right: m.right } : { top: m.top, bottom: m.bottom, left: span[0], right: span[1] };
      if (range.top === range.bottom && range.left === range.right) continue;
      this._mergeRange(range);
    }
  }

  /** @internal Merges a range without the overlap check of `mergeCells`. */
  _mergeRange(range: { top: number; left: number; bottom: number; right: number }): void {
    const master = this.getCell(range.top, range.left);
    master._unmerge();
    for (let r = range.top; r <= range.bottom; r++) {
      for (let c = range.left; c <= range.right; c++) {
        if (r === range.top && c === range.left) continue;
        const cell = this.getCell(r, c);
        cell._merge(master);
        cell.style = { ...master.style };
      }
    }
    master._mergeCount++;
    this._merges.set(`${columnLetter(range.left)}${String(range.top)}`, range);
  }

  /** @internal The autofilter range and every formula of the workbook follow a splice of this sheet. */
  _afterSplice(change: SpliceChange): void {
    if (this.autoFilter) {
      const f = this.autoFilter;
      const text = typeof f === 'string' ? f : `${typeof f.from === 'string' ? f.from : `${columnLetter(f.from.column)}${String(f.from.row)}`}:${typeof f.to === 'string' ? f.to : `${columnLetter(f.to.column)}${String(f.to.row)}`}`;
      const r = parseRangeArgs([text]);
      const span = change.axis === 'row' ? moveSpan(r.top, r.bottom, change) : moveSpan(r.left, r.right, change);
      if (!span) this.autoFilter = null;
      else {
        const nr = change.axis === 'row' ? { ...r, top: span[0], bottom: span[1] } : { ...r, left: span[0], right: span[1] };
        if (typeof f === 'string') this.autoFilter = `${columnLetter(nr.left)}${String(nr.top)}:${columnLetter(nr.right)}${String(nr.bottom)}`;
        else {
          const end = (x: string | { row: number; column: number }, row: number, col: number): string | { row: number; column: number } =>
            typeof x === 'string' ? `${columnLetter(col)}${String(row)}` : { row, column: col };
          this.autoFilter = { from: end(f.from, nr.top, nr.left), to: end(f.to, nr.bottom, nr.right) };
        }
      }
    }
    const names = this._sheetNames();
    this._eachFormula((sheet, cell, v, formula) => {
      const text = adjustFormulaText(formula, sheet.name, change, names);
      if (text !== null && text !== formula) cell.value = { ...v, formula: text } as CellValue;
    });
  }

  /** @internal Throws before anything changes when a formula of the workbook cannot follow `change`. */
  _checkFormulas(change: SpliceChange, method: string): void {
    const names = this._sheetNames();
    this._eachFormula((sheet, cell, _v, formula) => {
      if (adjustFormulaText(formula, sheet.name, change, names) === null) formulaNotAdjustable(method, sheet.name, cell.address, formula);
    });
  }

  private _sheetNames(): string[] {
    return (this.workbook.worksheets ?? [this]).map((s) => s.name);
  }

  private _eachFormula(fn: (sheet: Worksheet, cell: Cell, value: Record<string, unknown>, formula: string) => void): void {
    for (const sheet of this.workbook.worksheets ?? [this]) {
      for (const row of sheet._rows) {
        if (!row) continue;
        for (const cell of row._cells) {
          if (!cell || cell._master) continue;
          const v = cell.value;
          if (v === null || typeof v !== 'object' || v instanceof Date) continue;
          const formula = (v as { formula?: unknown }).formula;
          if (typeof formula === 'string') fn(sheet, cell, v as Record<string, unknown>, formula);
        }
      }
    }
  }

  eachRow(options: { includeEmpty?: boolean } | ((row: Row, n: number) => void), iteratee?: (row: Row, n: number) => void): void {
    const cb = typeof options === 'function' ? options : iteratee;
    const includeEmpty = typeof options === 'object' && options !== null && options.includeEmpty === true;
    if (!cb) return;
    if (includeEmpty) {
      const n = this._rows.length;
      for (let i = 1; i <= n; i++) cb(this.getRow(i), i);
      return;
    }
    this._rows.forEach((row) => {
      if (row && row.hasValues) cb(row, row.number);
    });
  }

  getSheetValues(): CellValue[][] {
    const out: CellValue[][] = [];
    this.eachRow((row) => {
      out[row.number] = row.values;
    });
    return out;
  }

  /** @internal Overridden by the streaming writer. */
  _commitRow(_row: Row): void {}

  // ---- cells --------------------------------------------------------------------------------

  findCell(r: number | string, c?: number): Cell | undefined {
    const { row, col } = typeof r === 'string' ? decodeAddress(r) : { row: r, col: c ?? 0 };
    return this.findRow(row)?.findCell(col);
  }

  getCell(r: number | string, c?: number | string): Cell {
    if (typeof r === 'string' && c === undefined) {
      const { row, col } = decodeAddress(r);
      return this.getRow(row).getCell(col);
    }
    return this.getRow(Number(r)).getCell(c as number | string);
  }

  get dimensions(): { top: number; left: number; bottom: number; right: number; tl: string; br: string } | null {
    let top = 0;
    let left = 0;
    let bottom = 0;
    let right = 0;
    this.eachRow((row) => {
      const d = row.dimensions;
      if (!d) return;
      if (!top) top = row.number;
      bottom = row.number;
      if (!left || d.min < left) left = d.min;
      if (d.max > right) right = d.max;
    });
    if (!top) return null;
    return { top, left, bottom, right, tl: `${columnLetter(left)}${String(top)}`, br: `${columnLetter(right)}${String(bottom)}` };
  }

  // ---- merges -------------------------------------------------------------------------------

  mergeCells(...args: unknown[]): void {
    const range = parseRangeArgs(args);
    for (const m of this._merges.values()) {
      if (range.top <= m.bottom && m.top <= range.bottom && range.left <= m.right && m.left <= range.right) {
        throw new Error('Cannot merge already merged cells');
      }
    }
    const master = this.getCell(range.top, range.left);
    for (let r = range.top; r <= range.bottom; r++) {
      for (let c = range.left; c <= range.right; c++) {
        if (r === range.top && c === range.left) continue;
        const cell = this.getCell(r, c);
        cell._merge(master);
        cell.style = { ...master.style };
      }
    }
    master._mergeCount++;
    this._merges.set(`${columnLetter(range.left)}${String(range.top)}`, range);
  }

  unMergeCells(...args: unknown[]): void {
    const range = parseRangeArgs(args);
    for (const [key, m] of this._merges) {
      if (range.top <= m.bottom && m.top <= range.bottom && range.left <= m.right && m.left <= range.right) {
        for (let r = m.top; r <= m.bottom; r++) for (let c = m.left; c <= m.right; c++) this.findCell(r, c)?._unmerge();
        const master = this.findCell(m.top, m.left);
        if (master) master._mergeCount = Math.max(0, master._mergeCount - 1);
        this._merges.delete(key);
      }
    }
  }

  get hasMerges(): boolean {
    return this._merges.size > 0;
  }

  // ---- features not supported ------------------------------------------------------------------

  addImage(): never {
    return notSupported('Images (worksheet.addImage)');
  }
  addBackgroundImage(): never {
    return notSupported('Background images');
  }
  addTable(): never {
    return notSupported('Tables (worksheet.addTable)');
  }
  addConditionalFormatting(): never {
    return notSupported('Conditional formatting');
  }
  protect(): never {
    return notSupported('Worksheet protection');
  }
  get dataValidations(): never {
    return notSupported('Data validation');
  }
}

export { shallowStyle };
