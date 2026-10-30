// Tables and images of the exceljs model, following exceljs 4.4.0 (lib/doc/table.js, image.js,
// anchor.js; docs/facts/exceljs.md X4-21 … X4-23). Writing and reading go through xlvane's
// tables and pictures (docs/decisions.md D-D02 … D-D04). Where exceljs writes a file Excel or the
// schema rejects, xlvane writes a valid one; each difference is in docs/exceljs-compat.md.
import { createWorkbook, imageFormat, XlvError, type SheetImage, type TableDefinition, type WorkbookImage } from 'xlvane';
import type { Style } from './types';
import type { Worksheet } from './model';

const MAX_COLS = 16_384;
const MAX_ROWS = 1_048_576;

function letters(n: number): string {
  let s = '';
  for (let c = n; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s;
  return s;
}

function decode(address: string): { row: number; col: number } {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(address.trim());
  if (!m) throw invalidTable(`"${address}" is not a cell address.`, 'Give the top-left cell of the table, for example "B2".');
  let col = 0;
  for (const ch of m[1]!.toUpperCase()) col = col * 26 + ch.charCodeAt(0) - 64;
  return { row: Number(m[2]), col };
}

function encode(top: number, left: number, bottom: number, right: number): string {
  return `${letters(left)}${String(top)}:${letters(right)}${String(bottom)}`;
}

function invalidTable(message: string, hint: string): XlvError {
  return new XlvError('XLV_E_INVALID_TABLE', message, { hint });
}

// ---- tables -----------------------------------------------------------------------------------

/** The totals functions of a table column (exceljs `TableColumnProperties.totalsRowFunction`). */
type TotalsFunction = 'none' | 'average' | 'countNums' | 'count' | 'max' | 'min' | 'stdDev' | 'var' | 'sum' | 'custom';

/**
 * A column of a table, as in exceljs `TableColumnProperties`. `style` is applied to the column's
 * cells (header, data and totals), as exceljs does.
 *
 * @example
 * ```ts
 * import type { TableColumnProperties } from '@xlvane/exceljs';
 *
 * const column: TableColumnProperties = { name: 'Revenue', totalsRowFunction: 'sum', filterButton: true };
 * ```
 */
export interface TableColumnProperties {
  name: string;
  filterButton?: boolean | undefined;
  totalsRowLabel?: string | undefined;
  totalsRowFunction?: TotalsFunction | undefined;
  totalsRowFormula?: string | undefined;
  totalsRowResult?: unknown;
  style?: Partial<Style> | undefined;
}

/**
 * The definition of a table, as in exceljs `TableProperties` (`worksheet.addTable`). `ref` is the
 * top-left cell; the table covers the header row, the rows and the totals row.
 *
 * @example
 * ```ts
 * import type { TableProperties } from '@xlvane/exceljs';
 *
 * const table: TableProperties = {
 *   name: 'Sales',
 *   ref: 'A1',
 *   totalsRow: true,
 *   columns: [{ name: 'Region', totalsRowLabel: 'Total' }, { name: 'Revenue', totalsRowFunction: 'sum' }],
 *   rows: [['North', 1200], ['South', 980]],
 * };
 * ```
 */
export interface TableProperties {
  name: string;
  displayName?: string | undefined;
  ref: string;
  headerRow?: boolean | undefined;
  totalsRow?: boolean | undefined;
  style?: { theme?: string | null | undefined; showFirstColumn?: boolean | undefined; showLastColumn?: boolean | undefined; showRowStripes?: boolean | undefined; showColumnStripes?: boolean | undefined } | undefined;
  columns: TableColumnProperties[];
  rows: unknown[][];
}

/** The model of a table: exceljs's fields plus the computed ranges (lib/doc/table.js L129–178). */
interface TableModel extends TableProperties {
  tl: { row: number; col: number };
  autoFilterRef: string;
  tableRef: string;
}

/**
 * The SUBTOTAL function numbers of the totals row (exceljs table.js L81–107). exceljs uses 106
 * (PRODUCT) for `stdDev` and 107 (STDEV) for `var`; xlvane writes the numbers of [MS-OI29500]
 * §2.1.1094 (107 and 110), as the core does (D-D02).
 */
const SUBTOTAL: Readonly<Record<string, number>> = { average: 101, countNums: 102, count: 103, max: 104, min: 105, stdDev: 107, sum: 109, var: 110 };

/**
 * A column of a table, as exceljs `Table.getColumn()` returns it. Changing `name` or the totals
 * settings takes effect at `table.commit()`.
 *
 * @example
 * ```ts
 * import ExcelJS from '@xlvane/exceljs';
 *
 * const sheet = new ExcelJS.Workbook().addWorksheet('Data');
 * const table = sheet.addTable({ name: 'Sales', ref: 'A1', columns: [{ name: 'Region' }, { name: 'Revenue' }], rows: [['North', 1200]] });
 * table.getColumn(1).totalsRowFunction = 'sum';
 * table.totalsRow = true;
 * table.commit();
 * ```
 */
export class TableColumn {
  /** @internal */
  constructor(
    readonly table: Table,
    readonly column: TableColumnProperties,
    readonly index: number,
  ) {}

  private set_<K extends keyof TableColumnProperties>(key: K, value: TableColumnProperties[K]): void {
    this.table.cacheState();
    this.column[key] = value;
  }

  get name(): string {
    return this.column.name;
  }
  set name(v: string) {
    this.set_('name', v);
  }
  get filterButton(): boolean | undefined {
    return this.column.filterButton;
  }
  set filterButton(v: boolean | undefined) {
    this.column.filterButton = v;
  }
  get style(): Partial<Style> | undefined {
    return this.column.style;
  }
  set style(v: Partial<Style> | undefined) {
    this.column.style = v;
  }
  get totalsRowLabel(): string | undefined {
    return this.column.totalsRowLabel;
  }
  set totalsRowLabel(v: string | undefined) {
    this.set_('totalsRowLabel', v);
  }
  get totalsRowFunction(): TotalsFunction | undefined {
    return this.column.totalsRowFunction;
  }
  set totalsRowFunction(v: TotalsFunction | undefined) {
    this.set_('totalsRowFunction', v);
  }
  get totalsRowResult(): unknown {
    return this.column.totalsRowResult;
  }
  set totalsRowResult(v: unknown) {
    this.set_('totalsRowResult', v);
  }
  get totalsRowFormula(): string | undefined {
    return this.column.totalsRowFormula;
  }
  set totalsRowFormula(v: string | undefined) {
    this.set_('totalsRowFormula', v);
  }
}

/**
 * A table on a worksheet, as exceljs `Table` (lib/doc/table.js). `worksheet.addTable()` writes
 * the header, rows and totals into the cells; methods that change the table take effect at
 * `commit()`, which clears the old area and writes the new one. Names and ranges follow Excel's
 * rules (`XLV_E_INVALID_TABLE`), which exceljs does not check.
 *
 * @example
 * ```ts
 * import ExcelJS from '@xlvane/exceljs';
 *
 * const sheet = new ExcelJS.Workbook().addWorksheet('Data');
 * const table = sheet.addTable({
 *   name: 'Sales',
 *   ref: 'B2',
 *   columns: [{ name: 'Region' }, { name: 'Revenue' }],
 *   rows: [['North', 1200]],
 * });
 * table.addRow(['South', 980]);
 * table.commit();
 * console.log(table.ref, sheet.getCell('C4').value); // "B2" 980
 * ```
 */
export class Table {
  /** @internal */
  table: TableModel;
  private _cache: { ref: string; width: number; tableHeight: number } | undefined;

  /** @internal Use `worksheet.addTable()`. */
  constructor(
    readonly worksheet: Worksheet,
    table: TableProperties,
    store = true,
  ) {
    this.table = table as TableModel;
    this.validate();
    if (store) this.store();
  }

  get width(): number {
    return this.table.columns.length;
  }
  get height(): number {
    return this.table.rows.length;
  }
  get filterHeight(): number {
    return this.height + (this.table.headerRow ? 1 : 0);
  }
  get tableHeight(): number {
    return this.filterHeight + (this.table.totalsRow ? 1 : 0);
  }

  /** The formula of a column's totals cell, or `null` (exceljs `getFormula`, with the codes of [MS-OI29500]). */
  getFormula(column: TableColumnProperties): string | null {
    const fn = column.totalsRowFunction ?? 'none';
    if (fn === 'none') return null;
    if (fn === 'custom') return column.totalsRowFormula ?? null;
    const code = SUBTOTAL[fn];
    if (code === undefined) throw invalidTable(`"${String(fn)}" is not a totals row function.`, 'Use none, average, countNums, count, max, min, stdDev, var, sum or custom.');
    return `SUBTOTAL(${String(code)},${this.table.name}[${column.name}])`;
  }

  /**
   * Sets the defaults of exceljs and computes the ranges (table.js L129–178), then checks the
   * table against Excel's rules through xlvane, so that a table Excel would reject is refused
   * here rather than when the file is written.
   */
  validate(): void {
    const t = this.table;
    if (t.headerRow === undefined) t.headerRow = true;
    if (t.totalsRow === undefined) t.totalsRow = false;
    t.style ??= {};
    if (t.style.theme === undefined) t.style.theme = 'TableStyleMedium2';
    t.style.showFirstColumn ??= false;
    t.style.showLastColumn ??= false;
    t.style.showRowStripes ??= false;
    t.style.showColumnStripes ??= false;
    if (!t.ref) throw invalidTable('A table needs a ref (its top-left cell).', "For example ref: 'A1'.");
    if (!Array.isArray(t.columns)) throw invalidTable('A table needs column definitions.', "For example columns: [{ name: 'Region' }, { name: 'Revenue' }].");
    if (!Array.isArray(t.rows)) throw invalidTable('A table needs row definitions.', 'Give rows: [] for a table without data rows.');
    t.tl = decode(t.ref);
    const { row, col } = t.tl;
    const bottom = row + this.tableHeight - 1;
    const right = col + this.width - 1;
    if (bottom > MAX_ROWS || right > MAX_COLS) throw invalidTable(`The table "${t.name}" would reach beyond the sheet.`, 'Start it higher or further left, or give it fewer rows or columns.');
    t.autoFilterRef = encode(row, col, row + this.filterHeight - 1, right);
    t.tableRef = encode(row, col, bottom, right);
    t.columns.forEach((column, i) => {
      if (!column || !column.name) throw invalidTable(`Column ${String(i)} of table "${t.name}" must have a name.`, "For example { name: 'Revenue' }.");
      // exceljs labels the first column "Total" and computes the others' formulas (table.js L169–176).
      if (i === 0) column.totalsRowLabel ??= 'Total';
      else {
        column.totalsRowFunction ??= 'none';
        if (column.totalsRowFunction !== 'custom') column.totalsRowFormula = this.getFormula(column) ?? undefined;
      }
    });
    // Excel's rules (names, columns, range), checked by the core on a scratch workbook.
    createWorkbook().addSheet('S').addTable(tableInput(this));
  }

  /** Writes the header, rows and totals into the worksheet's cells (exceljs `store`, table.js L180–234). */
  store(): void {
    const t = this.table;
    const { row, col } = t.tl;
    const assignStyle = (address: { row: number; col: number }, style: Partial<Style> | undefined): void => {
      if (!style) return;
      const cell = this.worksheet.getCell(address.row, address.col) as unknown as Record<string, unknown>;
      for (const key of Object.keys(style)) cell[key] = (style as Record<string, unknown>)[key];
    };
    let r = row;
    if (t.headerRow) {
      t.columns.forEach((c, j) => {
        this.worksheet.getCell(r, col + j).value = c.name;
        assignStyle({ row: r, col: col + j }, c.style);
      });
      r++;
    }
    for (const data of t.rows) {
      (data ?? []).forEach((value, j) => {
        this.worksheet.getCell(r, col + j).value = value as never;
        assignStyle({ row: r, col: col + j }, t.columns[j]?.style);
      });
      r++;
    }
    if (t.totalsRow) {
      t.columns.forEach((c, j) => {
        const cell = this.worksheet.getCell(r, col + j);
        const formula = this.getFormula(c);
        if (formula !== null) cell.value = (c.totalsRowResult === undefined ? { formula } : { formula, result: c.totalsRowResult }) as never;
        else cell.value = j === 0 || (c.totalsRowLabel !== undefined && (c.totalsRowFunction ?? 'none') === 'none') ? (c.totalsRowLabel ?? null) : null;
        assignStyle({ row: r, col: col + j }, c.style);
      });
    }
  }

  /** @internal Remembers the table's area before the first change, for `commit()`. */
  cacheState(): void {
    this._cache ??= { ref: this.ref, width: this.width, tableHeight: this.tableHeight };
  }

  /**
   * Applies the changes made since the last commit: clears the area the table left and writes
   * it again (table.js L295–335). The cache is cleared, so later changes are committed too
   * (exceljs keeps the first cache and clears the wrong cells after a second change).
   */
  commit(): void {
    const cache = this._cache;
    if (!cache) return;
    this.validate();
    const at = decode(cache.ref);
    const clear = (i: number, j: number): void => {
      this.worksheet.getCell(at.row + i, at.col + j).value = null;
    };
    if (this.ref !== cache.ref) {
      for (let i = 0; i < cache.tableHeight; i++) for (let j = 0; j < cache.width; j++) clear(i, j);
    } else {
      for (let i = this.tableHeight; i < cache.tableHeight; i++) for (let j = 0; j < cache.width; j++) clear(i, j);
      for (let i = 0; i < this.tableHeight; i++) for (let j = this.width; j < cache.width; j++) clear(i, j);
    }
    this._cache = undefined;
    this.store();
  }

  /** Adds a row of data at the end, or before the row with index `rowNumber` (0-based). */
  addRow(values: unknown[], rowNumber?: number): void {
    this.cacheState();
    if (rowNumber === undefined) this.table.rows.push(values);
    else this.table.rows.splice(rowNumber, 0, values);
  }

  /** Removes `count` rows of data from index `rowIndex` (0-based). */
  removeRows(rowIndex: number, count = 1): void {
    this.cacheState();
    this.table.rows.splice(rowIndex, count);
  }

  /** The column with index `colIndex` (0-based). */
  getColumn(colIndex: number): TableColumn {
    const column = this.table.columns[colIndex];
    if (!column) throw invalidTable(`Table "${this.name}" has no column ${String(colIndex)}.`, `Use an index from 0 to ${String(this.width - 1)}.`);
    return new TableColumn(this, column, colIndex);
  }

  /** Adds a column with its values, at the end or before index `colIndex` (0-based). */
  addColumn(column: TableColumnProperties, values: unknown[], colIndex?: number): void {
    this.cacheState();
    if (colIndex === undefined) {
      this.table.columns.push(column);
      this.table.rows.forEach((row, i) => row.push(values[i]));
    } else {
      this.table.columns.splice(colIndex, 0, column);
      this.table.rows.forEach((row, i) => row.splice(colIndex, 0, values[i]));
    }
  }

  /** Removes `count` columns from index `colIndex` (0-based), with their values. */
  removeColumns(colIndex: number, count = 1): void {
    this.cacheState();
    this.table.columns.splice(colIndex, count);
    this.table.rows.forEach((row) => row.splice(colIndex, count));
  }

  /** The table's model, as exceljs `table.model`, with `ref` and `rows`. */
  get model(): TableProperties & { tableRef: string; autoFilterRef: string } {
    return this.table;
  }
  set model(v: TableProperties & { tableRef: string; autoFilterRef: string }) {
    this.table = v as TableModel;
  }

  get ref(): string {
    return this.table.ref;
  }
  set ref(v: string) {
    this.cacheState();
    this.table.ref = v;
  }
  get name(): string {
    return this.table.name;
  }
  set name(v: string) {
    this.table.name = v;
  }
  /** The display name; `name` when not set. exceljs reads a misspelled field here (`displyName`). */
  get displayName(): string {
    return this.table.displayName ?? this.table.name;
  }
  set displayName(v: string) {
    this.table.displayName = v;
  }
  get headerRow(): boolean {
    return this.table.headerRow !== false;
  }
  set headerRow(v: boolean) {
    this.cacheState();
    this.table.headerRow = v;
  }
  get totalsRow(): boolean {
    return this.table.totalsRow === true;
  }
  set totalsRow(v: boolean) {
    this.cacheState();
    this.table.totalsRow = v;
  }
  /** The table style name. exceljs reads and writes `style.name` here, which is never written to the file. */
  get theme(): string | null | undefined {
    return this.table.style?.theme;
  }
  set theme(v: string | null | undefined) {
    this.table.style = { ...this.table.style, theme: v };
  }
  get showFirstColumn(): boolean {
    return this.table.style?.showFirstColumn === true;
  }
  set showFirstColumn(v: boolean) {
    this.table.style = { ...this.table.style, showFirstColumn: v };
  }
  get showLastColumn(): boolean {
    return this.table.style?.showLastColumn === true;
  }
  set showLastColumn(v: boolean) {
    this.table.style = { ...this.table.style, showLastColumn: v };
  }
  get showRowStripes(): boolean {
    return this.table.style?.showRowStripes === true;
  }
  set showRowStripes(v: boolean) {
    this.table.style = { ...this.table.style, showRowStripes: v };
  }
  get showColumnStripes(): boolean {
    return this.table.style?.showColumnStripes === true;
  }
  set showColumnStripes(v: boolean) {
    this.table.style = { ...this.table.style, showColumnStripes: v };
  }
}

/**
 * The core form of a table, without rows: the layer writes the cells itself and the core adds the
 * header names, the totals formulas and the table part (D-D02, D-D04).
 */
export function tableInput(table: Table): {
  name: string;
  displayName?: string;
  ref: string;
  headerRow: boolean;
  totalsRow: boolean;
  style: { theme: string | null; showFirstColumn: boolean; showLastColumn: boolean; showRowStripes: boolean; showColumnStripes: boolean };
  columns: { name: string; totalsRowFunction?: TotalsFunction; totalsRowLabel?: string; totalsRowFormula?: string; filterButton: boolean }[];
} {
  const t = table.table;
  const style = t.style ?? {};
  // Excel needs at least one data row (D-D02); exceljs writes a table without one, so a table
  // without rows gets one empty data row in the file.
  if (t.rows.length === 0 && t.totalsRow === true) {
    throw invalidTable(`Table "${t.name}" has a totals row but no data rows.`, 'Excel needs at least one data row in a table; add one, e.g. rows: [[]] for an empty row.');
  }
  const extra = (t.headerRow !== false ? 1 : 0) + (t.totalsRow === true ? 1 : 0);
  const bottom = t.tl.row + extra + Math.max(t.rows.length, 1) - 1;
  return {
    name: t.name,
    ...(t.displayName !== undefined && t.displayName !== t.name ? { displayName: t.displayName } : {}),
    ref: encode(t.tl.row, t.tl.col, bottom, t.tl.col + t.columns.length - 1),
    headerRow: t.headerRow !== false,
    totalsRow: t.totalsRow === true,
    style: {
      theme: style.theme === undefined ? 'TableStyleMedium2' : style.theme,
      showFirstColumn: style.showFirstColumn === true,
      showLastColumn: style.showLastColumn === true,
      showRowStripes: style.showRowStripes === true,
      showColumnStripes: style.showColumnStripes === true,
    },
    columns: t.columns.map((c) => {
      const fn = c.totalsRowFunction ?? 'none';
      return {
        name: c.name,
        ...(fn === 'none' ? {} : { totalsRowFunction: fn }),
        // exceljs labels column 0 "Total" even when it has a function; a cell holds either a label
        // or a function in Excel ([MS-OI29500] §2.1.681), so the label is dropped then (X4-21).
        ...(fn === 'none' && c.totalsRowLabel !== undefined ? { totalsRowLabel: c.totalsRowLabel } : {}),
        ...(fn === 'custom' && c.totalsRowFormula !== undefined ? { totalsRowFormula: c.totalsRowFormula } : {}),
        // filter-column-xform.js L35–51: the button is hidden unless filterButton is true.
        filterButton: c.filterButton === true,
      };
    }),
  };
}

/**
 * A table read from a file, in exceljs's shape (table-xform.js parse, X4-21) and with `ref` and
 * the rows taken from the cells, so that it can be changed and written again.
 */
export function tableOut(ws: Worksheet, def: TableDefinition): Table {
  const [first, last] = def.ref.split(':') as [string, string | undefined];
  const tl = decode(first);
  const br = decode(last ?? first);
  const headerRow = def.headerRowCount > 0;
  const totalsRow = def.totalsRowCount > 0;
  const rows: unknown[][] = [];
  for (let r = tl.row + (headerRow ? 1 : 0); r <= br.row - (totalsRow ? 1 : 0); r++) {
    const values: unknown[] = [];
    for (let c = tl.col; c <= br.col; c++) {
      const v = ws.findRow(r)?.findCell(c)?.value;
      values.push(v === undefined ? null : v);
    }
    rows.push(values);
  }
  const model: TableProperties = {
    name: def.name,
    displayName: def.displayName,
    ref: first,
    headerRow,
    totalsRow,
    style: def.style ? { ...def.style } : { theme: null, showFirstColumn: false, showLastColumn: false, showRowStripes: false, showColumnStripes: false },
    columns: def.columns.map((c) => ({
      name: c.name,
      filterButton: c.filterButton,
      ...(c.totalsRowLabel === null ? {} : { totalsRowLabel: c.totalsRowLabel }),
      totalsRowFunction: c.totalsRowFunction,
      ...(c.totalsRowFormula === null ? {} : { totalsRowFormula: c.totalsRowFormula }),
    })),
    rows,
  };
  return new Table(ws, model, false);
}

// ---- images -----------------------------------------------------------------------------------

/**
 * An image of the workbook, as exceljs `Image` (`workbook.addImage`). Give the file contents as
 * `buffer` (bytes) or `base64` (a data URL or plain base64). xlvane recognises PNG, JPEG and GIF
 * from the bytes; `extension` is kept but not used. `filename` is not supported: read the file
 * and pass `buffer`.
 *
 * @example
 * ```ts
 * import type { ImageData } from '@xlvane/exceljs';
 *
 * declare const png: Uint8Array;
 * const image: ImageData = { buffer: png, extension: 'png' };
 * ```
 */
export interface ImageData {
  extension?: string;
  buffer?: Uint8Array;
  base64?: string;
  filename?: string;
  /** Set when read from a file: the media part's name without extension, as exceljs reads it. */
  name?: string;
  /** `"image"`, as in exceljs's media list. */
  type?: 'image';
}

/** A corner of a picture as exceljs models it (lib/doc/anchor.js): a 0-based cell and an offset in EMU. */
interface AnchorModel {
  nativeCol: number;
  nativeColOff: number;
  nativeRow: number;
  nativeRowOff: number;
}

/**
 * A corner of a picture, as exceljs `Anchor` (lib/doc/anchor.js): `nativeCol` and `nativeRow`
 * are 0-based, the offsets are in EMU, and `col`/`row` give the position as a fraction of the
 * cell, using exceljs's cell sizes (640000 per default column, 180000 per default row,
 * `width * 10000` and `height * 10000` for custom ones).
 *
 * @example
 * ```ts
 * import ExcelJS from '@xlvane/exceljs';
 *
 * const workbook = new ExcelJS.Workbook();
 * const sheet = workbook.addWorksheet('Report');
 * declare const png: Uint8Array;
 * sheet.addImage(workbook.addImage({ buffer: png, extension: 'png' }), 'B2:D6');
 * const corner = sheet.getImages()[0]?.range?.tl;
 * console.log(corner?.nativeCol, corner?.col); // 1 1
 * ```
 */
export class Anchor implements AnchorModel {
  nativeCol = 0;
  nativeColOff = 0;
  nativeRow = 0;
  nativeRowOff = 0;

  /** @internal */
  constructor(
    readonly worksheet: Worksheet | undefined,
    address?: string | Partial<AnchorModel> | { col: number; row: number },
    offset = 0,
  ) {
    if (typeof address === 'string') {
      const d = decode(address);
      this.nativeCol = d.col + offset;
      this.nativeRow = d.row + offset;
    } else if (address && 'nativeCol' in address && address.nativeCol !== undefined) {
      this.nativeCol = address.nativeCol || 0;
      this.nativeColOff = address.nativeColOff || 0;
      this.nativeRow = address.nativeRow || 0;
      this.nativeRowOff = address.nativeRowOff || 0;
    } else if (address && 'col' in address && address.col !== undefined) {
      this.col = address.col + offset;
      this.row = address.row + offset;
    }
  }

  get colWidth(): number {
    const column = this.worksheet?._columns[this.nativeCol];
    return column?.isCustomWidth && column.width !== undefined ? Math.floor(column.width * 10000) : 640000;
  }
  get rowHeight(): number {
    const height = this.worksheet?.findRow(this.nativeRow + 1)?.height;
    return height ? Math.floor(height * 10000) : 180000;
  }
  get col(): number {
    return this.nativeCol + Math.min(this.colWidth - 1, this.nativeColOff) / this.colWidth;
  }
  set col(v: number) {
    this.nativeCol = Math.floor(v);
    this.nativeColOff = Math.floor((v - this.nativeCol) * this.colWidth);
  }
  get row(): number {
    return this.nativeRow + Math.min(this.rowHeight - 1, this.nativeRowOff) / this.rowHeight;
  }
  set row(v: number) {
    this.nativeRow = Math.floor(v);
    this.nativeRowOff = Math.floor((v - this.nativeRow) * this.rowHeight);
  }
  get model(): AnchorModel {
    return { nativeCol: this.nativeCol, nativeColOff: this.nativeColOff, nativeRow: this.nativeRow, nativeRowOff: this.nativeRowOff };
  }
}

/**
 * Where a picture goes, as in exceljs `worksheet.addImage` (index.d.ts L915–928): a range such as
 * `"B2:D6"`, two corners `{ tl, br }`, or a corner and a size in pixels `{ tl, ext }`. Corners
 * are 0-based `{ col, row }` (fractions are parts of a cell) or exceljs `nativeCol` objects.
 *
 * @example
 * ```ts
 * import type { ImageRangeInput } from '@xlvane/exceljs';
 *
 * const cells: ImageRangeInput = 'B2:D6';
 * const sized: ImageRangeInput = { tl: { col: 1, row: 1 }, ext: { width: 160, height: 90 } };
 * ```
 */
export type ImageRangeInput =
  | string
  | {
      tl: string | Partial<AnchorModel> | { col: number; row: number };
      br?: string | Partial<AnchorModel> | { col: number; row: number };
      ext?: { width: number; height: number };
      editAs?: 'oneCell' | 'twoCell' | 'absolute';
      hyperlinks?: { hyperlink?: string; tooltip?: string };
    };

/** A picture or background of a worksheet (exceljs lib/doc/image.js). */
export interface SheetMedia {
  type: 'image' | 'background';
  imageId: number;
  range?: { tl: Anchor; br?: Anchor; ext?: { width: number; height: number }; editAs?: string; hyperlinks?: { hyperlink?: string; tooltip?: string } };
}

/** The range of `worksheet.addImage`, as exceljs `Image.model` sets it (image.js L33–54). */
export function imageRange(ws: Worksheet, range: ImageRangeInput): NonNullable<SheetMedia['range']> {
  if (typeof range === 'string') {
    const [a, b] = range.split(':') as [string, string | undefined];
    const tl = decode(a);
    const br = decode(b ?? a);
    return { tl: new Anchor(ws, { nativeCol: tl.col - 1, nativeRow: tl.row - 1 }), br: new Anchor(ws, { nativeCol: br.col, nativeRow: br.row }), editAs: 'oneCell' };
  }
  if (!range || typeof range !== 'object' || range.tl === undefined) {
    throw new XlvError('XLV_E_INVALID_OPTION', 'worksheet.addImage needs a range such as "B2:D6" or { tl, br } or { tl, ext }.', { hint: "For example sheet.addImage(id, 'B2:D6') or sheet.addImage(id, { tl: { col: 1, row: 1 }, ext: { width: 160, height: 90 } }).", details: { option: 'range' } });
  }
  return {
    tl: new Anchor(ws, range.tl),
    ...(range.br === undefined ? {} : { br: new Anchor(ws, range.br) }),
    ...(range.ext === undefined ? {} : { ext: { ...range.ext } }),
    ...(range.editAs === undefined ? {} : { editAs: range.editAs }),
    ...(range.hyperlinks === undefined ? {} : { hyperlinks: { ...range.hyperlinks } }),
  };
}

/**
 * The bytes of an image given to `workbook.addImage`, checked: PNG, JPEG or GIF by their first
 * bytes (core `imageFormat`, TD-31). A base64 string may be a data URL: everything up to the
 * first comma is skipped, as exceljs does (xlsx.js addMedia).
 */
export function imageBytes(image: ImageData): Uint8Array {
  if (!image || typeof image !== 'object') throw invalidImage('workbook.addImage needs { buffer } or { base64 }.');
  let bytes: Uint8Array | undefined;
  if (image.buffer !== undefined) {
    if (!(image.buffer instanceof Uint8Array)) throw invalidImage('The image "buffer" must be a Buffer or Uint8Array.');
    bytes = image.buffer;
  } else if (typeof image.base64 === 'string') {
    const text = image.base64.slice(image.base64.indexOf(',') + 1).replace(/\s+/g, '');
    let binary: string;
    try {
      binary = atob(text);
    } catch {
      throw invalidImage('The image "base64" is not valid base64.');
    }
    bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  } else if (image.filename !== undefined) {
    throw new XlvError('XLV_E_NOT_SUPPORTED', 'workbook.addImage({ filename }) is not supported: @xlvane/exceljs does not read files by name in every runtime.', {
      hint: "Read the file and pass its contents: workbook.addImage({ buffer: fs.readFileSync(filename), extension: 'png' }).",
      details: { feature: 'image filename' },
    });
  } else throw invalidImage('workbook.addImage needs { buffer } or { base64 }.');
  if (imageFormat(bytes) === null) throw invalidImage('The image is not a PNG, JPEG or GIF file.');
  return bytes;
}

function invalidImage(message: string): XlvError {
  return new XlvError('XLV_E_INVALID_IMAGE', message, { hint: "Pass the bytes of a PNG, JPEG or GIF file, e.g. workbook.addImage({ buffer: pngBytes, extension: 'png' }).", details: {} });
}

/**
 * The column width and row height xlvane uses for pictures (TD-33: `round(width × 7)` pixels per
 * custom column, 64 by default; `height × 96 / 72` pixels per custom row, 20 by default; 9525
 * EMU per pixel). An exceljs offset in EMU is given to the core as a fraction of these sizes, so
 * that the core writes the same offset exceljs would write.
 */
function cellEmu(ws: Worksheet, axis: 'col' | 'row', index: number): number {
  if (axis === 'col') {
    const column = ws._columns[index];
    const px = column?.width !== undefined && column.isCustomWidth ? Math.round(column.width * 7) : 64;
    return Math.max(1, px) * 9525;
  }
  const height = ws.findRow(index + 1)?.height;
  return Math.max(1, height ? (height * 96) / 72 : 20) * 9525;
}

function position(ws: Worksheet, a: Anchor): { col: number; row: number } {
  return { col: a.nativeCol + a.nativeColOff / cellEmu(ws, 'col', a.nativeCol), row: a.nativeRow + a.nativeRowOff / cellEmu(ws, 'row', a.nativeRow) };
}

/** The core anchor and options of a picture (D-D03): two corners, or a corner and a size. */
export function imageAnchorIn(
  ws: Worksheet,
  range: NonNullable<SheetMedia['range']>,
): { anchor: { from: { col: number; row: number }; to: { col: number; row: number }; editAs: 'twoCell' | 'oneCell' | 'absolute' } | { from: { col: number; row: number }; size: { width: number; height: number } }; hyperlink?: string } {
  const from = position(ws, range.tl);
  const link = range.hyperlinks?.hyperlink;
  const options = link ? { hyperlink: link } : {};
  if (range.br) {
    // two-cell-anchor-xform.js L27–36: exceljs writes editAs="oneCell" when none is given.
    const editAs = range.editAs === 'twoCell' || range.editAs === 'absolute' ? range.editAs : 'oneCell';
    return { anchor: { from, to: position(ws, range.br), editAs }, ...options };
  }
  if (!range.ext) {
    // exceljs throws a TypeError here at write time (ext-xform.js reads ext.width).
    throw new XlvError('XLV_E_INVALID_OPTION', 'A picture with only a top-left corner needs "ext" (its size in pixels).', { hint: 'For example { tl: { col: 1, row: 1 }, ext: { width: 160, height: 90 } }.', details: { option: 'ext' } });
  }
  return { anchor: { from, size: { width: range.ext.width, height: range.ext.height } }, ...options };
}

/** A picture read from a file, in exceljs's shape (`getImages()`); `null` for absolute anchors, which exceljs drops. */
export function imageOut(ws: Worksheet, image: SheetImage, imageId: number): SheetMedia | null {
  const a = image.anchor;
  const corner = (m: { col: number; row: number; colOffset: number; rowOffset: number }): Anchor =>
    new Anchor(ws, { nativeCol: m.col, nativeColOff: m.colOffset, nativeRow: m.row, nativeRowOff: m.rowOffset });
  const hyperlinks = image.hyperlink === null ? {} : { hyperlinks: { hyperlink: image.hyperlink } };
  if (a.type === 'twoCell') return { type: 'image', imageId, range: { tl: corner(a.from), br: corner(a.to), editAs: a.editAs, ...hyperlinks } };
  if (a.type === 'oneCell') {
    // ext-xform.js: exceljs reads the size back in pixels (EMU / 9525) and editAs as "oneCell".
    return { type: 'image', imageId, range: { tl: corner(a.from), ext: { width: a.size.width / 9525, height: a.size.height / 9525 }, editAs: 'oneCell', ...hyperlinks } };
  }
  return null;
}

/** The media list of a workbook read from a file, as exceljs reads it (xlsx.js _processMediaEntry). */
export function mediaOut(images: readonly WorkbookImage[]): ImageData[] {
  return images.map((i) => {
    const file = i.part.slice(i.part.lastIndexOf('/') + 1);
    const dot = file.lastIndexOf('.');
    return { type: 'image', name: dot === -1 ? file : file.slice(0, dot), extension: dot === -1 ? '' : file.slice(dot + 1), buffer: i.bytes };
  });
}
