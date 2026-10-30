// exceljs `DefinedNames` (lib/doc/defined-names.js) and the cell matrix behind it
// (lib/utils/cell-matrix.js): names are sets of cells, grouped into ranges when written
// (docs/facts/exceljs.md X4-05 … X4-08).
import { columnLetter } from './model';

/**
 * A cell location of a defined name, as exceljs keeps it.
 *
 * @example
 * ```ts
 * import ExcelJS, { type NameCell } from '@xlvane/exceljs';
 *
 * const workbook = new ExcelJS.Workbook();
 * workbook.definedNames.add('Data!$B$2', 'total');
 * workbook.definedNames.forEach((name: string, cell: NameCell) => console.log(name, cell.sheetName, cell.address)); // total Data B2
 * ```
 */
export interface NameCell {
  sheetName: string | undefined;
  address: string;
  row: number;
  col: number;
  /** @internal Used while grouping cells into ranges. */
  mark?: boolean;
}

/**
 * A location parsed by exceljs `colCache.decodeEx`: one cell, a range, or an error reference.
 *
 * @example
 * ```ts
 * import ExcelJS, { type DecodedLocation } from '@xlvane/exceljs';
 *
 * const workbook = new ExcelJS.Workbook();
 * const location: DecodedLocation = { sheetName: 'Data', row: 2, col: 2 };
 * workbook.definedNames.addEx(location, 'total');
 * console.log(workbook.definedNames.getNamesEx(location)); // ['total']
 * ```
 */
export interface DecodedLocation {
  sheetName?: string | undefined;
  row?: number | undefined;
  col?: number | undefined;
  top?: number;
  left?: number;
  bottom?: number;
  right?: number;
  error?: string;
}

/** exceljs `colCache.decodeAddress` (lib/utils/col-cache.js L122): lenient, throws only past XFD. */
function decodeAddressLoose(value: string): { row: number | undefined; col: number | undefined } {
  let hasCol = false;
  let colNumber = 0;
  let hasRow = false;
  let rowNumber = 0;
  for (let i = 0; i < value.length; i++) {
    const char = value.charCodeAt(i);
    if (!hasRow && char >= 65 && char <= 90) {
      hasCol = true;
      colNumber = colNumber * 26 + char - 64;
    } else if (char >= 48 && char <= 57) {
      hasRow = true;
      rowNumber = rowNumber * 10 + char - 48;
    } else if (hasRow && hasCol && char !== 36) break;
  }
  if (hasCol && colNumber > 16_384) throw new Error(`Out of bounds. Invalid column letter: ${value}`);
  return { col: hasCol ? colNumber : undefined, row: hasRow ? rowNumber : undefined };
}

/** exceljs `colCache.decodeEx` (lib/utils/col-cache.js L216). */
export function decodeEx(value: string): DecodedLocation {
  const groups = /(?:(?:(?:'((?:[^']|'')*)')|([^'^ !]*))!)?(.*)/.exec(value) ?? [];
  const sheetName = groups[1] || groups[2] || undefined;
  const reference = groups[3] ?? '';
  const parts = reference.split(':');
  if (parts.length > 1) {
    const tl = decodeAddressLoose(parts[0] ?? '');
    const br = decodeAddressLoose(parts[1] ?? '');
    return {
      sheetName,
      top: Math.min(tl.row ?? Number.NaN, br.row ?? Number.NaN),
      left: Math.min(tl.col ?? Number.NaN, br.col ?? Number.NaN),
      bottom: Math.max(tl.row ?? Number.NaN, br.row ?? Number.NaN),
      right: Math.max(tl.col ?? Number.NaN, br.col ?? Number.NaN),
    };
  }
  if (reference.startsWith('#')) return sheetName ? { sheetName, error: reference } : { error: reference };
  const address = decodeAddressLoose(reference);
  return sheetName ? { sheetName, ...address } : address;
}

/** The sheet prefix exceljs writes (lib/doc/range.js L153): quoted unless letters and digits only. */
function sheetPrefix(sheetName: string | undefined): string {
  if (!sheetName) return '';
  return /^[a-zA-Z0-9]*$/.test(sheetName) ? `${sheetName}!` : `'${sheetName}'!`;
}

/**
 * The cells of one defined name, as exceljs `CellMatrix`: sheet → row → column.
 *
 * @example
 * ```ts
 * import ExcelJS from '@xlvane/exceljs';
 *
 * const workbook = new ExcelJS.Workbook();
 * workbook.definedNames.add('Data!$A$1:$A$3', 'items');
 * workbook.definedNames.getMatrix('items').forEach((cell) => console.log(cell.address)); // A1 A2 A3
 * ```
 */
export class CellMatrix {
  /** @internal */
  readonly sheets: Map<string | undefined, (NameCell | undefined)[][]> = new Map();

  addCell(addressStr: string): void {
    this.addCellEx(decodeEx(addressStr));
  }

  getCell(addressStr: string): NameCell | undefined {
    return this.findCellEx(decodeEx(addressStr), true);
  }

  findCell(addressStr: string): NameCell | undefined {
    return this.findCellEx(decodeEx(addressStr), false);
  }

  findCellAt(sheetName: string | undefined, rowNumber: number, colNumber: number): NameCell | undefined {
    return this.sheets.get(sheetName)?.[rowNumber]?.[colNumber];
  }

  addCellEx(address: DecodedLocation): void {
    if (address.top) {
      for (let row = address.top; row <= (address.bottom ?? 0); row++) {
        for (let col = address.left ?? 0; col <= (address.right ?? 0); col++) this.getCellAt(address.sheetName, row, col);
      }
    } else this.findCellEx(address, true);
  }

  getCellEx(address: DecodedLocation): NameCell | undefined {
    return this.findCellEx(address, true);
  }

  findCellEx(address: DecodedLocation, create: boolean): NameCell | undefined {
    const { row, col } = address;
    if (row === undefined || col === undefined || !Number.isInteger(row) || !Number.isInteger(col)) return undefined;
    const existing = this.findCellAt(address.sheetName, row, col);
    if (existing || !create) return existing;
    return this.getCellAt(address.sheetName, row, col);
  }

  getCellAt(sheetName: string | undefined, rowNumber: number, colNumber: number): NameCell {
    let sheet = this.sheets.get(sheetName);
    if (!sheet) {
      sheet = [];
      this.sheets.set(sheetName, sheet);
    }
    const row = (sheet[rowNumber] ??= []);
    return (row[colNumber] ??= { sheetName, address: `${columnLetter(colNumber)}${String(rowNumber)}`, row: rowNumber, col: colNumber });
  }

  removeCellEx(address: DecodedLocation): void {
    const row = address.row === undefined ? undefined : this.sheets.get(address.sheetName)?.[address.row];
    if (row && address.col !== undefined) row[address.col] = undefined;
  }

  forEachInSheet(sheetName: string | undefined, callback: (cell: NameCell, row: number, col: number) => void): void {
    this.sheets.get(sheetName)?.forEach((row, r) => {
      row?.forEach((cell, c) => {
        if (cell) callback(cell, r, c);
      });
    });
  }

  forEach(callback: (cell: NameCell) => void): void {
    for (const name of this.sheets.keys()) this.forEachInSheet(name, (cell) => callback(cell));
  }

  /** Removes `numDelete` rows at `start` and inserts `numInsert` empty ones (cell-matrix.js L138). */
  spliceRows(sheetName: string | undefined, start: number, numDelete: number, numInsert: number): void {
    const sheet = this.sheets.get(sheetName);
    if (sheet) sheet.splice(start, numDelete, ...Array.from({ length: numInsert }, () => []));
  }

  /** Removes `numDelete` columns at `start` and inserts `numInsert` empty ones (cell-matrix.js L150). */
  spliceColumns(sheetName: string | undefined, start: number, numDelete: number, numInsert: number): void {
    const sheet = this.sheets.get(sheetName);
    if (sheet) {
      for (const row of sheet) if (row) row.splice(start, numDelete, ...Array.from({ length: numInsert }, () => undefined));
    }
  }

  map<T>(callback: (cell: NameCell) => T): T[] {
    const out: T[] = [];
    this.forEach((cell) => out.push(callback(cell)));
    return out;
  }
}

/** exceljs `rangeRegexp` (lib/doc/defined-names.js L8): only absolute references are kept on load. */
const RANGE = /[$](\w+)[$](\d+)(:[$](\w+)[$](\d+))?/;

/**
 * Defined names of a workbook, as exceljs `workbook.definedNames`. A name is a set of cells; when
 * the workbook is written, the cells are grouped into rectangles with absolute references
 * (`Sheet!$A$1:$B$2`), and names that hold no cells are left out, as in exceljs.
 *
 * @example
 * ```ts
 * import ExcelJS from '@xlvane/exceljs';
 *
 * const workbook = new ExcelJS.Workbook();
 * workbook.addWorksheet('Data');
 * workbook.definedNames.add('Data!$A$1:$B$2', 'block');
 * console.log(workbook.definedNames.getRanges('block')); // { name: 'block', ranges: ['Data!$A$1:$B$2'] }
 * console.log(workbook.definedNames.getNames('Data!B2')); // ['block']
 * ```
 */
export class DefinedNames {
  /** @internal */
  matrixMap: Map<string, CellMatrix> = new Map();

  getMatrix(name: string): CellMatrix {
    let matrix = this.matrixMap.get(name);
    if (!matrix) {
      matrix = new CellMatrix();
      this.matrixMap.set(name, matrix);
    }
    return matrix;
  }

  /** Adds the cells of `locStr` (`Sheet!$A$1` or `Sheet!$A$1:$B$2`) to the name. */
  add(locStr: string, name: string): void {
    this.addEx(decodeEx(locStr), name);
  }

  addEx(location: DecodedLocation, name: string): void {
    const matrix = this.getMatrix(name);
    if (location.top) {
      for (let col = location.left ?? 0; col <= (location.right ?? 0); col++) {
        for (let row = location.top; row <= (location.bottom ?? 0); row++) matrix.getCellAt(location.sheetName, row, col);
      }
    } else matrix.addCellEx(location);
  }

  remove(locStr: string, name: string): void {
    this.removeEx(decodeEx(locStr), name);
  }

  removeEx(location: DecodedLocation, name: string): void {
    this.getMatrix(name).removeCellEx(location);
  }

  removeAllNames(location: DecodedLocation): void {
    for (const matrix of this.matrixMap.values()) matrix.removeCellEx(location);
  }

  forEach(callback: (name: string, cell: NameCell) => void): void {
    for (const [name, matrix] of this.matrixMap) matrix.forEach((cell) => callback(name, cell));
  }

  /** The names that include the cell `addressStr` (`Sheet!A1`). */
  getNames(addressStr: string): string[] {
    return this.getNamesEx(decodeEx(addressStr));
  }

  getNamesEx(address: DecodedLocation): string[] {
    const out: string[] = [];
    for (const [name, matrix] of this.matrixMap) if (matrix.findCellEx(address, false)) out.push(name);
    return out;
  }

  private explore(matrix: CellMatrix, cell: NameCell): { top: number; left: number; bottom: number; right: number; sheetName: string | undefined } {
    cell.mark = false;
    const { sheetName } = cell;
    const range = { top: cell.row, left: cell.col, bottom: cell.row, right: cell.col, sheetName };
    const vGrow = (yy: number, edge: 'top' | 'bottom'): boolean => {
      const c = matrix.findCellAt(sheetName, yy, cell.col);
      if (!c || !c.mark) return false;
      range[edge] = yy;
      c.mark = false;
      return true;
    };
    for (let y = cell.row - 1; vGrow(y, 'top'); y--);
    for (let y = cell.row + 1; vGrow(y, 'bottom'); y++);
    const hGrow = (xx: number, edge: 'left' | 'right'): boolean => {
      const cells: NameCell[] = [];
      for (let y = range.top; y <= range.bottom; y++) {
        const c = matrix.findCellAt(sheetName, y, xx);
        if (c && c.mark) cells.push(c);
        else return false;
      }
      range[edge] = xx;
      for (const c of cells) c.mark = false;
      return true;
    };
    for (let x = cell.col - 1; hGrow(x, 'left'); x--);
    for (let x = cell.col + 1; hGrow(x, 'right'); x++);
    return range;
  }

  /** Gives moved cells their new row and column (lib/doc/defined-names.js L147). */
  normaliseMatrix(matrix: CellMatrix, sheetName: string | undefined): void {
    matrix.forEachInSheet(sheetName, (cell, row, col) => {
      if (cell.row !== row || cell.col !== col) {
        cell.row = row;
        cell.col = col;
        cell.address = `${columnLetter(col)}${String(row)}`;
      }
    });
  }

  spliceRows(sheetName: string | undefined, start: number, numDelete: number, numInsert: number): void {
    for (const matrix of this.matrixMap.values()) {
      matrix.spliceRows(sheetName, start, numDelete, numInsert);
      this.normaliseMatrix(matrix, sheetName);
    }
  }

  spliceColumns(sheetName: string | undefined, start: number, numDelete: number, numInsert: number): void {
    for (const matrix of this.matrixMap.values()) {
      matrix.spliceColumns(sheetName, start, numDelete, numInsert);
      this.normaliseMatrix(matrix, sheetName);
    }
  }

  /**
   * xlvane addition: after `numInsert` rows (or columns) were inserted at `start`, adds the
   * inserted cells to every name that has cells on both sides of them in the same column (or row),
   * so a named range grows as in a spreadsheet application. exceljs leaves a gap (X4-15).
   */
  fillInserted(sheetName: string | undefined, axis: 'row' | 'col', start: number, numInsert: number): void {
    if (numInsert <= 0) return;
    for (const matrix of this.matrixMap.values()) {
      const lines = new Set<number>();
      matrix.forEachInSheet(sheetName, (_cell, row, col) => {
        if ((axis === 'row' ? row : col) === start - 1) lines.add(axis === 'row' ? col : row);
      });
      for (const line of lines) {
        const after = axis === 'row' ? matrix.findCellAt(sheetName, start + numInsert, line) : matrix.findCellAt(sheetName, line, start + numInsert);
        if (!after) continue;
        for (let i = start; i < start + numInsert; i++) {
          if (axis === 'row') matrix.getCellAt(sheetName, i, line);
          else matrix.getCellAt(sheetName, line, i);
        }
      }
    }
  }

  /** The cells of a name grouped into ranges, as exceljs `getRanges` (lib/doc/defined-names.js L125). */
  getRanges(name: string, matrix?: CellMatrix): { name: string; ranges: string[] } {
    const m = matrix ?? this.matrixMap.get(name);
    if (!m) return { name, ranges: [] };
    m.forEach((cell) => {
      cell.mark = true;
    });
    const ranges = m
      .map((cell) => (cell.mark ? this.explore(m, cell) : null))
      .filter((r): r is NonNullable<typeof r> => r !== null)
      .map((r) => {
        const tl = `$${columnLetter(r.left)}$${String(r.top)}`;
        const single = r.top === r.bottom && r.left === r.right;
        return sheetPrefix(r.sheetName) + (single ? tl : `${tl}:$${columnLetter(r.right)}$${String(r.bottom)}`);
      });
    return { name, ranges };
  }

  /** Every name with at least one cell, as `{ name, ranges }`. */
  get model(): { name: string; ranges: string[] }[] {
    return [...this.matrixMap].map(([name, matrix]) => this.getRanges(name, matrix)).filter((d) => d.ranges.length > 0);
  }

  set model(value: { name: string; ranges: string[] }[]) {
    this.matrixMap = new Map();
    for (const definedName of value) {
      const matrix = new CellMatrix();
      this.matrixMap.set(definedName.name, matrix);
      for (const rangeStr of definedName.ranges) if (RANGE.test(rangeStr.split('!').pop() ?? '')) matrix.addCell(rangeStr);
    }
  }
}

/** exceljs `extractRanges` (lib/xlsx/xform/book/defined-name-xform.js L251): the ranges of a name's formula text. */
export function extractRanges(parsedText: string): string[] {
  const valid = (range: string): boolean => {
    try {
      decodeEx(range);
      return true;
    } catch {
      return false;
    }
  };
  const ranges: string[] = [];
  let quotesOpened = false;
  let last = '';
  for (const item of parsedText.split(',')) {
    if (!item) continue;
    const quotes = (item.match(/'/g) ?? []).length;
    if (!quotes) {
      if (quotesOpened) last += `${item},`;
      else if (valid(item)) ranges.push(item);
      continue;
    }
    const quotesEven = quotes % 2 === 0;
    if (!quotesOpened && quotesEven && valid(item)) ranges.push(item);
    else if (quotesOpened && !quotesEven) {
      quotesOpened = false;
      if (valid(last + item)) ranges.push(last + item);
      last = '';
    } else {
      quotesOpened = true;
      last += `${item},`;
    }
  }
  return ranges;
}
