// editCells and appendRows after the K3 review (A2): a 4.4 MB file with 4.9 million number cells,
// below the default maxCells, made editCells run out of memory (about 0.8 KB per cell of the edited
// sheet), and appendRows with 4 000 rows of 1 200 values did the same while it planned the cells.
// The cells of the edited sheets plus the new cells are now counted against the edit budget
// (`limits.maxCells`, at most EDIT_MAX_CELLS by default) before anything per cell is built, and the
// edit keeps a few numbers per existing cell instead of objects.
import { strToU8, zipSync } from 'fflate';
import { describe, expect, it, vi } from 'vitest';
import { appendRows, editCells, EDIT_MAX_CELLS, type CellEdit } from '../../packages/xlvane/src/edit';
import { DEFAULT_READ_LIMITS } from '../../packages/xlvane/src/limits';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { formatCellRef } from '../../packages/xlvane/src/sml/cell-ref';
import { NewSheet, type CellInput } from '../../packages/xlvane/src/write';
import { expectXlvError } from '../helpers/expect-error';

const M = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** One sheet "S" of `rows` × `cols` number cells, written without cell references as some writers do. */
function grid(rows: number, cols: number, refs = false): Uint8Array {
  const parts: string[] = [];
  for (let r = 1; r <= rows; r++) {
    parts.push(refs ? `<row r="${String(r)}">` : '<row>');
    for (let c = 1; c <= cols; c++) parts.push(refs ? `<c r="${formatCellRef(r, c)}"><v>${String(c)}</v></c>` : `<c><v>${String(c)}</v></c>`);
    parts.push('</row>');
  }
  return zipSync({
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="${M}" xmlns:r="${R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="${M}"><sheetData>${parts.join('')}</sheetData></worksheet>`),
  });
}

/** Runs `fn` and returns how many cells the edit planned (NewSheet.set calls of planEdits). */
function plannedCells(fn: () => unknown): number {
  const set = vi.spyOn(NewSheet.prototype, 'set');
  try {
    try {
      fn();
    } catch {
      // The caller checks the error separately.
    }
    return set.mock.calls.length;
  } finally {
    set.mockRestore();
  }
}

describe('A2: the cells of an edit are counted before the edit is built', () => {
  const bytes = grid(30, 30);
  const newCells: CellEdit[] = Array.from({ length: 200 }, (_, i) => ({ sheet: 'S', cell: formatCellRef(31 + Math.floor(i / 30), 1 + (i % 30)), value: i }));

  it('editCells refuses new cells beyond maxCells, before planning them', () => {
    const error = expectXlvError(() => editCells(bytes, newCells, { limits: { maxCells: 1000 } }), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxCells', value: 1000 });
    expect(error.location.sheet).toBe('S');
    expect(plannedCells(() => editCells(bytes, newCells, { limits: { maxCells: 1000 } }))).toBe(0);
  });

  it('editCells counts only cells that are new: replacing cells of a sheet at the limit works', () => {
    const replace: CellEdit[] = Array.from({ length: 30 }, (_, i) => ({ sheet: 'S', cell: formatCellRef(1 + i, 2), value: -i }));
    const { bytes: out } = editCells(bytes, replace, { limits: { maxCells: 900 } });
    const sheet = readWorkbook(out, { limits: { maxCells: 900 } }).getSheet('S');
    expect(sheet.cellCount).toBe(900);
    expect(sheet.cell('B30')?.value).toBe(-29);
    expect(sheet.cell('C30')?.value).toBe(3);
  });

  it('editCells with more edits than maxCells is refused before the file is read', () => {
    const many: CellEdit[] = Array.from({ length: 11 }, (_, i) => ({ sheet: 'Missing', cell: `A${String(i + 1)}`, value: i }));
    expectXlvError(() => editCells(bytes, many, { limits: { maxCells: 10 } }), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxCells', value: 10 });
  });

  it('appendRows refuses rows beyond maxCells, before planning them', () => {
    const rows = Array.from({ length: 10 }, (_, r) => Array.from({ length: 30 }, (_, c) => r * 30 + c));
    expectXlvError(() => appendRows(bytes, 'S', rows, { limits: { maxCells: 1000 } }), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxCells', value: 1000 });
    expect(plannedCells(() => appendRows(bytes, 'S', rows, { limits: { maxCells: 1000 } }))).toBe(0);
    // Within the limit the rows are appended.
    const { bytes: out } = appendRows(bytes, 'S', rows, { limits: { maxCells: 1200 } });
    expect(readWorkbook(out).getSheet('S').cellCount).toBe(1200);
  });

  it('appendRows counts the values, not the rows: empty values are not cells', () => {
    // undefined is skipped like null; the type does not allow it, but JavaScript callers can pass it.
    const sparse = Array.from({ length: 200 }, () => [1, null, undefined] as unknown as CellInput[]);
    expect(readWorkbook(appendRows(bytes, 'S', sparse, { limits: { maxCells: 1100 } }).bytes).getSheet('S').cellCount).toBe(1100);
  });

  it('appendRows with millions of values is refused by the default edit budget instead of running out of memory', () => {
    const row = Array.from({ length: 1200 }, (_, i) => i);
    const rows = Array.from({ length: 4000 }, () => row);
    expect(EDIT_MAX_CELLS).toBeLessThan(4000 * 1200);
    expect(EDIT_MAX_CELLS).toBeLessThanOrEqual(DEFAULT_READ_LIMITS.maxCells);
    const error = expectXlvError(() => appendRows(bytes, 'S', rows), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxCells', value: EDIT_MAX_CELLS });
    expect(error.hint).toContain('maxCells');
    expect(plannedCells(() => appendRows(bytes, 'S', rows))).toBe(0);
  }, 60_000);

  it('a maxCells given by the caller replaces the edit budget', () => {
    const rows = [[1, 2, 3]];
    expect(readWorkbook(appendRows(bytes, 'S', rows, { limits: { maxCells: EDIT_MAX_CELLS + 1 } }).bytes).getSheet('S').cellCount).toBe(903);
  });

  it('edits sheets without cell references and keeps every cell', () => {
    const { bytes: out } = editCells(bytes, [
      { sheet: 'S', cell: 'B2', value: 'x' },
      { sheet: 'S', cell: 'AE1', value: 7 },
    ]);
    const sheet = readWorkbook(out).getSheet('S');
    expect(sheet.cellCount).toBe(901);
    expect([sheet.cell('A1')?.value, sheet.cell('B2')?.value, sheet.cell('C2')?.value, sheet.cell('AD30')?.value, sheet.cell('AE1')?.value]).toEqual([1, 'x', 3, 30, 7]);
  });

  it('edits sheets with cell references and copies unchanged rows as they were', () => {
    const withRefs = grid(5, 3, true);
    const { bytes: out } = editCells(withRefs, [{ sheet: 'S', cell: 'B3', value: 9 }]);
    const sheet = readWorkbook(out).getSheet('S');
    expect([...sheet.rows()].map((r) => r.map((c) => c.value))).toEqual([
      [1, 2, 3],
      [1, 2, 3],
      [1, 9, 3],
      [1, 2, 3],
      [1, 2, 3],
    ]);
  });
});
