// Cells holding only a style are kept in a compact list per row. Finding a cell in it searched
// the list whenever the column was not the largest so far, so a row written right to left took
// quadratic time, and a value replacing a style-only cell spliced the list (1.1.1 security
// review, D-H01). Rows out of order now get an index by column, and replaced entries are marked.
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { formatCellRef } from '../../packages/xlvane/src/sml/cell-ref';
import { buildXlsx } from '../helpers/xlsx-builder';

const styles = '<fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellXfs count="2"><xf/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="1"/></cellXfs>';

describe('style-only cells', () => {
  it('reads 32 rows of 16 384 style-only cells written right to left in linear time', () => {
    let data = '';
    for (let r = 1; r <= 32; r++) {
      data += `<row r="${String(r)}">`;
      for (let c = 16384; c >= 1; c--) data += `<c r="${formatCellRef(r, c)}" s="1"/>`;
      data += '</row>';
    }
    const bytes = buildXlsx({ sheets: [{ name: 'S', data }], styles });
    const start = performance.now();
    const wb = readWorkbook(bytes);
    expect(performance.now() - start).toBeLessThan(3000);
    const cells = [...(wb.sheets[0]?.styledCells() ?? [])];
    expect(cells).toHaveLength(32 * 16384);
    expect(cells[0]?.ref).toBe('A1');
    expect(cells[16383]?.ref).toBe('XFD1');
  }, 60_000);

  it('values replacing style-only cells of a long row in reverse order take linear time', () => {
    let data = '<row r="1">';
    for (let c = 1; c <= 16384; c++) data += `<c r="${formatCellRef(1, c)}" s="1"/>`;
    for (let c = 16384; c >= 1; c--) data += `<c r="${formatCellRef(1, c)}"><v>${String(c)}</v></c>`;
    data += '</row>';
    const bytes = buildXlsx({ sheets: [{ name: 'S', data }], styles });
    const start = performance.now();
    const wb = readWorkbook(bytes);
    expect(performance.now() - start).toBeLessThan(5000);
    const cells = [...(wb.sheets[0]?.cells() ?? [])];
    expect(cells).toHaveLength(16384);
    expect(cells.every((c) => c.type === 'number')).toBe(true);
    expect([...(wb.sheets[0]?.styledCells() ?? [])]).toHaveLength(0);
  }, 60_000);

  it('keeps the last definition when style-only and valued cells alternate out of order', () => {
    const data = '<row r="1"><c r="C1" s="1"/><c r="A1" s="1"/><c r="C1"><v>3</v></c><c r="B1" s="1"/><c r="C1" s="1"/><c r="A1"><v>1</v></c><c r="A1" s="1"/></row>';
    const wb = readWorkbook(buildXlsx({ sheets: [{ name: 'S', data }], styles }));
    expect([...(wb.sheets[0]?.cells() ?? [])]).toHaveLength(0);
    expect([...(wb.sheets[0]?.styledCells() ?? [])].map((c) => c.ref)).toEqual(['A1', 'B1', 'C1']);
  });
});
