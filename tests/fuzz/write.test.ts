// Property tests for the writer: any workbook built through the public API is written to a
// package that our strict reader accepts and that reads back with the same content.
import fc from 'fast-check';
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { formatCellRef } from '../../packages/xlvane/src/sml/cell-ref';
import { createWorkbook, writeWorkbook, type CellInput } from '../../packages/xlvane/src/write';
import { firstLoneSurrogate } from '../../packages/xlvane/src/xml/writer';
import { NUM_RUNS } from './runs';

const RUNS = { numRuns: Math.max(1, Math.floor(NUM_RUNS / 20)) };

const tricky = fc.constantFrom('_x0041_', '_x005F_', '\u0001', '\r', '\r\n', '\t', ' ', '<', '&', ']]>', '😀', '￿', 'ż');
const text = fc
  .oneof(fc.string({ unit: 'binary', maxLength: 30 }), fc.array(tricky, { maxLength: 8 }).map((a) => a.join('')))
  .filter((s) => firstLoneSurrogate(s) === -1);
const number = fc.double({ noNaN: true, noDefaultInfinity: true });
const date = fc.date({ min: new Date(Date.UTC(1900, 0, 1)), max: new Date(Date.UTC(9999, 11, 31)), noInvalidDate: true });
const value: fc.Arbitrary<CellInput> = fc.oneof(text, number, fc.boolean(), date);

const sheetName = fc
  .string({ minLength: 1, maxLength: 31, unit: fc.constantFrom('a', 'B', ' ', '_', '1', 'ż', '😀', '&', '<', '-') })
  .filter((s) => s.length <= 31);

const cells = fc.array(
  fc.record({ row: fc.integer({ min: 1, max: 60 }), col: fc.integer({ min: 1, max: 30 }), value }),
  { maxLength: 40 },
);

function expected(v: CellInput): unknown {
  if (typeof v === 'number') return Object.is(v, -0) ? 0 : v;
  if (v instanceof Date) {
    // Stored with millisecond precision in a double serial; compare to the millisecond.
    return v.getTime();
  }
  return v;
}

describe('writer properties', () => {
  it('reads back every value written', () => {
    fc.assert(
      fc.property(fc.uniqueArray(sheetName, { minLength: 1, maxLength: 3, selector: (s) => s.toLowerCase() }), cells, (names, list) => {
        const wb = createWorkbook();
        for (const name of names) wb.addSheet(name);
        const sheet = wb.sheets[0]!;
        const last = new Map<string, CellInput>();
        for (const { row, col, value: v } of list) {
          const ref = formatCellRef(row, col);
          // An explicit empty style replaces the cell's style; without it a later number would keep
          // an earlier date's number format (set() keeps the current style, as Excel does).
          sheet.set(ref, v, {});
          last.set(ref, v);
        }
        const read = readWorkbook(writeWorkbook(wb));
        expect(read.sheets.map((s) => s.name)).toEqual(names);
        const rs = read.getSheet(names[0]!);
        for (const [ref, v] of last) {
          const got = rs.cell(ref)?.value;
          if (v instanceof Date) expect(Math.abs((got as Date).getTime() - v.getTime())).toBeLessThanOrEqual(1);
          else expect(got).toEqual(expected(v));
        }
        expect(rs.cellCount).toBe(last.size);
      }),
      RUNS,
    );
  });
});
