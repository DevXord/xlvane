// Property tests for CSV: toCSV and the CSV parser are inverse, and hostile text never throws
// anything but XlvError.
import fc from 'fast-check';
import { describe, expect, it } from 'vitest';
import { parseCsvRows, toCSV } from '../../packages/xlvane/src/csv';
import { isXlvError } from '../../packages/xlvane/src/errors';
import { formatCellRef } from '../../packages/xlvane/src/sml/cell-ref';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { firstLoneSurrogate } from '../../packages/xlvane/src/xml/writer';
import { NUM_RUNS } from './runs';

const RUNS = { numRuns: Math.max(1, Math.floor(NUM_RUNS / 10)) };
const piece = fc.constantFrom('a', ',', ';', '"', '\r', '\n', '\r\n', ' ', 'ż', '😀', '=', "'", '\t', 'x"y');
const text = fc
  .oneof(fc.string({ maxLength: 12 }), fc.array(piece, { minLength: 1, maxLength: 6 }).map((a) => a.join('')))
  .filter((s) => s !== '' && firstLoneSurrogate(s) === -1);

describe('CSV properties', () => {
  it('parses what toCSV writes back to the same text fields', () => {
    fc.assert(
      fc.property(fc.array(fc.array(text, { minLength: 1, maxLength: 4 }), { minLength: 1, maxLength: 5 }), fc.constantFrom(',', ';', '\t'), (grid, delimiter) => {
        const wb = createWorkbook();
        const s = wb.addSheet('S');
        const width = Math.max(...grid.map((r) => r.length));
        grid.forEach((row, r) => row.forEach((v, c) => s.set(formatCellRef(r + 1, c + 1), v)));
        const csv = toCSV(writeWorkbook(wb), { delimiter, formulaInjection: 'off' });
        const rows = parseCsvRows(csv, delimiter);
        expect(rows).toEqual(grid.map((row) => [...row, ...Array<string>(width - row.length).fill('')]));
      }),
      RUNS,
    );
  });

  it('never throws anything but XlvError on arbitrary text', () => {
    fc.assert(
      fc.property(fc.oneof(fc.string({ maxLength: 200 }), fc.array(piece, { maxLength: 60 }).map((a) => a.join(''))), (input) => {
        try {
          parseCsvRows(input, ',');
        } catch (error) {
          expect(isXlvError(error)).toBe(true);
        }
      }),
      { numRuns: NUM_RUNS },
    );
  });
});
