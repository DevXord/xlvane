// Property tests for cell references, serial dates, ISO 8601 date cells and number-format
// classification: hostile strings never throw, and every conversion round-trips.
import fc from 'fast-check';
import { describe, expect, it } from 'vitest';
import type { Conformance, DateSystem } from '../../packages/xlvane/src/model';
import {
  columnLetters,
  columnNumber,
  formatCellRef,
  formatRange,
  MAX_COLUMNS,
  MAX_ROWS,
  parseCellRef,
  parseRange,
} from '../../packages/xlvane/src/sml/cell-ref';
import { dateToSerial, parseIsoDateTime, serialToDate } from '../../packages/xlvane/src/sml/dates';
import { isDateFormat } from '../../packages/xlvane/src/sml/numfmt';
import { NUM_RUNS } from './runs';

const RUNS = { numRuns: NUM_RUNS };
const DAY_MS = 86_400_000;

const row = fc.integer({ min: 1, max: MAX_ROWS });
const col = fc.integer({ min: 1, max: MAX_COLUMNS });

/** Strings that look like cell references often enough to reach the interesting branches. */
const refLike = fc.string({
  unit: fc.constantFrom('A', 'B', 'X', 'F', 'D', 'Z', 'a', 'z', '0', '1', '9', '$', ':', ' ', '-', '.', 'Ä', '１'),
  maxLength: 24,
});
const anyString = fc.oneof(fc.string({ unit: 'binary', maxLength: 40 }), refLike);

const systems: readonly (readonly [DateSystem, Conformance])[] = [
  ['1900', 'strict'],
  ['1900', 'transitional'],
  ['1904', 'strict'],
  ['1904', 'transitional'],
];
const systemArb = fc.constantFrom(...systems);

const MIN_DATE = new Date(Date.UTC(2000, 0, 1));
MIN_DATE.setUTCFullYear(1);
const MAX_DATE = new Date(Date.UTC(9999, 11, 31, 23, 59, 59, 999));
// Serial 0 (1899-12-31) is the earliest value: time-only values lie on that day (D-112).
const TRANSITIONAL_1900_MIN = new Date(Date.UTC(1899, 11, 31));

/** Lowest valid instant of a date system. */
function lowest(system: DateSystem, conformance: Conformance): Date {
  return system === '1900' && conformance === 'transitional' ? TRANSITIONAL_1900_MIN : MIN_DATE;
}

describe('cell references (property)', () => {
  it('columnLetters and columnNumber are inverse', () => {
    fc.assert(
      fc.property(col, (c) => {
        expect(columnNumber(columnLetters(c))).toBe(c);
        expect(columnNumber(columnLetters(c).toLowerCase())).toBe(c);
      }),
      RUNS,
    );
  });

  it('formatCellRef and parseCellRef are inverse', () => {
    fc.assert(
      fc.property(row, col, (r, c) => {
        const ref = formatCellRef(r, c);
        expect(parseCellRef(ref)).toEqual({ row: r, col: c });
        expect(parseCellRef(ref.toLowerCase())).toEqual({ row: r, col: c });
      }),
      RUNS,
    );
  });

  it('parseCellRef never throws and only returns addresses inside the grid', () => {
    fc.assert(
      fc.property(anyString, (s) => {
        const address = parseCellRef(s);
        if (address === null) return;
        expect(Number.isInteger(address.row) && address.row >= 1 && address.row <= MAX_ROWS).toBe(true);
        expect(Number.isInteger(address.col) && address.col >= 1 && address.col <= MAX_COLUMNS).toBe(true);
        expect(s).not.toContain('$');
        // Canonical form: upper-case letters, row without leading zeros.
        expect(parseCellRef(formatCellRef(address.row, address.col))).toEqual(address);
      }),
      RUNS,
    );
  });

  it('parseRange never throws and always returns a normalised range inside the grid', () => {
    fc.assert(
      fc.property(anyString, (s) => {
        const range = parseRange(s);
        if (range === null) return;
        expect(range.start.row).toBeLessThanOrEqual(range.end.row);
        expect(range.start.col).toBeLessThanOrEqual(range.end.col);
        expect(parseRange(formatRange(range))).toEqual(range);
      }),
      RUNS,
    );
  });

  it('parseRange of two formatted corners gives their bounding box', () => {
    fc.assert(
      fc.property(row, col, row, col, (r1, c1, r2, c2) => {
        const range = parseRange(`${formatCellRef(r1, c1)}:${formatCellRef(r2, c2)}`);
        expect(range).toEqual({
          start: { row: Math.min(r1, r2), col: Math.min(c1, c2) },
          end: { row: Math.max(r1, r2), col: Math.max(c1, c2) },
        });
        expect(formatRange(range!)).toBe(
          r1 === r2 && c1 === c2
            ? formatCellRef(r1, c1)
            : `${formatCellRef(Math.min(r1, r2), Math.min(c1, c2))}:${formatCellRef(Math.max(r1, r2), Math.max(c1, c2))}`,
        );
      }),
      RUNS,
    );
  });
});

describe('serial dates (property)', () => {
  it('serialToDate never throws and returns null or a date between 0001-01-01 and 9999-12-31T23:59:59.999', () => {
    fc.assert(
      fc.property(fc.double(), systemArb, (serial, [system, conformance]) => {
        const date = serialToDate(serial, system, conformance);
        if (date === null) return;
        expect(date.getTime()).toBeGreaterThanOrEqual(lowest(system, conformance).getTime());
        expect(date.getTime()).toBeLessThanOrEqual(MAX_DATE.getTime());
        expect(date.getUTCMilliseconds()).toBe(Math.round(date.getUTCMilliseconds()));
      }),
      RUNS,
    );
  });

  it('dates survive dateToSerial then serialToDate unchanged (millisecond precision)', () => {
    fc.assert(
      fc.property(systemArb, fc.date({ min: MIN_DATE, max: MAX_DATE, noInvalidDate: true }), ([system, conformance], date) => {
        fc.pre(date.getTime() >= lowest(system, conformance).getTime());
        expect(serialToDate(dateToSerial(date, system, conformance), system, conformance)).toEqual(date);
      }),
      RUNS,
    );
  });

  it('serials survive serialToDate then dateToSerial to within half a millisecond', () => {
    fc.assert(
      fc.property(systemArb, fc.double({ min: -700_000, max: 3_000_000, noNaN: true }), ([system, conformance], serial) => {
        const date = serialToDate(serial, system, conformance);
        if (date === null) return;
        const back = dateToSerial(date, system, conformance);
        expect(Math.abs(back - serial) * DAY_MS).toBeLessThanOrEqual(0.5 + 1e-6 * Math.max(1, Math.abs(serial)));
      }),
      RUNS,
    );
  });

  it('dateToSerial is strictly increasing in every date system', () => {
    fc.assert(
      fc.property(
        systemArb,
        fc.date({ min: MIN_DATE, max: MAX_DATE, noInvalidDate: true }),
        fc.date({ min: MIN_DATE, max: MAX_DATE, noInvalidDate: true }),
        ([system, conformance], a, b) => {
          const min = lowest(system, conformance).getTime();
          fc.pre(a.getTime() >= min && b.getTime() >= min && a.getTime() !== b.getTime());
          const [early, late] = a.getTime() < b.getTime() ? [a, b] : [b, a];
          expect(dateToSerial(early, system, conformance)).toBeLessThan(dateToSerial(late, system, conformance));
        },
      ),
      RUNS,
    );
  });

  it('Transitional 1900 serials in [60, 61) never map to a date', () => {
    fc.assert(
      fc.property(fc.double({ min: 60, max: 61, maxExcluded: true, noNaN: true }), (serial) => {
        const date = serialToDate(serial, '1900', 'transitional');
        // Rounding to the nearest millisecond may carry a serial just below 61 into 1900-03-01.
        if (date !== null) expect(date.toISOString()).toBe('1900-03-01T00:00:00.000Z');
      }),
      RUNS,
    );
  });
});

describe('ISO 8601 date cells (property)', () => {
  it('every instant in range survives formatting and parseIsoDateTime', () => {
    fc.assert(
      fc.property(fc.date({ min: MIN_DATE, max: MAX_DATE, noInvalidDate: true }), (date) => {
        const text = date.toISOString(); // YYYY-MM-DDTHH:mm:ss.sssZ for years 0001–9999
        expect(parseIsoDateTime(text)).toEqual(date);
        expect(parseIsoDateTime(text.slice(0, -1))).toEqual(date);
      }),
      RUNS,
    );
  });

  it('parseIsoDateTime never throws on arbitrary strings', () => {
    const isoLike = fc.string({ unit: fc.constantFrom('0', '1', '2', '9', '-', ':', 'T', '.', 'Z', '+', ' '), maxLength: 30 });
    fc.assert(
      fc.property(fc.oneof(isoLike, fc.string({ unit: 'binary', maxLength: 30 })), (s) => {
        const date = parseIsoDateTime(s);
        if (date !== null) {
          expect(Number.isNaN(date.getTime())).toBe(false);
          expect(date.getTime()).toBeLessThanOrEqual(MAX_DATE.getTime());
          expect(date.getTime()).toBeGreaterThanOrEqual(MIN_DATE.getTime());
        }
      }),
      RUNS,
    );
  });
});

describe('number-format classification (property)', () => {
  const codeUnit = fc.constantFrom('y', 'm', 'd', 'h', 's', 'e', 'E', '+', '-', '0', '#', '"', '\\', '_', '*', '[', ']', '$', ';', '@', 'G', 'A', '/', 'P', ':', '.');
  const codeLike = fc.oneof(fc.string({ unit: codeUnit, maxLength: 40 }), fc.string({ unit: 'binary', maxLength: 40 }));

  it('never throws', () => {
    fc.assert(
      fc.property(codeLike, (code) => {
        expect(typeof isDateFormat(code)).toBe('boolean');
      }),
      RUNS,
    );
  });

  it('text inside double quotes is never a date', () => {
    fc.assert(
      fc.property(codeLike, (s) => {
        expect(isDateFormat(`0"${s.replaceAll('"', '')}"`)).toBe(false);
      }),
      RUNS,
    );
  });

  it('characters escaped with a backslash are never a date', () => {
    fc.assert(
      fc.property(codeLike, (s) => {
        const escaped = [...s].map((ch) => `\\${ch}`).join('');
        expect(isDateFormat(`0${escaped}`)).toBe(false);
      }),
      RUNS,
    );
  });

  it('a date token after closed literals is always found', () => {
    fc.assert(
      fc.property(codeLike, fc.constantFrom('yyyy', 'd', 'mm', 'h', 'ss', '[h]', '[mm]', '[ss]'), (s, token) => {
        const literal = `"${s.replaceAll('"', '')}"`;
        expect(isDateFormat(`${literal}${token}`)).toBe(true);
        expect(isDateFormat(`0;${literal}${token}`)).toBe(true);
      }),
      RUNS,
    );
  });
});
