import { describe, expect, it } from 'vitest';
import { dateToSerial, parseIsoDateTime, serialToDate } from '../../packages/xlvane/src/sml/dates';

/** A Date whose UTC fields are the given wall-clock values (month is 1-based; years 0–99 work). */
function utc(year: number, month: number, day: number, hour = 0, minute = 0, second = 0, ms = 0): Date {
  const date = new Date(Date.UTC(2000, month - 1, day, hour, minute, second, ms));
  date.setUTCFullYear(year);
  return date;
}

/** ISO string of a serialToDate result, or null. */
function iso(date: Date | null): string | null {
  return date === null ? null : date.toISOString();
}

const DAY = 86_400;

describe('serialToDate, Strict 1900 date system (Part 1 §18.17.4.1)', () => {
  const s = (serial: number): string | null => iso(serialToDate(serial, '1900', 'strict'));

  it('has base date 1899-12-30T00:00:00 = serial 0', () => {
    expect(s(0)).toBe('1899-12-30T00:00:00.000Z');
    expect(s(-0)).toBe('1899-12-30T00:00:00.000Z');
  });

  it('matches the worked examples of §18.17.4.1 and §18.17.4.3', () => {
    expect(s(1.5)).toBe('1899-12-31T12:00:00.000Z');
    expect(s(-4.25)).toBe('1899-12-25T18:00:00.000Z');
    expect(s(-1.25)).toBe('1899-12-28T18:00:00.000Z');
    expect(s(2.5)).toBe('1900-01-01T12:00:00.000Z');
    expect(s(3687 + 36_354 / DAY)).toBe('1910-02-03T10:05:54.000Z');
    expect(s(-2338 + 1 / DAY)).toBe('1893-08-05T00:00:01.000Z');
    expect(s(2_958_465 + 86_399 / DAY)).toBe('9999-12-31T23:59:59.000Z');
  });

  it('matches the truncated spec serials to within the precision they are printed with', () => {
    // The spec prints these serials cut short ("..."), so they may be off by up to one unit in
    // the last printed decimal place.
    const near = (serial: number, decimals: number, expected: string): void => {
      const date = serialToDate(serial, '1900', 'strict');
      expect(date).not.toBeNull();
      const tolerance = 10 ** -decimals * DAY * 1000;
      expect(Math.abs((date as Date).getTime() - Date.parse(expected))).toBeLessThanOrEqual(tolerance);
    };
    near(3687.4207639, 7, '1910-02-03T10:05:54.000Z');
    near(-2337.999989, 6, '1893-08-05T00:00:01.000Z');
    near(2_958_465.9999884, 7, '9999-12-31T23:59:59.000Z');
  });

  it('matches the §18.3.1.96 example: 1976-11-22T08:30 is 28086.3541666667', () => {
    expect(s(28_086.354_166_666_7)).toBe('1976-11-22T08:30:00.000Z');
  });

  it('has no 29 February 1900: serial 60 is 1900-02-28 and 61 is 1900-03-01', () => {
    expect(s(59)).toBe('1900-02-27T00:00:00.000Z');
    expect(s(60)).toBe('1900-02-28T00:00:00.000Z');
    expect(s(61)).toBe('1900-03-01T00:00:00.000Z');
  });

  it('accepts negative serials down to 0001-01-01T00:00:00 = -693593', () => {
    expect(s(-693_593)).toBe('0001-01-01T00:00:00.000Z');
    expect(s(-693_593.000_000_1)).toBeNull();
    expect(s(-693_594)).toBeNull();
  });

  it('accepts serials up to 9999-12-31T23:59:59.999', () => {
    expect(s(2_958_465 + 86_399_999 / (DAY * 1000))).toBe('9999-12-31T23:59:59.999Z');
    expect(s(2_958_466)).toBeNull();
  });
});

describe('serialToDate, Transitional 1900 date system (Part 4 §15.3.4)', () => {
  const t = (serial: number): string | null => iso(serialToDate(serial, '1900', 'transitional'));

  it('has 1900-01-01T00:00:00 = serial 1 and rejects negative serials', () => {
    expect(t(1)).toBe('1900-01-01T00:00:00.000Z');
    expect(t(1.5)).toBe('1900-01-01T12:00:00.000Z');
    expect(t(-1)).toBeNull();
    expect(t(-0.5)).toBeNull();
  });

  it('places time-only values (0 <= serial < 1) on 1899-12-31, the day of serial 0 (D-112)', () => {
    expect(t(0)).toBe('1899-12-31T00:00:00.000Z');
    expect(t(0.5)).toBe('1899-12-31T12:00:00.000Z');
    expect(t(0.75)).toBe('1899-12-31T18:00:00.000Z');
  });

  it('is the default when no conformance class is given', () => {
    expect(iso(serialToDate(1, '1900'))).toBe('1900-01-01T00:00:00.000Z');
    expect(iso(serialToDate(0.5, '1900'))).toBe('1899-12-31T12:00:00.000Z');
  });

  it('counts a 29 February 1900 that does not exist: serial 60 has no date', () => {
    expect(t(59)).toBe('1900-02-28T00:00:00.000Z');
    expect(t(59.75)).toBe('1900-02-28T18:00:00.000Z');
    expect(t(60)).toBeNull();
    expect(t(60.5)).toBeNull();
    expect(t(61)).toBe('1900-03-01T00:00:00.000Z');
  });

  it('matches the DATE examples of Part 1 §18.17.7.74 for the backward-compatibility date base', () => {
    expect(t(1)).toBe('1900-01-01T00:00:00.000Z');
    expect(t(2_958_465)).toBe('9999-12-31T00:00:00.000Z');
  });

  it('agrees with the Strict system from 1900-03-01 onwards', () => {
    expect(t(28_086.354_166_666_7)).toBe('1976-11-22T08:30:00.000Z');
    expect(t(61)).toBe(iso(serialToDate(61, '1900', 'strict')));
  });

  it('has upper limit 9999-12-31T23:59:59.999', () => {
    expect(t(2_958_465 + 86_399 / DAY)).toBe('9999-12-31T23:59:59.000Z');
    expect(t(2_958_465 + 86_399_999 / (DAY * 1000))).toBe('9999-12-31T23:59:59.999Z');
    expect(t(2_958_466)).toBeNull();
  });
});

describe('serialToDate, 1904 date system (Part 1 §18.17.4.1, unchanged by Part 4 §15.3.4)', () => {
  for (const conformance of ['strict', 'transitional'] as const) {
    const f = (serial: number): string | null => iso(serialToDate(serial, '1904', conformance));

    it(`has base date 1904-01-01 = serial 0 (${conformance})`, () => {
      expect(f(0)).toBe('1904-01-01T00:00:00.000Z');
      expect(f(0.5)).toBe('1904-01-01T12:00:00.000Z');
      expect(f(-1)).toBe('1903-12-31T00:00:00.000Z');
    });

    it(`matches the worked examples of §18.17.4.3 (${conformance})`, () => {
      expect(f(-3800 + 1 / DAY)).toBe('1893-08-05T00:00:01.000Z');
      expect(f(2225 + 36_354 / DAY)).toBe('1910-02-03T10:05:54.000Z');
      expect(f(2_957_003 + 86_399 / DAY)).toBe('9999-12-31T23:59:59.000Z');
    });

    it(`has range 0001-01-01 (-695055) to 9999-12-31T23:59:59.999 (${conformance})`, () => {
      expect(f(-695_055)).toBe('0001-01-01T00:00:00.000Z');
      expect(f(-695_056)).toBeNull();
      expect(f(2_957_003 + 86_399_999 / (DAY * 1000))).toBe('9999-12-31T23:59:59.999Z');
      expect(f(2_957_004)).toBeNull();
    });
  }
});

describe('serialToDate, time fraction (Part 1 §18.17.4.2)', () => {
  it('rounds the time of day to the nearest millisecond', () => {
    expect(iso(serialToDate(0.5 + 0.4 / (DAY * 1000), '1900', 'strict'))).toBe('1899-12-30T12:00:00.000Z');
    expect(iso(serialToDate(0.5 + 0.6 / (DAY * 1000), '1900', 'strict'))).toBe('1899-12-30T12:00:00.001Z');
  });

  it('uses the greatest integer not exceeding the serial as the day (spec examples)', () => {
    // -2.00000001 is in day -3 with time 0.99999999; 4.66666667 is in day 4 with time 0.66666667.
    expect(iso(serialToDate(-2.000_000_01, '1900', 'strict'))).toBe('1899-12-27T23:59:59.999Z');
    expect(iso(serialToDate(4.666_666_67, '1900', 'strict'))).toBe('1900-01-03T16:00:00.000Z');
  });

  it('applies the range check after rounding', () => {
    expect(iso(serialToDate(0.999_999_999_9, '1900', 'transitional'))).toBe('1900-01-01T00:00:00.000Z');
    expect(serialToDate(59.999_999_999_9, '1900', 'transitional')).toBeNull();
    expect(iso(serialToDate(60.999_999_999_9, '1900', 'transitional'))).toBe('1900-03-01T00:00:00.000Z');
    expect(serialToDate(2_958_465.999_999_999_9, '1900', 'strict')).toBeNull();
  });

  it('returns null for NaN, infinities and huge values', () => {
    for (const bad of [Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY, 1e308, -1e308, Number.MAX_SAFE_INTEGER]) {
      expect(serialToDate(bad, '1900', 'strict')).toBeNull();
      expect(serialToDate(bad, '1900', 'transitional')).toBeNull();
      expect(serialToDate(bad, '1904', 'strict')).toBeNull();
    }
  });
});

describe('dateToSerial', () => {
  it('converts in the Strict 1900 system', () => {
    expect(dateToSerial(utc(1899, 12, 30), '1900', 'strict')).toBe(0);
    expect(dateToSerial(utc(1900, 1, 1, 12), '1900', 'strict')).toBe(2.5);
    expect(dateToSerial(utc(1899, 12, 25, 18), '1900', 'strict')).toBe(-4.25);
    expect(dateToSerial(utc(1900, 2, 28), '1900', 'strict')).toBe(60);
    expect(dateToSerial(utc(1900, 3, 1), '1900', 'strict')).toBe(61);
    expect(dateToSerial(utc(1, 1, 1), '1900', 'strict')).toBe(-693_593);
    expect(dateToSerial(utc(9999, 12, 31, 23, 59, 59, 999), '1900', 'strict')).toBeCloseTo(2_958_465.999_999_988, 6);
  });

  it('converts in the Transitional 1900 system (the default), skipping serial 60', () => {
    expect(dateToSerial(utc(1900, 1, 1), '1900')).toBe(1);
    expect(dateToSerial(utc(1900, 1, 1), '1900', 'transitional')).toBe(1);
    expect(dateToSerial(utc(1900, 2, 28, 18), '1900', 'transitional')).toBe(59.75);
    expect(dateToSerial(utc(1900, 3, 1), '1900', 'transitional')).toBe(61);
    expect(dateToSerial(utc(9999, 12, 31), '1900', 'transitional')).toBe(2_958_465);
    expect(dateToSerial(utc(1976, 11, 22, 8, 30), '1900', 'transitional')).toBeCloseTo(28_086.354_166_666_7, 9);
  });

  it('converts in the 1904 system for both conformance classes', () => {
    for (const conformance of ['strict', 'transitional'] as const) {
      expect(dateToSerial(utc(1904, 1, 1), '1904', conformance)).toBe(0);
      expect(dateToSerial(utc(1904, 1, 1, 12), '1904', conformance)).toBe(0.5);
      expect(dateToSerial(utc(1893, 8, 5), '1904', conformance)).toBe(-3800);
      expect(dateToSerial(utc(1, 1, 1), '1904', conformance)).toBe(-695_055);
    }
  });

  it('throws RangeError for dates outside the selected date system', () => {
    expect(() => dateToSerial(utc(1899, 12, 30, 23, 59, 59, 999), '1900', 'transitional')).toThrow(RangeError);
    expect(dateToSerial(utc(1899, 12, 31, 12), '1900', 'transitional')).toBe(0.5);
    expect(() => dateToSerial(utc(0, 12, 31, 23, 59, 59, 999), '1900', 'strict')).toThrow(RangeError);
    expect(() => dateToSerial(utc(0, 12, 31), '1904', 'transitional')).toThrow(RangeError);
    expect(() => dateToSerial(utc(10_000, 1, 1), '1900', 'strict')).toThrow(RangeError);
    expect(() => dateToSerial(utc(10_000, 1, 1), '1904', 'strict')).toThrow(RangeError);
  });

  it('throws RangeError for an invalid Date', () => {
    expect(() => dateToSerial(new Date(Number.NaN), '1900')).toThrow(RangeError);
  });

  it('round-trips with serialToDate', () => {
    const dates = [utc(1, 1, 1, 0, 0, 0, 1), utc(1900, 1, 1), utc(1900, 2, 28, 23, 59, 59, 999), utc(1900, 3, 1), utc(2024, 2, 29, 13, 37, 1, 234), utc(9999, 12, 31, 23, 59, 59, 999)];
    for (const date of dates) {
      for (const [system, conformance] of [
        ['1900', 'strict'],
        ['1900', 'transitional'],
        ['1904', 'strict'],
      ] as const) {
        // The Transitional 1900 system starts at 1900-01-01.
        if (system === '1900' && conformance === 'transitional' && date.getUTCFullYear() < 1900) continue;
        expect(serialToDate(dateToSerial(date, system, conformance), system, conformance)).toEqual(date);
      }
    }
  });
});

describe('parseIsoDateTime (Part 1 §18.17.4)', () => {
  const p = (text: string): string | null => iso(parseIsoDateTime(text));

  it('parses a date-only value (Complete, Extended Format Calendar Date)', () => {
    expect(p('1975-10-05')).toBe('1975-10-05T00:00:00.000Z');
    expect(p('1582-10-15')).toBe('1582-10-15T00:00:00.000Z');
    expect(p('0001-01-01')).toBe('0001-01-01T00:00:00.000Z');
    expect(p('9999-12-31')).toBe('9999-12-31T00:00:00.000Z');
    expect(p('2000-02-29')).toBe('2000-02-29T00:00:00.000Z');
  });

  it('parses date and time values with minutes, seconds and fractional seconds', () => {
    expect(p('1976-11-22T08:30')).toBe('1976-11-22T08:30:00.000Z');
    expect(p('1976-11-22T08:30:00')).toBe('1976-11-22T08:30:00.000Z');
    expect(p('1976-11-22T08:30:00.000')).toBe('1976-11-22T08:30:00.000Z');
    expect(p('1582-10-15T00:00')).toBe('1582-10-15T00:00:00.000Z');
    expect(p('9999-12-31T23:59:59.999')).toBe('9999-12-31T23:59:59.999Z');
  });

  it('parses time-only values onto 1899-12-30, the day of serial 0 in the 1900 date system', () => {
    expect(p('08:30')).toBe('1899-12-30T08:30:00.000Z');
    expect(p('08:30:00')).toBe('1899-12-30T08:30:00.000Z');
    expect(p('08:30:00.000')).toBe('1899-12-30T08:30:00.000Z');
    expect(p('00:00')).toBe('1899-12-30T00:00:00.000Z');
  });

  it('rounds fractional seconds to the nearest millisecond', () => {
    expect(p('2020-01-01T12:00:00.5')).toBe('2020-01-01T12:00:00.500Z');
    expect(p('2020-01-01T12:00:00.1234')).toBe('2020-01-01T12:00:00.123Z');
    expect(p('2020-01-01T12:00:00.0005')).toBe('2020-01-01T12:00:00.001Z');
    expect(p('1999-12-31T23:59:59.9996')).toBe('2000-01-01T00:00:00.000Z');
    expect(p('2020-01-01T12:00:00.123456789')).toBe('2020-01-01T12:00:00.123Z');
    expect(p('9999-12-31T23:59:59.9996')).toBeNull();
  });

  it('accepts a trailing "Z" after a time (Annex L.2.16.9.3 examples)', () => {
    expect(p('1893-08-05T00:00:01Z')).toBe('1893-08-05T00:00:01.000Z');
    expect(p('2020-01-01T12:00:00.000Z')).toBe('2020-01-01T12:00:00.000Z');
  });

  it('rejects invalid calendar dates (proleptic Gregorian) and out-of-range years', () => {
    for (const bad of ['1900-02-29', '1975-02-29', '1975-13-01', '1975-00-10', '1975-04-31', '1975-10-00', '1975-10-32', '0000-01-01', '10000-01-01']) {
      expect(p(bad), bad).toBeNull();
    }
  });

  it('rejects hour 24, minute 60 and leap seconds', () => {
    for (const bad of ['1975-10-05T24:00', '1975-10-05T24:00:00', '24:00', '1975-10-05T23:60', '1975-10-05T23:59:60', '23:59:60']) {
      expect(p(bad), bad).toBeNull();
    }
  });

  it('rejects other lexical forms', () => {
    const bad = [
      '',
      'T',
      '1975',
      '1975-10',
      '1975-10-5',
      '75-10-05',
      '19751005',
      '+01975-10-05',
      '-1975-10-05',
      '1975-10-05T',
      '1975-10-05T08',
      '1975-10-05T8:30',
      '1975-10-05T08:3',
      '1975-10-05 08:30',
      '1975-10-05t08:30',
      '1975-10-05T08:30:00.',
      '1975-10-05T08:30:00,5',
      '1975-10-05T08:30.5',
      '1975-10-05T08:30:00+02:00',
      '1975-10-05T08:30:00-05:00',
      '1975-10-05Z',
      '1975-10-05T08:30:00ZZ',
      ' 1975-10-05',
      '1975-10-05 ',
      '１９７５-10-05',
      '1975-10-05T08:30:00.1234567890',
    ];
    for (const text of bad) {
      expect(p(text), JSON.stringify(text)).toBeNull();
    }
  });

  it('rejects very long input quickly', () => {
    const start = performance.now();
    expect(parseIsoDateTime(`2020-01-01T00:00:00.${'0'.repeat(1_000_000)}`)).toBeNull();
    expect(parseIsoDateTime('9'.repeat(1_000_000))).toBeNull();
    expect(performance.now() - start).toBeLessThan(100);
  });
});
