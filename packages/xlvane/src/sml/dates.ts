// Serial date-times and ISO 8601 date cells (ECMA-376 Part 1 §18.17.4, Part 4 §15.3.4).
//
// Dates are `Date` objects whose UTC fields hold the spreadsheet's wall-clock date and time; no
// time-zone conversion is done (Part 1 §18.17.4.1: serial date-times represent local time).
//
// The date systems (see the facts fragment for the sources):
// - Strict 1900 (Part 1 §18.17.4.1): serial 0 = 1899-12-30T00:00; range 0001-01-01 (-693593) to
//   9999-12-31T23:59:59.999. There is no 29 February 1900 (§18.17.4 uses the proleptic Gregorian
//   calendar), so serial 60 is 1900-02-28.
// - Transitional 1900 (Part 4 §15.3.4 replaces Part 1 §18.17.4.1): serial 1 = 1900-01-01T00:00
//   is the lower limit and 2958465 = 9999-12-31. That span is one day longer than the calendar,
//   which matches Microsoft's documentation that Excel treats 1900 as a leap year (support
//   article "wrongly-assumes-1900-is-leap-year"). Serial 60 is the nonexistent 1900-02-29 and
//   has no `Date`.
// - 1904 (Part 1 §18.17.4.1, unchanged by Part 4): serial 0 = 1904-01-01T00:00; range
//   0001-01-01 (-695055) to 9999-12-31T23:59:59.999.
import type { Conformance, DateSystem } from '../model';

const DAY_MS = 86_400_000;

/** Serial 0 of the Strict 1900 system, and of the Transitional one from 1900-03-01 onwards. */
const EPOCH_1899_12_30 = Date.UTC(1899, 11, 30);
/** Serial 0 of the Transitional 1900 system for 1900-01-01 to 1900-02-28 (Part 4 §15.3.4). */
const EPOCH_1899_12_31 = Date.UTC(1899, 11, 31);
/** Serial 0 of the 1904 system (Part 1 §18.17.4.1). */
const EPOCH_1904_01_01 = Date.UTC(1904, 0, 1);
const TIME_1900_03_01 = Date.UTC(1900, 2, 1);

/** Earliest and latest permitted instants (Part 1 §18.17.4 and §18.17.4.1). */
const MIN_TIME = new Date(0).setUTCFullYear(1, 0, 1);
const MAX_TIME = Date.UTC(9999, 11, 31, 23, 59, 59, 999);

/** No date system reaches beyond these serials; checked before multiplying, to keep precision. */
const MIN_ANY_SERIAL = -695_056;
const MAX_ANY_SERIAL = 2_958_467;

/**
 * Converts a serial date-time to a `Date` whose UTC fields are the spreadsheet's wall-clock date
 * and time. The time of day is rounded to the nearest millisecond, and the result is `null` when
 * the rounded value is outside the date system's valid range (ECMA-376 Part 1 §18.17.4.1, or
 * Part 4 §15.3.4 for Transitional documents) or is the Transitional serial 60 (the nonexistent
 * 29 February 1900). `conformance` selects the 1900 date system's variant; it defaults to
 * `'transitional'`, the class of documents without `workbook@conformance` (Part 1 §18.2.27).
 */
export function serialToDate(serial: number, system: DateSystem, conformance: Conformance = 'transitional'): Date | null {
  if (!Number.isFinite(serial) || serial < MIN_ANY_SERIAL || serial > MAX_ANY_SERIAL) return null;
  // The day is the greatest integer not exceeding the serial, the rest is the time of day
  // (Part 1 §18.17.4.2); rounding the whole value to milliseconds keeps both consistent.
  const ms = Math.round(serial * DAY_MS);
  let time: number;
  if (system === '1904') {
    time = EPOCH_1904_01_01 + ms;
  } else if (conformance === 'strict') {
    time = EPOCH_1899_12_30 + ms;
  } else {
    // Part 4 §15.3.4: the lower limit is 1900-01-01T00:00:00 (serial 1) and serial 0 is
    // 1899-12-31. Values in [0, 1) are time-only values (e.g. 0.5 formatted as h:mm), common in
    // real files, so they are placed on 1899-12-31 instead of being rejected (D-112).
    if (ms < 0) return null;
    if (ms < 60 * DAY_MS) time = EPOCH_1899_12_31 + ms;
    // Serial 60 is 29 February 1900, which the proleptic Gregorian calendar does not have.
    else if (ms < 61 * DAY_MS) return null;
    else time = EPOCH_1899_12_30 + ms;
  }
  if (time < MIN_TIME || time > MAX_TIME) return null;
  return new Date(time);
}

/**
 * Converts a `Date` (read through its UTC fields as wall-clock time) to a serial date-time in the
 * given date system; `conformance` defaults to `'transitional'` as in {@link serialToDate}.
 * Throws `RangeError` for an invalid `Date` or one outside the date system's valid range.
 */
export function dateToSerial(date: Date, system: DateSystem, conformance: Conformance = 'transitional'): number {
  const time = date.getTime();
  if (Number.isNaN(time)) throw new RangeError('Cannot convert an invalid Date to a serial date-time.');
  if (time < MIN_TIME || time > MAX_TIME) {
    throw new RangeError(`Date ${date.toISOString()} is outside 0001-01-01T00:00:00.000 to 9999-12-31T23:59:59.999.`);
  }
  if (system === '1904') return (time - EPOCH_1904_01_01) / DAY_MS;
  if (conformance === 'strict') return (time - EPOCH_1899_12_30) / DAY_MS;
  if (time < EPOCH_1899_12_31) {
    throw new RangeError(
      `Date ${date.toISOString()} is before 1899-12-31 (serial 0, used for time-only values), the start of the Transitional 1900 date system.`,
    );
  }
  // Serials 1–59.99… are 1900-01-01 to 1900-02-28; serial 60 is skipped (Part 4 §15.3.4).
  if (time < TIME_1900_03_01) return (time - EPOCH_1899_12_31) / DAY_MS;
  return (time - EPOCH_1899_12_30) / DAY_MS;
}

const CODE_0 = 0x30;
const CODE_9 = 0x39;
const CODE_HYPHEN = 0x2d;
const CODE_COLON = 0x3a;
const CODE_PERIOD = 0x2e;
const CODE_T = 0x54;
const CODE_Z = 0x5a;

/** Longest accepted text: "YYYY-MM-DDThh:mm:ss." + 9 fraction digits + "Z". */
const MAX_ISO_LENGTH = 30;
const MAX_FRACTION_DIGITS = 9;

/** The number formed by `count` ASCII digits at `text[at]`, or -1 if any of them is not a digit. */
function digitsAt(text: string, at: number, count: number): number {
  if (at + count > text.length) return -1;
  let value = 0;
  for (let i = at; i < at + count; i++) {
    const code = text.charCodeAt(i);
    if (code < CODE_0 || code > CODE_9) return -1;
    value = value * 10 + (code - CODE_0);
  }
  return value;
}

function daysInMonth(year: number, month: number): number {
  if (month === 2) {
    // Proleptic Gregorian leap years (Part 1 §18.17.4).
    const leap = (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
    return leap ? 29 : 28;
  }
  return month === 4 || month === 6 || month === 9 || month === 11 ? 30 : 31;
}

/**
 * Parses the ISO 8601 lexical forms of Part 1 §18.17.4 used by `t="d"` cells: a date
 * (`1975-10-05`), a time of day (`08:30`, `08:30:00`, `08:30:00.000`) or both joined by `T`
 * (`1976-11-22T08:30:00.000`). A trailing `Z` after the time is accepted (as in the Annex L
 * examples); other time-zone designators are not. Fractional seconds (1–9 digits) are rounded to
 * the nearest millisecond. A time without a date is placed on 1899-12-30, the day of serial 0 in
 * the Strict 1900 date system. Returns `null` for any other text, an impossible date or time
 * (hour 24 and leap seconds are not permitted), or a value outside 0001-01-01 to
 * 9999-12-31T23:59:59.999. The result's UTC fields hold the wall-clock value.
 */
export function parseIsoDateTime(text: string): Date | null {
  const length = text.length;
  if (length < 5 || length > MAX_ISO_LENGTH) return null;

  let year = 1899;
  let month = 12;
  let day = 30;
  let i = 0;
  const hasDate = length >= 10 && text.charCodeAt(4) === CODE_HYPHEN;
  if (hasDate) {
    // Complete, extended format calendar date: YYYY-MM-DD.
    year = digitsAt(text, 0, 4);
    month = digitsAt(text, 5, 2);
    day = digitsAt(text, 8, 2);
    if (year < 1 || text.charCodeAt(7) !== CODE_HYPHEN || month < 1 || month > 12) return null;
    if (day < 1 || day > daysInMonth(year, month)) return null;
    i = 10;
  }

  let hour = 0;
  let minute = 0;
  let second = 0;
  let fractionMs = 0;
  if (!hasDate || i < length) {
    if (hasDate) {
      if (text.charCodeAt(i) !== CODE_T) return null;
      i++;
    }
    // Complete, extended format time of day: hh:mm, hh:mm:ss or hh:mm:ss.s…
    hour = digitsAt(text, i, 2);
    minute = digitsAt(text, i + 3, 2);
    if (hour < 0 || text.charCodeAt(i + 2) !== CODE_COLON || minute < 0) return null;
    i += 5;
    if (i < length && text.charCodeAt(i) === CODE_COLON) {
      second = digitsAt(text, i + 1, 2);
      if (second < 0) return null;
      i += 3;
      // Only seconds may have a decimal separator, and it is a full stop (Part 1 §18.17.4).
      if (i < length && text.charCodeAt(i) === CODE_PERIOD) {
        i++;
        const fractionStart = i;
        let fraction = 0;
        while (i < length) {
          const code = text.charCodeAt(i);
          if (code < CODE_0 || code > CODE_9) break;
          fraction = fraction * 10 + (code - CODE_0);
          i++;
        }
        const digits = i - fractionStart;
        if (digits < 1 || digits > MAX_FRACTION_DIGITS) return null;
        fractionMs = Math.round((fraction * 1000) / 10 ** digits);
      }
    }
    if (i < length && text.charCodeAt(i) === CODE_Z) i++;
    if (i !== length) return null;
    // Midnight is hour 0, never 24, and leap seconds are not permitted (Part 1 §18.17.4).
    if (hour > 23 || minute > 59 || second > 59) return null;
  }

  const time = new Date(0).setUTCFullYear(year, month - 1, day) + ((hour * 60 + minute) * 60 + second) * 1000 + fractionMs;
  if (time < MIN_TIME || time > MAX_TIME) return null;
  return new Date(time);
}
