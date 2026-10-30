import { describe, expect, it } from 'vitest';
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

describe('grid limits', () => {
  it('match Part 1 §18.17.5.1 (A1 to XFD1048576)', () => {
    expect(MAX_ROWS).toBe(1_048_576);
    expect(MAX_COLUMNS).toBe(16_384);
    expect(columnNumber('XFD')).toBe(MAX_COLUMNS);
  });
});

describe('columnLetters', () => {
  it('names columns A–Z, AA–AZ, BA–BZ, …, ZA–ZZ, AAA… (Part 1 §18.17.2.3.1)', () => {
    expect(columnLetters(1)).toBe('A');
    expect(columnLetters(26)).toBe('Z');
    expect(columnLetters(27)).toBe('AA');
    expect(columnLetters(52)).toBe('AZ');
    expect(columnLetters(53)).toBe('BA');
    expect(columnLetters(677)).toBe('ZA');
    expect(columnLetters(702)).toBe('ZZ');
    expect(columnLetters(703)).toBe('AAA');
    expect(columnLetters(728)).toBe('AAZ');
    expect(columnLetters(729)).toBe('ABA');
    expect(columnLetters(16_384)).toBe('XFD');
  });

  it('throws RangeError outside 1..MAX_COLUMNS or for non-integers', () => {
    for (const bad of [0, -1, 16_385, 1.5, Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY]) {
      expect(() => columnLetters(bad)).toThrow(RangeError);
    }
  });

  it('round-trips with columnNumber for every column', () => {
    for (let col = 1; col <= MAX_COLUMNS; col++) {
      expect(columnNumber(columnLetters(col))).toBe(col);
    }
  });
});

describe('columnNumber', () => {
  it('converts letters to 1-based column numbers', () => {
    expect(columnNumber('A')).toBe(1);
    expect(columnNumber('Z')).toBe(26);
    expect(columnNumber('AA')).toBe(27);
    expect(columnNumber('ZZ')).toBe(702);
    expect(columnNumber('AAA')).toBe(703);
    expect(columnNumber('XFD')).toBe(16_384);
  });

  it('is not case-sensitive (Part 1 §18.17.2.3.1)', () => {
    expect(columnNumber('xfd')).toBe(16_384);
    expect(columnNumber('xFd')).toBe(16_384);
    expect(columnNumber('a')).toBe(1);
  });

  it('returns null for columns beyond XFD', () => {
    expect(columnNumber('XFE')).toBeNull();
    expect(columnNumber('ZZZ')).toBeNull();
    expect(columnNumber('AAAA')).toBeNull();
  });

  it('returns null for anything that is not 1–3 ASCII letters', () => {
    for (const bad of ['', '$A', 'A1', ' A', 'A ', 'Ä', 'ａ', 'A-', '1', 'A'.repeat(100_000)]) {
      expect(columnNumber(bad)).toBeNull();
    }
  });
});

describe('parseCellRef', () => {
  it('parses A1-style references stored in c@r', () => {
    expect(parseCellRef('A1')).toEqual({ row: 1, col: 1 });
    expect(parseCellRef('B7')).toEqual({ row: 7, col: 2 });
    expect(parseCellRef('AA10')).toEqual({ row: 10, col: 27 });
    expect(parseCellRef('XFD1048576')).toEqual({ row: 1_048_576, col: 16_384 });
  });

  it('accepts lowercase column letters (column letters are not case-sensitive)', () => {
    expect(parseCellRef('b7')).toEqual({ row: 7, col: 2 });
    expect(parseCellRef('xfd1048576')).toEqual({ row: 1_048_576, col: 16_384 });
    expect(parseCellRef('aB12')).toEqual({ row: 12, col: 28 });
  });

  it('rejects $ markers: ST_CellRef is not the "might be absolute" ST_RefA type', () => {
    for (const bad of ['$A$1', '$A1', 'A$1', '$B$7']) {
      expect(parseCellRef(bad)).toBeNull();
    }
  });

  it('accepts leading zeros in the row digits (a digit-sequence whose value is non-zero)', () => {
    expect(parseCellRef('A01')).toEqual({ row: 1, col: 1 });
    expect(parseCellRef('C0000007')).toEqual({ row: 7, col: 3 });
  });

  it('rejects row 0 and references outside the grid', () => {
    expect(parseCellRef('A0')).toBeNull();
    expect(parseCellRef('A000')).toBeNull();
    expect(parseCellRef('A1048577')).toBeNull();
    expect(parseCellRef('XFE1')).toBeNull();
    expect(parseCellRef('AAAA1')).toBeNull();
    expect(parseCellRef('A12345678')).toBeNull();
  });

  it('rejects malformed references', () => {
    const bad = [
      '',
      'A',
      '1',
      '1A',
      ' A1',
      'A1 ',
      'A 1',
      'A-1',
      'A+1',
      'A1.0',
      'A1e3',
      'A1:B2',
      'R1C1',
      'A１',
      'A١',
      'Sheet1!A1',
      'A1\u0000',
    ];
    for (const ref of bad) {
      expect(parseCellRef(ref), JSON.stringify(ref)).toBeNull();
    }
  });

  it('rejects very long input quickly', () => {
    const start = performance.now();
    expect(parseCellRef(`A${'1'.repeat(1_000_000)}`)).toBeNull();
    expect(parseCellRef('A'.repeat(1_000_000))).toBeNull();
    expect(performance.now() - start).toBeLessThan(100);
  });
});

describe('formatCellRef', () => {
  it('formats row and column as an A1 reference', () => {
    expect(formatCellRef(7, 2)).toBe('B7');
    expect(formatCellRef(1, 1)).toBe('A1');
    expect(formatCellRef(1_048_576, 16_384)).toBe('XFD1048576');
  });

  it('throws RangeError outside the grid', () => {
    expect(() => formatCellRef(0, 1)).toThrow(RangeError);
    expect(() => formatCellRef(1_048_577, 1)).toThrow(RangeError);
    expect(() => formatCellRef(1, 0)).toThrow(RangeError);
    expect(() => formatCellRef(1, 16_385)).toThrow(RangeError);
    expect(() => formatCellRef(1.5, 1)).toThrow(RangeError);
    expect(() => formatCellRef(Number.NaN, 1)).toThrow(RangeError);
  });

  it('round-trips with parseCellRef', () => {
    for (const [row, col] of [
      [1, 1],
      [7, 2],
      [999, 703],
      [1_048_576, 16_384],
    ] as const) {
      expect(parseCellRef(formatCellRef(row, col))).toEqual({ row, col });
    }
  });
});

describe('parseRange', () => {
  it('parses a two-cell range (mergeCell@ref, dimension@ref)', () => {
    expect(parseRange('A1:C3')).toEqual({ start: { row: 1, col: 1 }, end: { row: 3, col: 3 } });
    expect(parseRange('B5:D8')).toEqual({ start: { row: 5, col: 2 }, end: { row: 8, col: 4 } });
  });

  it('parses a single cell as a one-cell range', () => {
    expect(parseRange('B2')).toEqual({ start: { row: 2, col: 2 }, end: { row: 2, col: 2 } });
    expect(parseRange('B2:B2')).toEqual({ start: { row: 2, col: 2 }, end: { row: 2, col: 2 } });
  });

  it('normalises corners so that start <= end in both dimensions', () => {
    expect(parseRange('C3:A1')).toEqual({ start: { row: 1, col: 1 }, end: { row: 3, col: 3 } });
    expect(parseRange('A3:C1')).toEqual({ start: { row: 1, col: 1 }, end: { row: 3, col: 3 } });
    expect(parseRange('C1:A3')).toEqual({ start: { row: 1, col: 1 }, end: { row: 3, col: 3 } });
  });

  it('accepts lowercase letters', () => {
    expect(parseRange('a1:c3')).toEqual({ start: { row: 1, col: 1 }, end: { row: 3, col: 3 } });
  });

  it('rejects whole-row and whole-column references (dimension does not support them)', () => {
    expect(parseRange('A:C')).toBeNull();
    expect(parseRange('1:3')).toBeNull();
    expect(parseRange('A:A')).toBeNull();
  });

  it('rejects $ markers and malformed ranges', () => {
    const bad = ['', ':', 'A1:', ':A1', 'A1:B2:C3', 'A1 :B2', 'A1: B2', '$A$1:$B$2', 'A1:$B$2', 'A1:XFE1', 'A0:B2', 'A1;B2', 'A1,B2', 'A1 B2'];
    for (const ref of bad) {
      expect(parseRange(ref), JSON.stringify(ref)).toBeNull();
    }
  });

  it('rejects very long input quickly', () => {
    const start = performance.now();
    expect(parseRange(`A1:${'B'.repeat(1_000_000)}`)).toBeNull();
    expect(parseRange(':'.repeat(1_000_000))).toBeNull();
    expect(performance.now() - start).toBeLessThan(100);
  });
});

describe('formatRange', () => {
  it('formats a range as "start:end"', () => {
    expect(formatRange({ start: { row: 1, col: 1 }, end: { row: 3, col: 3 } })).toBe('A1:C3');
  });

  it('formats a single-cell range as one reference', () => {
    expect(formatRange({ start: { row: 2, col: 2 }, end: { row: 2, col: 2 } })).toBe('B2');
  });

  it('throws RangeError for corners outside the grid', () => {
    expect(() => formatRange({ start: { row: 0, col: 1 }, end: { row: 1, col: 1 } })).toThrow(RangeError);
    expect(() => formatRange({ start: { row: 1, col: 1 }, end: { row: 1, col: 16_385 } })).toThrow(RangeError);
  });

  it('round-trips with parseRange', () => {
    for (const ref of ['A1:C3', 'B2', 'A1:XFD1048576', 'Z9:AA10']) {
      const range = parseRange(ref);
      expect(range).not.toBeNull();
      expect(formatRange(range!)).toBe(ref);
    }
  });
});
