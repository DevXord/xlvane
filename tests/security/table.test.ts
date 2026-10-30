// Untrusted CSV, JSON and sheet content in the task API: prototype keys, injection, limits and
// linear-time parsing.
import { describe, expect, it } from 'vitest';
import { fromCSV, parseCsvRows, toCSV } from '../../packages/xlvane/src/csv';
import { fromJSON, readTable, writeTable } from '../../packages/xlvane/src/table';
import { expectXlvError } from '../helpers/expect-error';
import { expectPrototypesUntouched } from '../helpers/prototypes';

describe('prototype keys as data', () => {
  it('reads __proto__, constructor and prototype headers as ordinary keys', () => {
    const t = readTable(fromCSV('__proto__,constructor,prototype,polluted\n{"x":1},1,2,yes\n'));
    const row = t.rows[0]!;
    expect(Object.keys(row)).toEqual(['__proto__', 'constructor', 'prototype', 'polluted']);
    expect(Object.getPrototypeOf(row)).toBe(Object.prototype);
    expect(row['__proto__']).toBe('{"x":1}');
    expectPrototypesUntouched();
  });

  it('writes JSON with __proto__ keys as a column, without following them', () => {
    const bytes = fromJSON('[{"__proto__":"p","a":1}]');
    expect(readTable(bytes).columns.map((c) => c.key)).toEqual(['__proto__', 'a']);
    expect(readTable(bytes).rows[0]?.['__proto__']).toBe('p');
    expectXlvError(() => fromJSON('[{"__proto__":{"polluted":"yes"}}]'), 'XLV_E_TABLE_INVALID');
    expectPrototypesUntouched();
  });

  it('ignores inherited properties of row objects', () => {
    const row = Object.create({ inherited: 1 }) as Record<string, unknown>;
    row['own'] = 2;
    expect(readTable(writeTable([row])).columns.map((c) => c.key)).toEqual(['own']);
  });
});

describe('formula injection', () => {
  it('escapes injected text even when it came from a CSV import', () => {
    const csv = toCSV(fromCSV('a\n"=cmd|\' /C calc\'!A0"\n', { types: 'text' }), { newline: '\n' });
    expect(csv).toBe(`a\n"'=cmd|' /C calc'!A0"\n`);
  });

  it('never creates formula cells from CSV or JSON input', () => {
    for (const bytes of [fromCSV('a\n=1+1\n+1\n@x\n'), fromJSON([{ a: '=1+1' }])]) {
      const wb = readTable(bytes);
      expect(wb.rows.every((r) => typeof r['a'] === 'string' || typeof r['a'] === 'number')).toBe(true);
    }
  });
});

describe('CSV limits and parser cost', () => {
  it('rejects more than 16 384 fields on a line', () => {
    expectXlvError(() => fromCSV(','.repeat(16_384)), 'XLV_E_LIMIT_EXCEEDED', { limit: 'columns' });
  });

  it('rejects text over maxChars', () => {
    expectXlvError(() => fromCSV('a,b\n1,2\n', { maxChars: 4 }), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxChars' });
  });

  it('rejects fields longer than a cell can hold', () => {
    expectXlvError(() => fromCSV(`a\n${'x'.repeat(32_768)}\n`), 'XLV_E_INVALID_VALUE');
  });

  it('parses hostile quoting in linear time', () => {
    const inputs = ['"'.repeat(200_001), '""'.repeat(100_000), `"${'""'.repeat(100_000)}"`, 'a,'.repeat(10_000) + '\n'.repeat(10_000), '"x\n'.repeat(1)];
    for (const text of inputs) {
      const start = performance.now();
      try {
        parseCsvRows(text, ',');
      } catch {
        // malformed input is fine; only the time matters here
      }
      expect(performance.now() - start).toBeLessThan(500);
    }
  });
});
