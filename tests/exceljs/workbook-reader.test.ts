// Ported from exceljs v4.4.0 spec/integration/workbook-xlsx-reader.spec.js (MIT, Copyright (c)
// 2014-2019 Guyon Roche), changes: every upstream test reads a binary fixture from
// spec/integration/data/ or uses createTestBook; no fixtures were copied, so they are listed as
// skipped. The first two tests are added for xlvane (not upstream): one streams the rich-text
// sample (tests/corpus/excel/exceljs-rich-text-sample.xlsx); one writes a file with Workbook
// and reads it back with stream.xlsx.WorkbookReader, using the reader shape from exceljs
// lib/stream/xlsx/workbook-reader.js and worksheet-reader.js (async iteration).
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { ExcelJS, withTempDir } from './helpers';

describe('WorkbookReader', () => {
  it('streams rich text values as { richText } (added for xlvane)', async () => {
    const bytes = readFileSync(new URL('../corpus/excel/exceljs-rich-text-sample.xlsx', import.meta.url));
    const reader = new ExcelJS.stream.xlsx.WorkbookReader(bytes, {});
    for await (const worksheetReader of reader) {
      for await (const row of worksheetReader) {
        if (row.number !== 1) continue;
        const value = row.getCell(1).value as { richText: { text: string }[] };
        expect(value.richText.map((r) => r.text).join('')).toBe('This is a colorful text with in-cell format');
        expect(row.getCell(1).type).toBe(ExcelJS.ValueType.RichText);
      }
    }
  });

  it('streams worksheets and rows written by Workbook (added for xlvane)', async () => {
    await withTempDir(async (dir) => {
      const filename = join(dir, 'wb.test.xlsx');
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('first');
      ws.addRow(['name', 'total']);
      ws.addRow(['a', 1]);
      ws.addRow(['b', 2]);
      wb.addWorksheet('second').addRow([true]);
      await wb.xlsx.writeFile(filename);

      const seen: { sheet: string; number: number; values: unknown[] }[] = [];
      const reader = new ExcelJS.stream.xlsx.WorkbookReader(filename, {});
      for await (const worksheetReader of reader) {
        for await (const row of worksheetReader) {
          seen.push({ sheet: worksheetReader.name, number: row.number, values: row.values });
        }
      }
      expect(seen.map((r) => [r.sheet, r.number])).toEqual([
        ['first', 1],
        ['first', 2],
        ['first', 3],
        ['second', 1],
      ]);
      expect(seen[1]?.values).toEqual([undefined, 'a', 1]);
      expect(seen[3]?.values).toEqual([undefined, true]);
    });
  });

  for (const name of [
    'Serialise: xlsx file (createTestBook)',
    '#readFile Row limit: should bail out if the file contains more rows than the limit',
    '#readFile Row limit: should fail fast on a huge file',
    '#readFile Row limit: should parse fine if the limit is not exceeded',
    '#readFile Column limit: should bail out if the file contains more cells than the limit',
    '#readFile Column limit: should fail fast on a huge file',
    '#readFile Column limit: should parse fine if the limit is not exceeded',
    '#read Row limit: should bail out if the file contains more rows than the limit',
    '#read Row limit: should parse fine if the limit is not exceeded',
    'edit styles in existing file: edit styles of single row instead of all',
    'regular formula: should be classified as a formula cell',
    'regular formula: should have text corresponding to the evaluated formula result',
    'regular formula: should have the formula source',
    'hyperlinked formula: should be classified as a formula cell',
    'hyperlinked formula: should have text corresponding to the evaluated formula result',
    'hyperlinked formula: should have the formula source',
    'hyperlinked formula: should contain the linked url',
    'shared string with an escaped underscore: should decode the underscore',
    'XML parse error in a worksheet: should reject the promise with the sax error',
    'missing some files in the zip container: should not break',
    "image tl anchor: integer part of col equals nativeCol",
    "image tl anchor: integer part of row equals nativeRow",
    "image tl anchor: anchor width equals column width when custom",
    "image tl anchor: anchor height equals row height",
    "image br anchor: integer part of col equals nativeCol",
    "image br anchor: integer part of row equals nativeRow",
    "image br anchor: anchor width equals column width when custom",
    "image br anchor: anchor height equals row height",
    'defined name that looks like a range: should not crash',
  ]) {
    it.skip(`${name} — fixture from spec/integration/data not copied`, () => {});
  }
});
