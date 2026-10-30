import { describe, expect, it } from 'vitest';
import { inspect } from '../../packages/xlvane/src/inspect';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { serialToDate } from '../../packages/xlvane/src/sml/dates';
import { expectXlvError } from '../helpers/expect-error';
import { buildXlsx, c, inline, row } from '../helpers/xlsx-builder';

/** inspect() writes dates as ISO 8601 without a time zone: YYYY-MM-DDTHH:mm:ss.sss */
const iso = (serial: number): string => (serialToDate(serial, '1900') as Date).toISOString().slice(0, 23);

const STYLES = '<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs>';

const report = buildXlsx({
  styles: STYLES,
  sharedStrings: ['Date', 'Customer', 'Amount', 'Paid', 'ACME', 'Globex', 'Quarterly report'],
  sheets: [
    {
      name: 'Sales',
      before: '<dimension ref="A1:D5"/>',
      data:
        row(1, c('A1', 6, 't="s"')) +
        row(2, c('A2', 0, 't="s"'), c('B2', 1, 't="s"'), c('C2', 2, 't="s"'), c('D2', 3, 't="s"')) +
        row(3, c('A3', 45292, 's="1"'), c('B3', 4, 't="s"'), c('C3', 100.5), c('D3', 1, 't="b"')) +
        row(4, c('A4', 45293, 's="1"'), c('B4', 5, 't="s"'), c('C4', 200), c('D4', 0, 't="b"')) +
        row(5, c('A5', 45294, 's="1"'), c('B5', 4, 't="s"'), '<c r="C5"><f>SUM(C3:C4)</f><v>300.5</v></c>'),
      after: '<mergeCells count="1"><mergeCell ref="A1:D1"/></mergeCells>',
    },
    { name: 'Notes', state: 'hidden', data: row(1, inline('A1', 'free text')) },
  ],
});

describe('inspect', () => {
  const result = inspect(report);

  it('describes the workbook', () => {
    expect(result).toMatchObject({
      schemaVersion: 1,
      format: 'xlsx',
      conformance: 'transitional',
      dateSystem: '1900',
      hasMacros: false,
      externalLinkCount: 0,
      sheetCount: 2,
    });
    expect(result.sheets.map((s) => [s.name, s.state])).toEqual([
      ['Sales', 'visible'],
      ['Notes', 'hidden'],
    ]);
  });

  it('describes ranges, counts and merged cells of a sheet', () => {
    expect(result.sheets[0]).toMatchObject({
      name: 'Sales',
      index: 0,
      kind: 'worksheet',
      dimension: 'A1:D5',
      usedRange: 'A1:D5',
      rowCount: 5,
      columnCount: 4,
      cellCount: 16,
      formulaCount: 1,
      mergedRanges: ['A1:D1'],
      formulas: [{ cell: 'C5', formula: 'SUM(C3:C4)' }],
    });
  });

  it('detects the header row below a title and infers column types', () => {
    const sheet = result.sheets[0];
    expect(sheet?.headerRow).toEqual({ row: 2, confidence: 'high' });
    expect(sheet?.columns).toEqual([
      {
        index: 1,
        letter: 'A',
        header: 'Date',
        dominantType: 'date',
        types: { number: 0, string: 0, boolean: 0, date: 3, error: 0, empty: 0 },
        samples: [iso(45292), iso(45293), iso(45294)],
      },
      {
        index: 2,
        letter: 'B',
        header: 'Customer',
        dominantType: 'string',
        types: { number: 0, string: 3, boolean: 0, date: 0, error: 0, empty: 0 },
        samples: ['ACME', 'Globex'],
      },
      {
        index: 3,
        letter: 'C',
        header: 'Amount',
        dominantType: 'number',
        types: { number: 3, string: 0, boolean: 0, date: 0, error: 0, empty: 0 },
        samples: [100.5, 200, 300.5],
      },
      {
        index: 4,
        letter: 'D',
        header: 'Paid',
        dominantType: 'boolean',
        types: { number: 0, string: 0, boolean: 2, date: 0, error: 0, empty: 1 },
        samples: [true, false],
      },
    ]);
  });

  it('returns plain JSON data', () => {
    expect(JSON.parse(JSON.stringify(result))).toEqual(result);
  });

  it('reports no header when the first rows are not all text', () => {
    const r = inspect(buildXlsx({ sheets: [{ name: 'N', data: row(1, c('A1', 1), c('B1', 2)) + row(2, c('A2', 3), c('B2', 4)) }] }));
    expect(r.sheets[0]?.headerRow).toBeNull();
    expect(r.sheets[0]?.columns[0]).toMatchObject({ header: null, dominantType: 'number', types: { number: 2 } });
  });

  it('gives low confidence to a text-only sheet', () => {
    const r = inspect(buildXlsx({ sheets: [{ name: 'T', data: row(1, inline('A1', 'a'), inline('B1', 'b')) + row(2, inline('A2', 'c'), inline('B2', 'd')) }] }));
    expect(r.sheets[0]?.headerRow).toEqual({ row: 1, confidence: 'low' });
  });

  it('describes an empty sheet', () => {
    const r = inspect(buildXlsx({ sheets: [{ name: 'Empty', data: '' }] }));
    expect(r.sheets[0]).toMatchObject({ usedRange: null, rowCount: 0, columnCount: 0, cellCount: 0, headerRow: null, columns: [] });
  });

  it('warns about cells outside the declared dimension and values hidden under merges', () => {
    const r = inspect(
      buildXlsx({
        sheets: [
          {
            name: 'W',
            before: '<dimension ref="A1"/>',
            data: row(1, c('A1', 1), c('B1', 2)) + row(2, c('C2', 3)),
            after: '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>',
          },
        ],
      }),
    );
    expect(r.warnings).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ code: 'XLV_W_DIMENSION_MISMATCH', location: { sheet: 'W' } }),
        expect.objectContaining({ code: 'XLV_W_VALUES_IN_MERGED_CELLS', cells: ['B1'] }),
      ]),
    );
  });

  it('caps the number of described columns', () => {
    const cells = Array.from({ length: 26 }, (_, i) => c(`${String.fromCharCode(65 + i)}1`, i));
    const r = inspect(buildXlsx({ sheets: [{ name: 'Wide', data: row(1, ...cells) }] }), { maxColumns: 5 });
    expect(r.sheets[0]?.columns).toHaveLength(5);
    expect(r.sheets[0]?.columnsTruncated).toBe(true);
  });

  it('accepts a workbook that was already read', () => {
    expect(inspect(readWorkbook(report)).sheets).toHaveLength(2);
  });

  it('includes reading warnings', () => {
    const r = inspect(buildXlsx({ vba: true }));
    expect(r.hasMacros).toBe(true);
    expect(r.warnings.map((w) => w.code)).toContain('XLV_W_MACROS_PRESENT');
  });

  it('rejects unknown options', () => {
    expectXlvError(() => inspect(report, { maxColumn: 3 } as never), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => inspect(report, { maxColumns: -1 }), 'XLV_E_INVALID_OPTION');
  });
});
