import { describe, expect, it } from 'vitest';
import { inspect } from '../../packages/xlvane/src/inspect';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { serialToDate } from '../../packages/xlvane/src/sml/dates';
import { isoDate } from '../../packages/xlvane/src/util/iso-date';
import { expectXlvError } from '../helpers/expect-error';
import { buildXlsx, c, inline, row } from '../helpers/xlsx-builder';

/** inspect() writes dates as ISO 8601 without a time zone, as CSV does: the date alone at midnight. */
const iso = (serial: number): string => isoDate(serialToDate(serial, '1900') as Date);

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

describe('inspect: phase A1 metadata (D-A07)', () => {
  it('reports set document properties, defined names, rich-text cells and styled columns', async () => {
    const { createWorkbook, writeWorkbook } = await import('../../packages/xlvane/src/write');
    const wb = createWorkbook({ properties: { title: 'Q3', created: new Date(Date.UTC(2026, 9, 1)) } });
    const s = wb.addSheet('Data');
    s.set('A1', { richText: [{ text: 'Bold', font: { bold: true } }, { text: ' plain' }] }).set('A2', 'x');
    s.column('B', { style: { font: { italic: true } } }).column('D', { width: 12 });
    wb.defineName('Rate', 'Data!$A$2', { sheet: 'Data', hidden: true });
    wb.defineName('Long', `SUM(${'Data!$A$1,'.repeat(20)}Data!$A$2)`);
    const r = inspect(writeWorkbook(wb));
    expect(r.properties).toEqual({ title: 'Q3', created: '2026-10-01T00:00:00.000Z' });
    expect(r.definedNames[0]).toEqual({ name: 'Rate', refersTo: 'Data!$A$2', sheet: 'Data', hidden: true });
    expect(r.definedNames[1]?.refersTo).toMatch(/^SUM\(.{96}…$/);
    expect(r.definedNamesTruncated).toBe(false);
    const sheet = r.sheets[0]!;
    expect(sheet.richTextCellCount).toBe(1);
    expect(sheet.richTextCells).toEqual(['A1']);
    expect(sheet.styledColumns).toEqual(['B']);
  });

  it('keeps the lists bounded', async () => {
    const { createWorkbook, writeWorkbook } = await import('../../packages/xlvane/src/write');
    const wb = createWorkbook();
    const s = wb.addSheet('Data');
    for (let i = 1; i <= 15; i++) s.set(`A${String(i)}`, { richText: [{ text: `r${String(i)}`, font: { bold: true } }] });
    for (let i = 0; i < 105; i++) wb.defineName(`n_${String(i)}`, 'Data!$A$1');
    const r = inspect(writeWorkbook(wb));
    expect(r.definedNames).toHaveLength(100);
    expect(r.definedNamesTruncated).toBe(true);
    expect(r.sheets[0]?.richTextCellCount).toBe(15);
    expect(r.sheets[0]?.richTextCells).toHaveLength(10);
    expect(r.properties).toEqual({});
  });
});

describe('header detection (D-R02)', () => {
  const header = (data: string) => inspect(buildXlsx({ sheets: [{ name: 'H', data }] })).sheets[0]?.headerRow ?? null;

  it('accepts consecutive years as column names', () => {
    const data = row(1, inline('A1', 'Region'), c('B1', 2023), c('C1', 2024)) + row(2, inline('A2', 'North'), c('B2', 15000), c('C2', 16000));
    expect(header(data)).toEqual({ row: 1, confidence: 'high' });
    const r = inspect(buildXlsx({ sheets: [{ name: 'H', data }] }));
    expect(r.sheets[0]?.columns.map((col) => col.header)).toEqual(['Region', '2023', '2024']);
  });

  it('does not take a data row with years for a header', () => {
    expect(header(row(1, inline('A1', 'Alice'), c('B1', 1985)) + row(2, inline('A2', 'Bob'), c('B2', 1990)))).toBeNull();
    expect(header(row(1, inline('A1', 'North'), c('B1', 1950), c('C1', 2010)) + row(2, inline('A2', 'South'), c('B2', 5), c('C2', 6)))).toBeNull();
  });

  it('ignores a summary formula next to the column names', () => {
    const data = row(1, inline('A1', 'Item'), inline('B1', 'Cost'), '<c r="C1"><f>SUM(B2:B3)</f><v>30</v></c>') + row(2, inline('A2', 'a'), c('B2', 10)) + row(3, inline('A3', 'b'), c('B3', 20));
    expect(header(data)).toEqual({ row: 1, confidence: 'high' });
  });

  it('accepts repeated names with low confidence', () => {
    const data = row(1, inline('A1', 'Amount'), inline('B1', 'Amount')) + row(2, c('A2', 1), c('B2', 2));
    expect(header(data)).toEqual({ row: 1, confidence: 'low' });
  });
});

describe('merged ranges in the report', () => {
  it('lists at most 100 and gives the count', () => {
    const merges = Array.from({ length: 150 }, (_, i) => `<mergeCell ref="A${String(2 * i + 1)}:B${String(2 * i + 1)}"/>`).join('');
    const r = inspect(buildXlsx({ sheets: [{ name: 'M', data: row(1, c('A1', 1)), after: `<mergeCells count="150">${merges}</mergeCells>` }] }));
    expect(r.sheets[0]?.mergedRanges).toHaveLength(100);
    expect(r.sheets[0]).toMatchObject({ mergedRangeCount: 150, mergedRangesTruncated: true });
  });
});
