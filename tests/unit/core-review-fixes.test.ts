// Functional fixes from the pre-release review (docs/decisions.md D-R03): tables kept in step
// with appendRows and header edits, error locations, merged-cell writes, workbook windows.
import { describe, expect, it } from 'vitest';
import { appendRows, editCells } from '../../packages/xlvane/src/edit';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { insertRows } from '../../packages/xlvane/src/rows';
import { openWorkbookStream } from '../../packages/xlvane/src/stream-read';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';

function caught(fn: () => unknown): { code: string; location: Record<string, unknown>; hint: string; message: string } {
  try {
    fn();
  } catch (error) {
    return error as never;
  }
  throw new Error('expected an error');
}

function tableBook(totalsRow = false): Uint8Array {
  const wb = createWorkbook();
  wb.addSheet('S').addTable({
    name: 'Sales',
    ref: 'A1',
    totalsRow,
    columns: [{ name: 'Region', ...(totalsRow ? { totalsRowLabel: 'Total' } : {}) }, { name: 'Revenue', ...(totalsRow ? { totalsRowFunction: 'sum' as const } : {}) }],
    rows: [
      ['North', 1],
      ['South', 2],
    ],
  });
  return writeWorkbook(wb);
}

describe('appendRows and tables', () => {
  it('grows a table that ends on the last used row, with its filter', () => {
    const { bytes } = appendRows(tableBook(), 'S', [{ Region: 'East', Revenue: 3 }, { Region: 'West', Revenue: 4 }]);
    const sheet = readWorkbook(bytes).getSheet('S');
    expect(sheet.tables[0]?.ref).toBe('A1:B5');
    expect(sheet.tables[0]?.autoFilter).toBe('A1:B5');
    expect(sheet.cell('A5')?.value).toBe('West');
  });

  it('refuses a table with a totals row and points to insertRows', () => {
    const input = tableBook(true);
    const error = caught(() => appendRows(input, 'S', [['East', 3]]));
    expect(error.code).toBe('XLV_E_NOT_SUPPORTED');
    expect(error.location).toMatchObject({ sheet: 'S', cell: 'A4' });
    expect(error.hint).toContain('insertRows(bytes, { sheet: "S", at: 4, count: 1 })');
    // The hint works: the table grows above its totals row.
    const sheet = readWorkbook(insertRows(input, { sheet: 'S', at: 4, count: 1 }).bytes).getSheet('S');
    expect(sheet.tables[0]?.ref).toBe('A1:B5');
  });

  it('leaves a table alone that ends above the last used row', () => {
    const wb = createWorkbook();
    wb.addSheet('S').addTable({ name: 'T', ref: 'A1', columns: [{ name: 'a' }], rows: [[1]] }).set('A5', 'note');
    const { bytes } = appendRows(writeWorkbook(wb), 'S', [['x']]);
    expect(readWorkbook(bytes).getSheet('S').tables[0]?.ref).toBe('A1:A2');
  });
});

describe('editCells on table headers', () => {
  it('renames the table column with the header cell', () => {
    const { bytes } = editCells(tableBook(), [{ sheet: 'S', cell: 'B1', value: 'Income' }]);
    const table = readWorkbook(bytes).getSheet('S').tables[0];
    expect(table?.columns.map((c) => c.name)).toEqual(['Region', 'Income']);
  });

  it('refuses a header that is not text, empty or a duplicate name', () => {
    const input = tableBook();
    for (const value of [5, null, '', 'region']) {
      const error = caught(() => editCells(input, [{ sheet: 'S', cell: 'B1', value }]));
      expect(error.code, String(value)).toBe('XLV_E_INVALID_VALUE');
      expect(error.location).toMatchObject({ sheet: 'S', cell: 'B1' });
    }
  });

  it('allows swapping two header names in one call', () => {
    const { bytes } = editCells(tableBook(), [
      { sheet: 'S', cell: 'A1', value: 'Revenue' },
      { sheet: 'S', cell: 'B1', value: 'Region' },
    ]);
    expect(readWorkbook(bytes).getSheet('S').tables[0]?.columns.map((c) => c.name)).toEqual(['Revenue', 'Region']);
  });
});

describe('addTable calculated columns', () => {
  it('fills data cells given as null as well as missing ones', () => {
    const wb = createWorkbook();
    wb.addSheet('S').addTable({
      name: 'T',
      ref: 'A1',
      columns: [{ name: 'a' }, { name: 'b', calculatedColumnFormula: 'T[[#This Row],[a]]*2' }],
      rows: [[1, null], [2], [3, 9]],
    });
    const sheet = readWorkbook(writeWorkbook(wb)).getSheet('S');
    expect(sheet.cell('B2')?.formula?.text).toBe('T[[#This Row],[a]]*2');
    expect(sheet.cell('B3')?.formula?.text).toBe('T[[#This Row],[a]]*2');
    expect(sheet.cell('B4')?.value).toBe(9);
  });
});

describe('error locations', () => {
  const sheet = () => createWorkbook().addSheet('Data');
  const bad = { font: { size: -1 } } as never;

  it('names the sheet and cell, range, column or row of a bad style', () => {
    expect(caught(() => sheet().set('B2', 1, bad))).toMatchObject({ code: 'XLV_E_INVALID_STYLE', location: { sheet: 'Data', cell: 'B2' } });
    expect(caught(() => sheet().style('a1:b2', bad))).toMatchObject({ code: 'XLV_E_INVALID_STYLE', location: { sheet: 'Data', cell: 'A1:B2' } });
    expect(caught(() => sheet().column('C', { style: bad }))).toMatchObject({ code: 'XLV_E_INVALID_STYLE', location: { sheet: 'Data', cell: 'C:C' } });
    expect(caught(() => sheet().tabColor('nope'))).toMatchObject({ location: { sheet: 'Data' } });
  });

  it('names the column or row of a bad width or height', () => {
    expect(caught(() => sheet().column(3, { width: -1 }))).toMatchObject({ code: 'XLV_E_INVALID_VALUE', location: { sheet: 'Data', cell: 'C:C' } });
    expect(caught(() => sheet().row(7, { height: 1e6 }))).toMatchObject({ code: 'XLV_E_INVALID_VALUE', location: { sheet: 'Data', cell: '7:7' } });
  });

  it('names the sheet of other sheet options', () => {
    expect(caught(() => sheet().row(2, { hidden: 'yes' as never }))).toMatchObject({ code: 'XLV_E_INVALID_OPTION', location: { sheet: 'Data' } });
    expect(caught(() => sheet().freeze({ rows: -1 }))).toMatchObject({ location: { sheet: 'Data' } });
    expect(caught(() => sheet().view({ zoom: 5 }))).toMatchObject({ location: { sheet: 'Data' } });
    expect(caught(() => sheet().properties({ outlineLevelRow: 9 } as never))).toMatchObject({ location: { sheet: 'Data' } });
  });

  it('names the cell of a bad style in editCells', () => {
    expect(caught(() => editCells(tableBook(), [{ sheet: 'S', cell: 'C3', value: 1, style: bad }]))).toMatchObject({ code: 'XLV_E_INVALID_STYLE', location: { sheet: 'S', cell: 'C3' } });
  });

  it('suggests sheet names and explains case for an unknown sheet', async () => {
    const wb = createWorkbook();
    wb.addSheet('Data');
    const error = caught(() => wb.getSheet('data'));
    expect(error).toMatchObject({ code: 'XLV_E_SHEET_NOT_FOUND', location: { sheet: 'data' } });
    expect(error.hint).toContain('Did you mean "Data"?');
    expect(error.hint).toContain('case-sensitive');
    const stream = await openWorkbookStream(writeWorkbook(wb));
    const streamed = await (async () => {
      try {
        for await (const _ of stream.rows('DATA')) void _;
      } catch (e) {
        return e as { code: string; hint: string };
      }
      return null;
    })();
    expect(streamed?.code).toBe('XLV_E_SHEET_NOT_FOUND');
    expect(streamed?.hint).toContain('Did you mean "Data"?');
    expect(caught(() => readWorkbook(writeWorkbook(wb)).getSheet('Dta')).hint).toContain('Did you mean "Data"?');
  });
});

describe('editCells into merged ranges', () => {
  it('warns when a value goes into a cell other than the top-left one', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', 'title').merge('A1:C1');
    const input = writeWorkbook(wb);
    const { warnings } = editCells(input, [{ sheet: 'S', cell: 'B1', value: 'hidden' }]);
    expect(warnings.map((w) => w.code)).toContain('XLV_W_MERGED_CELL_WRITE');
    expect(warnings.find((w) => w.code === 'XLV_W_MERGED_CELL_WRITE')).toMatchObject({ location: { sheet: 'S' }, cells: ['B1'] });
    expect(editCells(input, [{ sheet: 'S', cell: 'A1', value: 'new' }]).warnings.map((w) => w.code)).not.toContain('XLV_W_MERGED_CELL_WRITE');
    expect(editCells(input, [{ sheet: 'S', cell: 'B1', value: null }]).warnings.map((w) => w.code)).not.toContain('XLV_W_MERGED_CELL_WRITE');
  });
});

describe('workbook windows', () => {
  it('refuses activeTab or firstSheet beyond the last sheet', () => {
    for (const view of [{ activeTab: 2 }, { firstSheet: 5 }]) {
      const wb = createWorkbook().views([view]);
      wb.addSheet('A');
      wb.addSheet('B');
      const error = caught(() => writeWorkbook(wb));
      expect(error.code).toBe('XLV_E_INVALID_OPTION');
      expect(error.hint).toContain('0 = "A"');
    }
    const ok = createWorkbook().views([{ activeTab: 1, firstSheet: 1 }]);
    ok.addSheet('A');
    ok.addSheet('B');
    expect(readWorkbook(writeWorkbook(ok)).views[0]?.activeTab).toBe(1);
  });
});
