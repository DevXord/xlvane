// Regressions found by testing the packed @xlvane/compat against exceljs 4.4.0 in outside
// projects before the first release (docs/decisions.md D-R09). exceljs behaviour is cited from the
// npm package exceljs 4.4.0 (lib/...).
import { describe, expect, it } from 'vitest';
import { ExcelJS } from './helpers';

async function reload(wb: any): Promise<any> {
  const back = new ExcelJS.Workbook();
  await back.xlsx.load(await wb.xlsx.writeBuffer());
  return back;
}

describe('worksheet.columns covers every used column (lib/doc/row.js L48, worksheet.js L243-261)', () => {
  it('after addRow, so the usual auto-fit loop sets widths that are written', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('s');
    ws.addRow(['Name', 'Description of item', 'Qty']);
    ws.addRow(['Widget', 'A long description text here', 5]);
    expect(ws.columns.length).toBe(3);
    ws.columns.forEach((col: any) => {
      let max = 0;
      col.eachCell({ includeEmpty: true }, (cell: any) => {
        max = Math.max(max, String(cell.value ?? '').length);
      });
      col.width = max + 2;
    });
    const back = (await reload(wb)).getWorksheet('s');
    expect([1, 2, 3].map((i) => back.getColumn(i).width)).toEqual([8, 30, 5]);
  });

  it('on a loaded file without <cols>, and default columns are not written', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('t').getCell('C3').value = 'abc';
    const back = await reload(wb);
    expect(back.getWorksheet('t').columns.length).toBe(3);
    const again = await reload(back);
    expect(again.getWorksheet('t').getColumn(1).width).toBeUndefined();
  });

  it('a streamed sheet whose later rows use more columns keeps its layout', async () => {
    const chunks: Uint8Array[] = [];
    const stream = { write: (c: Uint8Array, cb?: () => void) => (chunks.push(c), cb?.(), true), end: (cb?: () => void) => cb?.() };
    const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ stream });
    const ws = wb.addWorksheet('s');
    ws.addRow([1]).commit();
    ws.addRow([1, 2, 3, 4]).commit();
    ws.commit();
    await wb.commit();
    const bytes = new Uint8Array(chunks.reduce((n, c) => n + c.byteLength, 0));
    let at = 0;
    for (const c of chunks) (bytes.set(c, at), (at += c.byteLength));
    const back = new ExcelJS.Workbook();
    await back.xlsx.load(bytes);
    expect(back.getWorksheet('s').getRow(2).values).toEqual([undefined, 1, 2, 3, 4]);
  });
});

describe('streaming WorkbookWriter stays linear in the number of rows', () => {
  async function timeRows(n: number): Promise<number> {
    const stream = { write: (_c: Uint8Array, cb?: () => void) => (cb?.(), true), end: (cb?: () => void) => cb?.() };
    const start = performance.now();
    const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ stream });
    const ws = wb.addWorksheet('a');
    for (let i = 1; i <= n; i++) ws.addRow([i, i * 2, 'text']).commit();
    await wb.commit();
    return performance.now() - start;
  }

  it('four times the rows take well under eight times as long', async () => {
    await timeRows(2000); // warm-up
    const small = Math.min(await timeRows(10_000), await timeRows(10_000));
    const large = Math.min(await timeRows(40_000), await timeRows(40_000));
    // Linear is about 4; the queue's Array#shift and the last-row scan over committed rows made it about 16.
    expect(large / small).toBeLessThan(8);
  }, 120_000);
});

describe('split view active cell (sheet-view-xform.js L81-85, L167-169)', () => {
  it('is written and read back', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('s', { views: [{ state: 'split', xSplit: 2000, ySplit: 3000, topLeftCell: 'G10', activeCell: 'B5' }] });
    const view = (await reload(wb)).getWorksheet('s').views[0];
    expect(view).toMatchObject({ state: 'split', xSplit: 2000, ySplit: 3000, activePane: 'bottomRight', activeCell: 'B5' });
  });
});

describe('single underline reads back as true (underline-xform.js L14-30, L40-45)', () => {
  it('in cell fonts, rich text and conditional-format fonts', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('u');
    ws.getCell('A1').value = 'x';
    ws.getCell('A1').font = { underline: true };
    ws.getCell('A2').value = 'y';
    ws.getCell('A2').font = { underline: 'single' };
    ws.getCell('A3').value = 'z';
    ws.getCell('A3').font = { underline: 'double' };
    ws.getCell('A4').value = { richText: [{ text: 'r', font: { underline: true } }] };
    ws.addConditionalFormatting({ ref: 'A1:A3', rules: [{ type: 'expression', formulae: ['TRUE'], style: { font: { underline: true } } }] });
    const back = (await reload(wb)).getWorksheet('u');
    expect(back.getCell('A1').font.underline).toBe(true);
    expect(back.getCell('A2').font.underline).toBe(true);
    expect(back.getCell('A3').font.underline).toBe('double');
    expect(back.getCell('A4').value.richText[0].font.underline).toBe(true);
    expect(back.conditionalFormattings[0].rules[0].style.font.underline).toBe(true);
  });
});

describe('smaller exceljs shapes', () => {
  it('getSheetValues gives [] for an existing empty row (lib/doc/worksheet.js L571-579)', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('s');
    ws.getCell('A1').value = 1;
    ws.getRow(2);
    ws.getCell('A3').value = 3;
    const values = ws.getSheetValues();
    expect(values[2]).toEqual([]);
    expect(values[3]).toEqual([undefined, 3]);
  });

  it('orderNo as exceljs (lib/doc/workbook.js L81-85, L203)', async () => {
    const wb = new ExcelJS.Workbook();
    const a = wb.addWorksheet('A');
    const b = wb.addWorksheet('B');
    expect([a.orderNo, b.orderNo]).toEqual([1, 2]);
    wb.removeWorksheet(a.id);
    expect(wb.addWorksheet('C').orderNo).toBe(3);
    b.orderNo = 10;
    expect(wb.worksheets.map((w: any) => w.name)).toEqual(['C', 'B']);
    const back = await reload(wb);
    expect(back.worksheets.map((w: any) => [w.name, w.orderNo])).toEqual([
      ['C', 0],
      ['B', 1],
    ]);
  });

  it('a table column without a totals function reads back without totalsRowFunction (table-column-xform.js L29)', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('t').addTable({ name: 'T', ref: 'A1', totalsRow: true, columns: [{ name: 'a', totalsRowLabel: 'Sum' }, { name: 'b', totalsRowFunction: 'sum' }, { name: 'c' }], rows: [[1, 2, 3]] });
    const columns = (await reload(wb)).getWorksheet('t').getTable('T').table.columns;
    expect(columns[1].totalsRowFunction).toBe('sum');
    expect(columns[2].totalsRowFunction).toBeUndefined();
  });
});
