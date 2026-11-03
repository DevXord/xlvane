// @xlvane/compat shared formulas on load after the K2 penetration test: xlsx.load and
// stream.xlsx.WorkbookReader read the text of every shared-formula dependent, which makes the core
// move the master's text to each of them. A 4.6 KB file (a 1 000 000-character master shared with
// 200 cells) ran out of heap; 30 dependents took 950 MB (the core alone, without the texts: 75 MB).
// Dependents are now read as exceljs's xlsx.load reads them, `{ sharedFormula: master, result }`
// (also by the streaming reader), and `cell.formula` moves the master's text when it is asked for.
import { describe, expect, it } from 'vitest';
import { buildXlsx } from '../helpers/xlsx-builder';
import { ExcelJS } from '../exceljs/helpers';

/** A master in A1 shared with A2…A(dependents + 1); the master's text has about `length` characters. */
function shared(length: number, dependents: number, extra = ''): Uint8Array {
  const master = Array.from({ length: Math.floor(length / 3) }, () => 'B1').join('+');
  let data = `<row r="1"><c r="A1"><f t="shared" ref="A1:A${String(dependents + 1)}" si="0">${master}</f><v>0</v></c></row>`;
  for (let r = 2; r <= dependents + 1; r++) data += `<row r="${String(r)}"><c r="A${String(r)}"><f t="shared" si="0"/><v>${String(r)}</v></c></row>`;
  return buildXlsx({ sheets: [{ name: 'S', data: data + extra }] });
}

describe('K2 item 4: shared formulas on load', () => {
  const bytes = shared(1_000_000, 40);

  it('xlsx.load does not move the master text to every dependent', async () => {
    expect(bytes.length).toBeLessThan(20_000);
    const start = performance.now();
    const wb = await new ExcelJS.Workbook().xlsx.load(bytes);
    const ms = performance.now() - start;
    const ws = wb.getWorksheet('S');
    expect(ws.getCell('A1').value).toMatchObject({ shareType: 'shared', ref: 'A1:A41', result: 0 });
    expect(ws.getCell('A1').value.formula.length).toBeGreaterThan(900_000);
    expect(ws.getCell('A2').value).toEqual({ sharedFormula: 'A1', result: 2 });
    expect(ws.getCell('A41').value).toEqual({ sharedFormula: 'A1', result: 41 });
    expect(ms).toBeLessThan(2000);
    // The moved text is still there when asked for.
    expect(ws.getCell('A3').formula.startsWith('B3+B3+')).toBe(true);
  }, 60_000);

  it('the loaded workbook is written again as a shared formula', async () => {
    const wb = await new ExcelJS.Workbook().xlsx.load(shared(30, 3));
    const again = await new ExcelJS.Workbook().xlsx.load(await wb.xlsx.writeBuffer());
    const ws = again.getWorksheet('S');
    expect(ws.getCell('A1').value).toEqual({ formula: 'B1+B1+B1+B1+B1+B1+B1+B1+B1+B1', shareType: 'shared', ref: 'A1:A4', result: 0 });
    expect(ws.getCell('A4').value).toEqual({ sharedFormula: 'A1', result: 4 });
    expect(ws.getCell('A4').formula).toBe('B4+B4+B4+B4+B4+B4+B4+B4+B4+B4');
  });

  it('a dependent the core cannot move to (outside the master range) keeps its cached value', async () => {
    const wb = await new ExcelJS.Workbook().xlsx.load(shared(9, 2, '<row r="9"><c r="A9"><f t="shared" si="0"/><v>7</v></c></row>'));
    expect(wb.getWorksheet('S').getCell('A9').value).toBe(7);
    expect(wb.getWorksheet('S').getCell('A2').value).toEqual({ sharedFormula: 'A1', result: 2 });
  });

  it('stream.xlsx.WorkbookReader gives dependents as { sharedFormula, result }, quickly', async () => {
    const start = performance.now();
    const values: unknown[] = [];
    let formula: string | undefined;
    for await (const ws of new ExcelJS.stream.xlsx.WorkbookReader(bytes, {})) {
      for await (const row of ws) {
        values.push(row.getCell(1).value);
        if (row.number === 3) formula = row.getCell(1).formula;
      }
    }
    expect(performance.now() - start).toBeLessThan(2000);
    expect(values).toHaveLength(41);
    // The master as exceljs's reader gives it: { formula, result }.
    expect(values[0]).toMatchObject({ result: 0 });
    expect(Object.keys(values[0] as object)).toEqual(['formula', 'result']);
    expect(values[1]).toEqual({ sharedFormula: 'A1', result: 2 });
    expect(values[40]).toEqual({ sharedFormula: 'A1', result: 41 });
    // The master's row is no longer in the worksheet, but its formula is kept for the dependents.
    expect(formula?.startsWith('B3+B3+')).toBe(true);
  }, 60_000);

  it('stream.xlsx.WorkbookReader keeps the cached value of a dependent outside the master range', async () => {
    const values: unknown[] = [];
    for await (const ws of new ExcelJS.stream.xlsx.WorkbookReader(shared(9, 2, '<row r="9"><c r="A9"><f t="shared" si="0"/><v>7</v></c></row>'), {})) {
      for await (const row of ws) values.push(row.getCell(1).value);
    }
    expect(values).toEqual([{ formula: 'B1+B1+B1', result: 0 }, { sharedFormula: 'A1', result: 2 }, { sharedFormula: 'A1', result: 3 }, 7]);
  });
});
