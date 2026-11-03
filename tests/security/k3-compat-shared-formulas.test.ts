// @xlvane/compat shared formulas after the K3 review of the K2 fixes: stream.xlsx.WorkbookReader
// gives a shared-formula dependent as `{ sharedFormula: master, result }` (K2), and copying its
// rows into a stream.xlsx.WorkbookWriter (`target.addRow(row.values).commit()`) then failed on
// the first dependent with XLV_E_INVALID_VALUE ("A2 shares the formula of A1, which holds no
// shared formula master written before it"): the writer looks the master up in its own sheet,
// where its row was committed and dropped. Before K2 the reader gave `{ formula: 'B2*2' }` and the
// copy worked. A dependent from the reader now also carries its moved formula as a non-enumerable
// `formula`, built only when it is read, which the writer writes when the master is not there.
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { buildXlsx } from '../helpers/xlsx-builder';
import { ExcelJS } from '../exceljs/helpers';

const gc = (globalThis as { gc?: () => void }).gc;

/** A master in A1 (`text`, shared over A1:A(dependents + 1)) and its dependents below, with B values. */
function shared(text: string, dependents: number): Uint8Array {
  let data = `<row r="1"><c r="A1"><f t="shared" ref="A1:A${String(dependents + 1)}" si="0">${text}</f><v>2</v></c><c r="B1"><v>1</v></c></row>`;
  for (let r = 2; r <= dependents + 1; r++) data += `<row r="${String(r)}"><c r="A${String(r)}"><f t="shared" si="0"/><v>${String(2 * r)}</v></c><c r="B${String(r)}"><v>${String(r)}</v></c></row>`;
  return buildXlsx({ sheets: [{ name: 'S', data }] });
}

/** Copies every row of the first sheet of `bytes` with the streaming reader into a streaming writer. */
async function streamCopy(bytes: Uint8Array): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  const stream = { write: (c: Uint8Array, cb?: () => void) => (chunks.push(c), cb?.(), true), end: (cb?: () => void) => cb?.() };
  const writer = new ExcelJS.stream.xlsx.WorkbookWriter({ stream });
  const target = writer.addWorksheet('Copy');
  for await (const ws of new ExcelJS.stream.xlsx.WorkbookReader(bytes, {})) {
    for await (const row of ws) target.addRow(row.values).commit();
  }
  target.commit();
  await writer.commit();
  const out = new Uint8Array(chunks.reduce((n, c) => n + c.byteLength, 0));
  let at = 0;
  for (const c of chunks) (out.set(c, at), (at += c.byteLength));
  return out;
}

describe('K3 item 2: copying shared formulas from the streaming reader to the streaming writer', () => {
  it('writes each dependent with its moved formula and result', async () => {
    const out = await streamCopy(shared('B1*2', 3));
    const cells = [...readWorkbook(out).sheets[0]!.cells()].filter((c) => c.ref.startsWith('A'));
    expect(cells.map((c) => [c.ref, c.formula?.text, c.value])).toEqual([
      ['A1', 'B1*2', 2],
      ['A2', 'B2*2', 4],
      ['A3', 'B3*2', 6],
      ['A4', 'B4*2', 8],
    ]);
  });

  it('also when the rows are committed only at the end (the master is still in the writer sheet)', async () => {
    const chunks: Uint8Array[] = [];
    const stream = { write: (c: Uint8Array, cb?: () => void) => (chunks.push(c), cb?.(), true), end: (cb?: () => void) => cb?.() };
    const writer = new ExcelJS.stream.xlsx.WorkbookWriter({ stream });
    const target = writer.addWorksheet('Copy');
    for await (const ws of new ExcelJS.stream.xlsx.WorkbookReader(shared('$B1*2', 2), {})) {
      for await (const row of ws) target.addRow(row.values);
    }
    target.commit();
    await writer.commit();
    const out = new Uint8Array(chunks.reduce((n, c) => n + c.byteLength, 0));
    let at = 0;
    for (const c of chunks) (out.set(c, at), (at += c.byteLength));
    const formulas = [...readWorkbook(out).sheets[0]!.cells()].filter((c) => c.ref.startsWith('A')).map((c) => c.formula?.text);
    expect(formulas).toEqual(['$B1*2', '$B2*2', '$B3*2']);
  });

  it('a dependent read from the stream reads as before: { sharedFormula, result }, with its formula on request', async () => {
    const values: Record<string, unknown>[] = [];
    for await (const ws of new ExcelJS.stream.xlsx.WorkbookReader(shared('B1*2', 2), {})) {
      for await (const row of ws) values.push(row.getCell(1).value);
    }
    expect(values[1]).toEqual({ sharedFormula: 'A1', result: 4 });
    expect(Object.keys(values[1]!)).toEqual(['sharedFormula', 'result']);
    expect(JSON.stringify(values[1])).toBe('{"sharedFormula":"A1","result":4}');
    expect(values[2]!['formula']).toBe('B3*2');
    // It can be set like any property.
    values[2]!['formula'] = 'C3';
    expect(values[2]).toEqual({ sharedFormula: 'A1', result: 6, formula: 'C3' });
  });

  it('a value copied into a loaded workbook is written as before', async () => {
    const wb = new ExcelJS.Workbook();
    const target = wb.addWorksheet('Copy');
    for await (const ws of new ExcelJS.stream.xlsx.WorkbookReader(shared('B1*2', 2), {})) {
      for await (const row of ws) target.addRow(row.values);
    }
    const back = await new ExcelJS.Workbook().xlsx.load(await wb.xlsx.writeBuffer());
    const sheet = back.getWorksheet('Copy');
    expect(sheet.getCell('A3').value).toEqual({ sharedFormula: 'A1', result: 6 });
    expect(sheet.getCell('A3').formula).toBe('B3*2');
  });

  it('reading alone builds no moved texts: a 1 000 000-character master shared with 200 cells', async () => {
    const master = Array.from({ length: 333_333 }, () => 'B1').join('+');
    const bytes = shared(master, 200);
    expect(bytes.length).toBeLessThan(20_000);
    const values: unknown[] = [];
    gc?.();
    const before = process.memoryUsage().heapUsed;
    const start = performance.now();
    for await (const ws of new ExcelJS.stream.xlsx.WorkbookReader(bytes, {})) {
      for await (const row of ws) {
        // What a caller commonly reads of a row and its cells.
        row.eachCell((cell: { type: unknown; text: unknown; effectiveType: unknown; formulaType: unknown; result: unknown }) => void [cell.type, cell.text, cell.effectiveType, cell.formulaType, cell.result]);
        JSON.stringify(row.values);
        values.push(row.getCell(1).value);
      }
    }
    const ms = performance.now() - start;
    gc?.();
    const grown = process.memoryUsage().heapUsed - before;
    expect(values).toHaveLength(201);
    expect(ms).toBeLessThan(2000);
    // 200 moved texts of 1 000 000 characters would be about 200 MB.
    expect(grown).toBeLessThan(50 * 1024 * 1024);
    // Each is still there when asked for.
    expect(((values[200] as Record<string, unknown>)['formula'] as string).startsWith('B201+B201+')).toBe(true);
  }, 60_000);
});
