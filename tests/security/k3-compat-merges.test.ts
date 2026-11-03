// @xlvane/compat merged ranges after the K3 review of the K2 fixes: K2 writes a merged range whose
// cells all carry the master's style as one range style per block of rows, but a note on any cell
// of the range made it fall back to a styled cell write per position (twice the time for a range
// of 200 000 positions). The cells of such a range are now styled as a block, and their notes are
// written on their own.
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { ExcelJS } from '../exceljs/helpers';

function merged(note: boolean): any {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').value = 1;
  ws.getCell('A1').font = { bold: true };
  ws.mergeCells('A1:E40000');
  if (note) {
    ws.getCell('C3').note = 'inside';
    ws.getCell('A1').note = 'master';
  }
  return wb;
}

async function writeMs(note: boolean): Promise<number> {
  const wb = merged(note);
  const start = performance.now();
  await wb.xlsx.writeBuffer();
  return performance.now() - start;
}

describe('K3 item 4: a note inside a merged range', () => {
  it('keeps the block writing: about as fast as the same range without a note', async () => {
    await writeMs(false);
    const plain = Math.min(await writeMs(false), await writeMs(false));
    const noted = Math.min(await writeMs(true), await writeMs(true));
    // Before: about twice as long (a styled cell write per merged position).
    expect(noted / plain).toBeLessThan(1.4);
  }, 120_000);

  it('writes the notes and the style of every merged cell', async () => {
    const out = (await merged(true).xlsx.writeBuffer()) as Uint8Array;
    const sheet = readWorkbook(out).sheets[0]!;
    expect(sheet.merges).toEqual(['A1:E40000']);
    const bold = new Map<string, boolean>();
    for (const c of [...sheet.cells(), ...sheet.styledCells()]) if (['A1', 'B1', 'C3', 'E40000'].includes(c.ref)) bold.set(c.ref, c.style.font?.bold === true);
    expect(Object.fromEntries(bold)).toEqual({ A1: true, B1: true, C3: true, E40000: true });
    const back = await new ExcelJS.Workbook().xlsx.load(out);
    const ws = back.getWorksheet('S');
    expect(ws.getCell('A1').note).toBe('master');
    expect(ws.getCell('C3').note).toBe('inside');
    expect(ws.getCell('C3').font).toMatchObject({ bold: true });
  });
});
