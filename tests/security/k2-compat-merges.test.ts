// @xlvane/compat merged ranges with a styled top-left cell after the K2 penetration test: a 2.4 KB
// file with <mergeCell ref="A1:E400000"/> (2 000 000 positions, within the load budget) and a
// styled A1 loaded in seconds, but xlsx.writeBuffer then took 15 s and 1.4 GB, and ran out of
// memory at about 4 000 000 positions. Each merged cell carries a copy of the master's style, as in
// exceljs, so each became a separate style conversion and a separate styled cell write. Merged
// ranges whose cells all carry the master's style are now written with one range style per block,
// and xlsx.load accepts merged ranges of at most max(1 048 576, cells read) positions in total.
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { buildXlsx } from '../helpers/xlsx-builder';
import { ExcelJS } from '../exceljs/helpers';

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>
</styleSheet>`;

function mergedFile(rows: number): Uint8Array {
  return buildXlsx({
    sheets: [{ name: 'S', data: '<row r="1"><c r="A1" s="1"><v>1</v></c></row>', after: `<mergeCells count="1"><mergeCell ref="A1:E${String(rows)}"/></mergeCells>` }],
    styles: STYLES,
  });
}

/** The style of each cell in the written file, by address, for the given addresses. */
function writtenFonts(bytes: Uint8Array, refs: string[]): Record<string, unknown> {
  const sheet = readWorkbook(bytes).sheets[0]!;
  const out: Record<string, unknown> = {};
  for (const c of sheet.styledCells()) if (refs.includes(c.ref)) out[c.ref] = c.style.font?.bold ?? false;
  for (const c of sheet.cells()) if (refs.includes(c.ref)) out[c.ref] = c.style.font?.bold ?? false;
  return out;
}

describe('K2 item 5: merged ranges with a styled master', () => {
  it('a loaded merge of 1 000 000 positions is written quickly, every merged cell with the style', async () => {
    const bytes = mergedFile(200_000);
    expect(bytes.length).toBeLessThan(5000);
    const wb = await new ExcelJS.Workbook().xlsx.load(bytes);
    const start = performance.now();
    const out = (await wb.xlsx.writeBuffer()) as Uint8Array;
    const ms = performance.now() - start;
    expect(ms).toBeLessThan(15000);
    const read = readWorkbook(out).sheets[0]!;
    expect(read.merges).toEqual(['A1:E200000']);
    expect(writtenFonts(out, ['A1', 'B1', 'A2', 'E200000'])).toEqual({ A1: true, B1: true, A2: true, E200000: true });
  }, 120_000);

  it('xlsx.load refuses merged ranges of more than 1 048 576 positions in a file with few cells', async () => {
    const bytes = mergedFile(400_000);
    const start = performance.now();
    const error = await new ExcelJS.Workbook().xlsx.load(bytes).then(
      () => null,
      (e: unknown) => e,
    );
    expect(performance.now() - start).toBeLessThan(6000);
    expect(error).toMatchObject({ code: 'XLV_E_LIMIT_EXCEEDED', details: { limit: 'mergedCells', value: 1_048_576 }, location: { sheet: 'S' } });
    expect((error as { hint?: string }).hint).toContain('ignoreNodes');
    // A merge over a whole column is still read.
    const column = buildXlsx({ sheets: [{ name: 'S', data: '<row r="1"><c r="A1"><v>1</v></c></row>', after: '<mergeCells count="1"><mergeCell ref="A1:A1048576"/></mergeCells>' }] });
    const wb = await new ExcelJS.Workbook().xlsx.load(column);
    expect(Object.keys(wb.getWorksheet('S').model.merges)).toHaveLength(1);
  }, 60_000);

  it('a merge made through the API with a styled master is written the same way', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('B2').value = 'x';
    ws.getCell('B2').font = { italic: true };
    ws.mergeCells('B2:D3');
    const out = (await wb.xlsx.writeBuffer()) as Uint8Array;
    const sheet = readWorkbook(out).sheets[0]!;
    const styled = new Map<string, boolean>();
    for (const c of [...sheet.cells(), ...sheet.styledCells()]) styled.set(c.ref, c.style.font?.italic === true);
    expect([...styled.entries()].sort()).toEqual(['B2', 'B3', 'C2', 'C3', 'D2', 'D3'].map((r) => [r, true]));
  });

  it('a merged cell whose style was changed keeps its own style', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 1;
    ws.getCell('A1').font = { bold: true };
    ws.mergeCells('A1:C2');
    ws.getCell('C2').border = { bottom: { style: 'thin' } };
    const out = (await wb.xlsx.writeBuffer()) as Uint8Array;
    const sheet = readWorkbook(out).sheets[0]!;
    const cells = new Map([...sheet.cells(), ...sheet.styledCells()].map((c) => [c.ref, c.style]));
    expect(cells.get('B1')?.font?.bold).toBe(true);
    expect(cells.get('C2')?.font?.bold).toBe(true);
    expect(cells.get('C2')?.border?.bottom?.style).toBe('thin');
    expect(cells.get('B2')?.border).toBeUndefined();
    // The merged cells share the master's cell format, as when each was written on its own.
    expect(sheet.cells().next().value?.styleIndex).toBe([...sheet.styledCells()].find((c) => c.ref === 'B1')?.styleIndex);
  });

  it('a date master gives its merged cells the date format also without styles', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = new Date(Date.UTC(2024, 0, 31));
    ws.mergeCells('A1:B2');
    const out = (await wb.xlsx.writeBuffer({ useStyles: false } as never)) as Uint8Array;
    const sheet = readWorkbook(out).sheets[0]!;
    const formats = new Map([...sheet.cells(), ...sheet.styledCells()].map((c) => [c.ref, c.numFmt]));
    expect(Object.fromEntries(formats)).toEqual({ A1: 'mm-dd-yy', B1: 'mm-dd-yy', A2: 'mm-dd-yy', B2: 'mm-dd-yy' });
  });
});
