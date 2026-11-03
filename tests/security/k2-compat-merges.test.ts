// @xlvane/compat merged ranges with a styled top-left cell after the K2 penetration test: a 2.4 KB
// file with <mergeCell ref="A1:E400000"/> (2 000 000 positions, within the load budget) and a
// styled A1 loaded in seconds, but xlsx.writeBuffer then took 15 s and 1.4 GB, and ran out of
// memory at about 4 000 000 positions. Each merged cell carried a copy of the master's style, so
// each became a separate style conversion and a separate styled cell write. Merged ranges whose
// cells all carry the master's style are now written with one range style per block, and
// xlsx.load accepts merged ranges of at most max(1 048 576, cells read) positions in total.
// From 1.3.1 (F2) xlsx.load merges without styles, as exceljs does (lib/doc/worksheet.js L895-901
// `_parseMergeCells`), so a loaded merged cell keeps its own style from the file; cells that the
// file leaves out get the row and column style. A column style can still give every merged cell
// the master's style, and such ranges must still be written in blocks.
import { describe, expect, it, vi } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { NewSheet } from '../../packages/xlvane/src/write';
import { buildXlsx } from '../helpers/xlsx-builder';
import { ExcelJS } from '../exceljs/helpers';

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>
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
  it('a loaded merge of 1 000 000 positions is written quickly, the merged cells with their own style', async () => {
    const bytes = mergedFile(200_000);
    expect(bytes.length).toBeLessThan(5000);
    const wb = await new ExcelJS.Workbook().xlsx.load(bytes);
    const start = performance.now();
    const out = (await wb.xlsx.writeBuffer()) as Uint8Array;
    const ms = performance.now() - start;
    expect(ms).toBeLessThan(15000);
    const read = readWorkbook(out).sheets[0]!;
    expect(read.merges).toEqual(['A1:E200000']);
    // exceljs 4.4.0 writes the merged cells of this file as <c r="B1"/> without a style (F2).
    expect(writtenFonts(out, ['A1', 'B1', 'A2', 'E200000'])).toEqual({ A1: true });
  }, 120_000);

  // The same file with a bold column style on A:E: every merged cell is created with the column
  // style, equal to the master's but not the same object (exceljs, lib/doc/cell.js L27, gives a
  // new cell the row and column style too).
  function columnStyledFile(rows: number, extra = ''): Uint8Array {
    return buildXlsx({
      sheets: [
        {
          name: 'S',
          before: '<cols><col min="1" max="5" width="9" style="1"/></cols>',
          data: `<row r="1"><c r="A1" s="1"><v>1</v></c>${extra}</row>`,
          after: `<mergeCells count="1"><mergeCell ref="A1:E${String(rows)}"/></mergeCells>`,
        },
      ],
      styles: STYLES,
    });
  }

  /** Writes a workbook, expecting at most `max` styled cell writes (counted, not timed, as in K3). */
  async function countedWrite(wb: any, max: number): Promise<Uint8Array> {
    const set = vi.spyOn(NewSheet.prototype, 'set');
    try {
      const out = (await wb.xlsx.writeBuffer()) as Uint8Array;
      expect(set.mock.calls.length).toBeLessThanOrEqual(max);
      return out;
    } finally {
      set.mockRestore();
    }
  }

  it('a loaded merge whose cells take a column style equal to the master is written in blocks', async () => {
    const bytes = columnStyledFile(200_000);
    expect(bytes.length).toBeLessThan(5000);
    const wb = await new ExcelJS.Workbook().xlsx.load(bytes);
    expect(wb.getWorksheet('S').getCell('E200000').font?.bold).toBe(true);
    const out = await countedWrite(wb, 10);
    expect(writtenFonts(out, ['A1', 'B1', 'A2', 'E200000'])).toEqual({ A1: true, B1: true, A2: true, E200000: true });
  }, 120_000);

  it('a loaded merge with one cell of another style is still written quickly', async () => {
    // B1 has a style of its own in the file; the other 999 998 merged cells equal the master.
    const bytes = columnStyledFile(200_000, '<c r="B1" s="2"/>');
    const wb = await new ExcelJS.Workbook().xlsx.load(bytes);
    // Only the cells of row 1 are written one by one; rows 2 to 200 000 are styled as a block.
    const out = await countedWrite(wb, 10);
    const fonts = writtenFonts(out, ['A1', 'B1', 'A2', 'E200000']);
    expect(fonts).toEqual({ A1: true, B1: false, A2: true, E200000: true });
  }, 120_000);

  it('a loaded merge whose cells take a row style equal to the master is written in blocks', async () => {
    // From 1.3.1 (F3) row styles are read and new cells take them, as in exceljs: 60 styled rows
    // of a few bytes each give 983 040 merged cells the master's style.
    const rows = Array.from({ length: 60 }, (_, i) => `<row r="${String(i + 1)}" s="1" customFormat="1">${i === 0 ? '<c r="A1" s="1"><v>1</v></c>' : ''}</row>`).join('');
    const bytes = buildXlsx({ sheets: [{ name: 'S', before: '<dimension ref="A1:XFD60"/>', data: rows, after: '<mergeCells count="1"><mergeCell ref="A1:XFD60"/></mergeCells>' }], styles: STYLES });
    expect(bytes.length).toBeLessThan(5000);
    const wb = await new ExcelJS.Workbook().xlsx.load(bytes);
    expect(wb.getWorksheet('S').getCell('XFD60').font?.bold).toBe(true);
    const out = await countedWrite(wb, 10);
    expect(writtenFonts(out, ['A1', 'B1', 'A60', 'XFD60'])).toEqual({ A1: true, B1: true, A60: true, XFD60: true });
  }, 120_000);

  // Review of 1.3.1: merged cells that take a column or row style different from the master's
  // are written one by one, and each, having a style object of its own, was converted on its own
  // (twice the time of 1.3.0 for A1:B524288 with columns A and B styled differently). Styles
  // are now converted once per distinct set of parts: counted as the distinct style objects the
  // core is given, not timed.
  async function distinctStyles(wb: any): Promise<number> {
    const set = vi.spyOn(NewSheet.prototype, 'set');
    try {
      await wb.xlsx.writeBuffer();
      return new Set(set.mock.calls.map((c) => c[2])).size;
    } finally {
      set.mockRestore();
    }
  }

  it('merged cells with a column style other than the master are converted once per style', async () => {
    const bytes = buildXlsx({
      sheets: [{ name: 'S', before: '<cols><col min="1" max="1" width="9" style="1"/><col min="2" max="2" width="9" style="2"/></cols>', data: '<row r="1"><c r="A1" s="1"><v>1</v></c></row>', after: '<mergeCells count="1"><mergeCell ref="A1:B20000"/></mergeCells>' }],
      styles: STYLES,
    });
    const wb = await new ExcelJS.Workbook().xlsx.load(bytes);
    expect(wb.getWorksheet('S').getCell('B7').numFmt).toBe('0.00');
    expect(await distinctStyles(wb)).toBeLessThanOrEqual(3);
    const sheet = readWorkbook((await wb.xlsx.writeBuffer()) as Uint8Array).sheets[0]!;
    const cells = new Map([...sheet.cells(), ...sheet.styledCells()].map((c) => [c.ref, c]));
    expect(cells.get('B20000')?.numFmt).toBe('0.00');
    expect(cells.get('A20000')?.style.font?.bold).toBe(true);
  }, 120_000);

  it('merged cells with a row style other than the master are converted once per style', async () => {
    const rows = Array.from({ length: 20 }, (_, i) => `<row r="${String(i + 1)}" s="2" customFormat="1">${i === 0 ? '<c r="A1" s="1"><v>1</v></c>' : ''}</row>`).join('');
    const bytes = buildXlsx({ sheets: [{ name: 'S', before: '<dimension ref="A1:CV20"/>', data: rows, after: '<mergeCells count="1"><mergeCell ref="A1:CV20"/></mergeCells>' }], styles: STYLES });
    const wb = await new ExcelJS.Workbook().xlsx.load(bytes);
    expect(wb.getWorksheet('S').getCell('CV20').numFmt).toBe('0.00');
    expect(await distinctStyles(wb)).toBeLessThanOrEqual(3);
    const sheet = readWorkbook((await wb.xlsx.writeBuffer()) as Uint8Array).sheets[0]!;
    const cells = new Map([...sheet.cells(), ...sheet.styledCells()].map((c) => [c.ref, c]));
    expect(cells.get('A1')?.style.font?.bold).toBe(true);
    expect(cells.get('B1')?.numFmt).toBe('0.00');
    expect(cells.get('CV20')?.numFmt).toBe('0.00');
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
    // Assigning `style` gives C2 an object of its own; `C2.border = ...` would change the style
    // object shared by the whole range, as in exceljs (F1, lib/doc/cell.js L79-81, L169-175).
    ws.getCell('C2').style = { ...ws.getCell('C2').style, border: { bottom: { style: 'thin' } } };
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
