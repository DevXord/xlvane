// Regression tests for the @xlvane/compat differences from exceljs 4.4.0 fixed in 1.3.1 (the
// list "Known differences planned for 1.3.1" of docs/exceljs-compat.md). Each expectation was
// checked by running exceljs 4.4.0 (bench/node_modules/exceljs) on the same input; the source
// lines cited are of that package. Test files are built here (tests/helpers/xlsx-builder.ts) or
// taken from the POI corpus (tests/corpus/LICENSES.md).
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { buildXlsx } from '../helpers/xlsx-builder';
import { ExcelJS } from './helpers';

const POI = fileURLToPath(new URL('../corpus/poi/', import.meta.url));

async function load(bytes: Uint8Array): Promise<any> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes);
  return wb;
}

/** Every cell of the first sheet of a written file, with its style, by reference. */
function writtenStyles(bytes: Uint8Array): Map<string, any> {
  const sheet = readWorkbook(bytes).sheets[0]!;
  return new Map<string, any>([...sheet.styledCells(), ...sheet.cells()].map((c) => [c.ref, c.style]));
}

const THIN = { style: 'thin' };
const RED = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } };

// Fonts 0-1, borders 0-4 (none, left+top, right+top, left+bottom, right+bottom), cell formats
// 0-5: 1 bold with border 1, 2-4 borders 2-4, 5 format 14 (a date).
const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="5"><border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"/><right/><top style="thin"/><bottom/><diagonal/></border>
<border><left/><right style="thin"/><top style="thin"/><bottom/><diagonal/></border>
<border><left style="thin"/><right/><top/><bottom style="thin"/><diagonal/></border>
<border><left/><right style="thin"/><top/><bottom style="thin"/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="6"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="3" xfId="0" applyBorder="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="4" xfId="0" applyBorder="1"/>
<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>
</styleSheet>`;

describe('F1: merged cells share the master style object', () => {
  // lib/doc/cell.js L169-175 `merge(master, ignoreStyle)`: `this.style = master.style`; the style
  // setters change that object in place (L55-97, for example L87-89 `this.style.fill = value`).
  it('a style set on any cell of a merged range after merging shows on all of them', () => {
    const ws = new ExcelJS.Workbook().addWorksheet('S');
    ws.getCell('A1').value = 'x';
    ws.getCell('B2').font = { bold: true };
    ws.mergeCells('A1:C3');
    expect(ws.getCell('B2').style).toBe(ws.getCell('A1').style);
    // B2's own font was replaced by the master's style, which has none.
    expect(ws.getCell('B2').font).toBeUndefined();
    ws.getCell('A1').fill = RED;
    ws.getCell('A1').border = { top: THIN };
    expect(ws.getCell('C3').fill).toEqual(RED);
    expect(ws.getCell('C3').border).toEqual({ top: THIN });
    // A setter on another cell of the range changes the shared object, so the master too.
    ws.getCell('B2').font = { italic: true };
    expect(ws.getCell('A1').font).toEqual({ italic: true });
    expect(ws.getCell('C3').font).toEqual({ italic: true });
    // Assigning `style` replaces the object of that cell only.
    ws.getCell('B2').style = { numFmt: '0.00' };
    expect(ws.getCell('A1').numFmt).toBeUndefined();
    expect(ws.getCell('C3').numFmt).toBeUndefined();
    expect(ws.getCell('B2').style).toEqual({ numFmt: '0.00' });
  });

  it('mergeCellsWithoutStyle leaves each cell its own style object', () => {
    const ws = new ExcelJS.Workbook().addWorksheet('S');
    ws.getCell('B1').font = { bold: true };
    ws.mergeCellsWithoutStyle('A1:B1');
    expect(ws.getCell('B1').style).not.toBe(ws.getCell('A1').style);
    ws.getCell('A1').fill = RED;
    expect(ws.getCell('B1').fill).toBeUndefined();
    expect(ws.getCell('B1').font).toEqual({ bold: true });
  });

  // lib/doc/cell.js L177-183 `unmerge`: a merged cell (not the master) gets the row and column
  // style again (`_mergeStyle(this._row.style, this._column.style, {})`); the master keeps its own.
  it('unMergeCells gives the other cells the row and column style again', () => {
    const ws = new ExcelJS.Workbook().addWorksheet('S');
    ws.getColumn(2).numFmt = '0.0';
    ws.mergeCells('A1:B2');
    ws.getCell('A1').fill = RED;
    ws.unMergeCells('A1:B2');
    expect(ws.getCell('A1').fill).toEqual(RED);
    expect(ws.getCell('A2').style).toEqual({});
    expect(ws.getCell('B2').style).toEqual({ numFmt: '0.0' });
    ws.getCell('A1').font = { bold: true };
    expect(ws.getCell('A2').font).toBeUndefined();
  });

  it('a fill and border set on the master after merging are written on every merged cell', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.mergeCells('B2:C3');
    ws.getCell('B2').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00FF00' } };
    ws.getCell('B2').border = { top: THIN, left: THIN, bottom: THIN, right: THIN };
    const styles = writtenStyles(await wb.xlsx.writeBuffer());
    for (const ref of ['B2', 'C2', 'B3', 'C3']) {
      expect(styles.get(ref)?.fill, ref).toMatchObject({ pattern: 'solid' });
      expect(styles.get(ref)?.border?.right?.style, ref).toBe('thin');
    }
    // Read back by the layer, each merged cell has the same style (exceljs gives the same).
    const back = (await load(await wb.xlsx.writeBuffer())).getWorksheet('S');
    expect(back.getCell('C3').fill).toEqual(back.getCell('B2').fill);
    expect(back.getCell('C3').border).toEqual(back.getCell('B2').border);
  });
});

describe('F2: xlsx.load merges without styles', () => {
  // lib/doc/worksheet.js L895-901 `_parseMergeCells` calls `mergeCellsWithoutStyle`, so each cell
  // of a loaded merged range keeps the style it has in the file.
  const box = buildXlsx({
    sheets: [
      {
        name: 'S',
        data: '<row r="2"><c r="B2" s="1"><v>1</v></c><c r="C2" s="2"/></row><row r="3"><c r="B3" s="3"/><c r="C3" s="4"/></row>',
        after: '<mergeCells count="1"><mergeCell ref="B2:C3"/></mergeCells>',
      },
    ],
    styles: STYLES,
  });

  it('a box border around a merged range, stored on its edge cells, is kept', async () => {
    const ws = (await load(box)).getWorksheet('S');
    expect(ws.getCell('C2').isMerged).toBe(true);
    expect(ws.getCell('B2').border).toEqual({ left: THIN, top: THIN });
    expect(ws.getCell('C2').border).toEqual({ right: THIN, top: THIN });
    expect(ws.getCell('B3').border).toEqual({ left: THIN, bottom: THIN });
    expect(ws.getCell('C3').border).toEqual({ right: THIN, bottom: THIN });
    expect(ws.getCell('C3').font?.bold).toBeUndefined();
  });

  it('the box border survives load and writeBuffer', async () => {
    const wb = await load(box);
    const styles = writtenStyles(await wb.xlsx.writeBuffer());
    expect(styles.get('B2')?.border).toMatchObject({ left: { style: 'thin' }, top: { style: 'thin' } });
    expect(styles.get('C2')?.border).toMatchObject({ right: { style: 'thin' }, top: { style: 'thin' } });
    expect(styles.get('B3')?.border).toMatchObject({ left: { style: 'thin' }, bottom: { style: 'thin' } });
    expect(styles.get('C3')?.border).toMatchObject({ right: { style: 'thin' }, bottom: { style: 'thin' } });
    expect(styles.get('C3')?.border?.top).toBeUndefined();
    expect(styles.get('C2')?.font?.bold).toBeUndefined();
  });

  it('a merged cell without a cell in the file has no style, as exceljs writes it', async () => {
    const bytes = buildXlsx({
      sheets: [{ name: 'S', data: '<row r="1"><c r="A1" s="1"><v>1</v></c></row>', after: '<mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells>' }],
      styles: STYLES,
    });
    const wb = await load(bytes);
    expect(wb.getWorksheet('S').getCell('B2').style).toEqual({});
    // exceljs writes <c r="B1"/>, <c r="A2"/>, <c r="B2"/> without a style.
    const styles = writtenStyles(await wb.xlsx.writeBuffer());
    expect(styles.get('A1')?.font?.bold).toBe(true);
    for (const ref of ['B1', 'A2', 'B2']) expect(styles.get(ref)?.font?.bold, ref).toBeUndefined();
  });

  it('POI corpus: merged cells keep their own borders and fonts', async () => {
    // FillWithoutColor.xlsx, sheet LPH: H6:H7 is merged; exceljs reads H7's border as left, right
    // and bottom (the master H6 has top, left, right).
    const lph = (await load(readFileSync(`${POI}FillWithoutColor.xlsx`))).getWorksheet('LPH');
    const h7 = lph.getCell('H7');
    expect(h7.isMerged).toBe(true);
    expect(Object.keys(h7.border).sort()).toEqual(['bottom', 'left', 'right']);
    // 54288-ref.xlsx, sheet MIDCON-PJM: N18:O18 is merged; exceljs reads O18's font as Arial 10,
    // not the master's bold italic 14. (O18 uses the workbook's default font, which the layer
    // reads as `undefined`: F4 of docs/exceljs-compat.md.)
    const pjm = (await load(readFileSync(`${POI}54288-ref.xlsx`))).getWorksheet('MIDCON-PJM');
    const o18 = pjm.getCell('O18');
    expect(o18.master.address).toBe('N18');
    expect(o18.font?.bold).toBeUndefined();
    expect(o18.font?.size).not.toBe(14);
    expect(o18.fill).toMatchObject({ pattern: 'solid', fgColor: { argb: 'FFFF0000' } });
    expect(pjm.getCell('N18').font).toMatchObject({ size: 14, bold: true, italic: true });
  });
});
