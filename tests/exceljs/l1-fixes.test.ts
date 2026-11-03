// Regression tests for the @xlvane/compat differences from exceljs 4.4.0 fixed in 1.3.1 (the
// list "Known differences planned for 1.3.1" of docs/exceljs-compat.md). Each expectation was
// checked by running exceljs 4.4.0 (bench/node_modules/exceljs) on the same input; the source
// lines cited are of that package. Test files are built here (tests/helpers/xlsx-builder.ts) or
// taken from the POI corpus (tests/corpus/LICENSES.md).
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { buildXlsx } from '../helpers/xlsx-builder';
import { ExcelJS, withTempDir } from './helpers';

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

/** The `<v>` of each cell of the first worksheet of a written file, by reference. */
function writtenValues(bytes: Uint8Array): Record<string, string> {
  const xml = strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']!);
  const out: Record<string, string> = {};
  for (const m of xml.matchAll(/<c r="([A-Z]+\d+)"[^>]*>(?:<f>[^<]*<\/f>)?<v>([^<]*)<\/v>/g)) out[m[1]!] = m[2]!;
  return out;
}

/** The values of column A of the first worksheet, read by xlsx.load and by the streaming reader. */
async function columnA(bytes: Uint8Array): Promise<{ load: unknown[]; stream: unknown[] }> {
  const ws = (await load(bytes)).worksheets[0];
  const loaded: unknown[] = [];
  ws.eachRow((row: any) => loaded.push(row.getCell(1).value));
  const streamed: unknown[] = [];
  for await (const sheet of new ExcelJS.stream.xlsx.WorkbookReader(bytes, {})) for await (const row of sheet as any) streamed.push(row.getCell(1).value);
  return { load: loaded, stream: streamed };
}

const iso = (v: unknown): unknown => (v instanceof Date ? v.toISOString() : v);

/** A file with the given serials in column A, each with the cell format `xf` of `numFmts`. */
function serialFile(serials: number[], numFmt: { id: number; code?: string }, workbookPr?: string): Uint8Array {
  const fmts = numFmt.code === undefined ? '' : `<numFmts count="1"><numFmt numFmtId="${String(numFmt.id)}" formatCode="${numFmt.code.replace(/"/g, '&quot;')}"/></numFmts>`;
  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${fmts}
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="${String(numFmt.id)}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>
</styleSheet>`;
  const data = serials.map((v, i) => `<row r="${String(i + 1)}"><c r="A${String(i + 1)}" s="1"><v>${String(v)}</v></c></row>`).join('');
  return buildXlsx({ sheets: [{ name: 'S', data }], styles, ...(workbookPr === undefined ? {} : { workbookPr }) });
}

describe('F6: time-only dates', () => {
  // exceljs writes a Date as 25569 + ms / 86400000 (lib/utils/utils.js L55-57 `dateToExcel`), so
  // the time 13:20 given as 1899-12-30T13:20Z is serial 0.5555...; checked with exceljs 4.4.0,
  // which writes <v>0.5555555555547471</v>, and for 1899-12-31T06:00 <v>1.25</v>.
  const time = new Date(Date.UTC(1899, 11, 30, 13, 20));

  it('a Date before 1900-01-01 is written as its serial, with the date format', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = time;
    ws.getCell('B1').value = time;
    ws.getCell('B1').numFmt = 'hh:mm';
    ws.getCell('C1').value = { formula: 'A1', result: time };
    ws.getCell('A2').value = new Date(Date.UTC(1899, 11, 30));
    ws.getCell('A3').value = new Date(Date.UTC(1899, 11, 31, 6));
    const bytes = (await wb.xlsx.writeBuffer()) as Uint8Array;
    const v = writtenValues(bytes);
    expect(Number(v['A1'])).toBeCloseTo(0.5555555555555556, 10);
    expect(Number(v['B1'])).toBeCloseTo(0.5555555555555556, 10);
    expect(Number(v['C1'])).toBeCloseTo(0.5555555555555556, 10);
    expect(v['A2']).toBe('0');
    expect(v['A3']).toBe('1.25');
    const sheet = readWorkbook(bytes).sheets[0]!;
    const formats = new Map([...sheet.cells()].map((c) => [c.ref, c.numFmt]));
    expect(formats.get('A1')).toBe('mm-dd-yy');
    expect(formats.get('B1')).toBe('hh:mm');
    expect(formats.get('C1')).toBe('mm-dd-yy');
    // Read back, the time is the same Date again.
    const back = (await load(bytes)).getWorksheet('S');
    expect(iso(back.getCell('A1').value)).toBe('1899-12-30T13:20:00.000Z');
    expect(iso(back.getCell('B1').value)).toBe('1899-12-30T13:20:00.000Z');
    expect(iso(back.getCell('C1').value.result)).toBe('1899-12-30T13:20:00.000Z');
  });

  it('the streaming writer writes a time-only Date too', async () => {
    await withTempDir(async (dir) => {
      const filename = join(dir, 'time.xlsx');
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename });
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = time;
      ws.getRow(1).commit();
      await wb.commit();
      const bytes = new Uint8Array(readFileSync(filename));
      expect(Number(writtenValues(bytes)['A1'])).toBeCloseTo(0.5555555555555556, 10);
      expect(iso((await load(bytes)).getWorksheet('S').getCell('A1').value)).toBe('1899-12-30T13:20:00.000Z');
    });
  });

  // exceljs reads a serial as 1899-12-30 plus the serial (lib/utils/utils.js L58-61
  // `excelToDate`): checked with exceljs 4.4.0 on this file, serials 0, 0.5, 0.25 read as
  // 1899-12-30T00:00, T12:00, T06:00; 1, 59, 60 as 1899-12-31, 1900-02-27, 1900-02-28 (the
  // difference for 1900-01-01 to 1900-02-28 is kept, see docs/exceljs-compat.md); 61 as 1900-03-01.
  it('a serial below 1 with a date format reads as 1899-12-30 plus the time', async () => {
    const { load: loaded, stream } = await columnA(serialFile([0, 0.5, 0.25, 1, 59, 61], { id: 14 }));
    const expected = ['1899-12-30T00:00:00.000Z', '1899-12-30T12:00:00.000Z', '1899-12-30T06:00:00.000Z', '1900-01-01T00:00:00.000Z', '1900-02-28T00:00:00.000Z', '1900-03-01T00:00:00.000Z'];
    expect(loaded.map(iso)).toEqual(expected);
    expect(stream.map(iso)).toEqual(expected);
  });

  // In the 1904 date system exceljs counts from 1904-01-01 (`excelToDate` with date1904 adds
  // 1462 days); checked with exceljs 4.4.0: 0.5 reads as 1904-01-01T12:00, -1 as 1903-12-31.
  it('the 1904 date system is unchanged', async () => {
    const { load: loaded, stream } = await columnA(serialFile([0, 0.5, 1, 61], { id: 14 }, 'date1904="1"'));
    const expected = ['1904-01-01T00:00:00.000Z', '1904-01-01T12:00:00.000Z', '1904-01-02T00:00:00.000Z', '1904-03-02T00:00:00.000Z'];
    expect(loaded.map(iso)).toEqual(expected);
    expect(stream.map(iso)).toEqual(expected);
    // A 1904 workbook still refuses a Date before 1904-01-01 (exceljs writes a negative serial).
    const wb = await load(serialFile([1], { id: 14 }, 'date1904="1"'));
    wb.getWorksheet('S').getCell('A1').value = time;
    await expect(wb.xlsx.writeBuffer()).rejects.toMatchObject({ code: 'XLV_E_INVALID_VALUE' });
  });
});

describe('F7: upper-case date formats', () => {
  // exceljs reads a number as a Date when its format, without quoted and bracketed parts,
  // matches /[ymdhMsb]+/ (lib/utils/utils.js L140-151 `isDateFmt`, used by
  // lib/xlsx/xform/sheet/cell-xform.js L449-453 and lib/stream/xlsx/worksheet-reader.js L329-334).
  it('POI corpus 64508.xlsx: DD.MM.YYYY reads as a Date', async () => {
    const bytes = new Uint8Array(readFileSync(`${POI}64508.xlsx`));
    const ws = (await load(bytes)).worksheets[0];
    // Checked with exceljs 4.4.0: C2 is 2020-06-08T00:00:00.000Z, D2 (#,###,##0) stays 1.
    expect(ws.getCell('C2').numFmt).toBe('DD.MM.YYYY');
    expect(iso(ws.getCell('C2').value)).toBe('2020-06-08T00:00:00.000Z');
    expect(ws.getCell('C2').type).toBe(ExcelJS.ValueType.Date);
    expect(ws.getCell('D2').value).toBe(1);
    let streamed: unknown;
    for await (const sheet of new ExcelJS.stream.xlsx.WorkbookReader(bytes, {})) {
      for await (const row of sheet as any) if (row.number === 2) streamed = row.getCell(3).value;
      break;
    }
    expect(iso(streamed)).toBe('2020-06-08T00:00:00.000Z');
  });

  it('YYYY-MM-DD and HH:MM read as dates in both date systems; quoted letters do not', async () => {
    // 45000 is 2023-03-15 in the 1900 date system and 2027-03-16 in the 1904 one.
    const iso1900 = await columnA(serialFile([45000, 0.5], { id: 164, code: 'YYYY-MM-DD' }));
    expect(iso1900.load.map(iso)).toEqual(['2023-03-15T00:00:00.000Z', '1899-12-30T12:00:00.000Z']);
    expect(iso1900.stream.map(iso)).toEqual(['2023-03-15T00:00:00.000Z', '1899-12-30T12:00:00.000Z']);
    const iso1904 = await columnA(serialFile([45000], { id: 164, code: 'YYYY-MM-DD' }, 'date1904="1"'));
    expect(iso1904.load.map(iso)).toEqual(['2027-03-16T00:00:00.000Z']);
    expect(iso1904.stream.map(iso)).toEqual(['2027-03-16T00:00:00.000Z']);
    const hours = await columnA(serialFile([0.75], { id: 164, code: 'HH:MM' }));
    expect(hours.load.map(iso)).toEqual(['1899-12-30T18:00:00.000Z']);
    for (const code of ['0.0 "MB"', '[Red]0.00', '#,##0.00 [$EUR]']) {
      const plain = await columnA(serialFile([2.5], { id: 164, code }));
      expect(plain.load, code).toEqual([2.5]);
      expect(plain.stream, code).toEqual([2.5]);
    }
  });

  it('a formula result with such a format reads as a Date (xlsx.load)', async () => {
    // cell-xform.js L456-459: a formula's result is converted when the format is a date format.
    const bytes = buildXlsx({
      sheets: [{ name: 'S', data: '<row r="1"><c r="A1" s="1"><f>40000+5000</f><v>45000</v></c></row>' }],
      styles: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="DD.MM.YYYY"/></numFmts>
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>
</styleSheet>`,
    });
    const value = (await load(bytes)).getWorksheet('S').getCell('A1').value;
    expect(value.formula).toBe('40000+5000');
    expect(iso(value.result)).toBe('2023-03-15T00:00:00.000Z');
  });
});

describe('F10: csv.read with latin1', () => {
  // fast-csv decodes with Node.js's StringDecoder (@fast-csv/parse/build/src/CsvParserStream.js
  // L20), whose 'latin1' (alias 'binary') maps each byte to the code point of the same value.
  // Checked with exceljs 4.4.0: bytes 80, 93, 9F read as U+0080, U+0093, U+009F (the WHATWG
  // 'latin1' label is windows-1252, which gives U+20AC, U+201C, U+0178).
  it('decodes each byte as the code point of the same value', async () => {
    const bytes = new Uint8Array([0x80, 0x2c, 0x93, 0x41, 0x2c, 0xe9, 0x0a, 0x9f, 0x2c, 0xff]);
    for (const encoding of ['latin1', 'binary']) {
      const ws = await new ExcelJS.Workbook().csv.read([bytes], { parserOptions: { encoding } });
      expect([ws.getCell('A1').value, ws.getCell('B1').value, ws.getCell('C1').value, ws.getCell('A2').value, ws.getCell('B2').value], encoding).toEqual(['\u0080', '\u0093A', 'é', '\u009f', 'ÿ']);
    }
    // A chunk boundary does not matter: one byte per chunk.
    const ws = await new ExcelJS.Workbook().csv.read([...bytes].map((b) => new Uint8Array([b])), { parserOptions: { encoding: 'latin1' } });
    expect(ws.getCell('B1').value).toBe('\u0093A');
  });

  it('a latin1 file with U+0080 to U+009F reads back unchanged', async () => {
    await withTempDir(async (dir) => {
      const file = join(dir, 'latin1.csv');
      const wb = new ExcelJS.Workbook();
      const text = `a${Array.from({ length: 32 }, (_, i) => String.fromCharCode(0x80 + i)).join('')}éÿ`;
      wb.addWorksheet('S').getCell('A1').value = text;
      await wb.csv.writeFile(file, { encoding: 'latin1' });
      expect(readFileSync(file).length).toBe(text.length);
      const back = await new ExcelJS.Workbook().csv.readFile(file, { parserOptions: { encoding: 'latin1' } });
      expect(back.getCell('A1').value).toBe(text);
    });
  });
});
