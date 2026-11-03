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

describe('F9: csv formatter options', () => {
  // exceljs passes formatterOptions to fast-csv's format (lib/csv/csv.js L118); a row is the
  // array of mapped values (L156-163), and a missing row is written as [] (L153-156). Unless
  // noted, each expected text is exceljs 4.4.0's output for the same worksheet and options (with
  // fast-csv 4.3.6, @fast-csv/format 4.3.5); the files cited are of @fast-csv/format/build/src.
  // The worksheet starts at row 1, where exceljs and xlvane write the same blank lines.
  function sheet(): any {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 'a';
    ws.getCell('B1').value = 'b,c';
    ws.getCell('C1').value = 3;
    ws.getCell('A3').value = 'x';
    ws.getCell('B3').value = 'y';
    return wb;
  }
  const text = async (wb: any, formatterOptions: Record<string, unknown>, more: Record<string, unknown> = {}): Promise<string> => new TextDecoder('utf-8', { ignoreBOM: true }).decode(await wb.csv.writeBuffer({ formatterOptions, ...more }));

  it('transform is applied to each row', async () => {
    // formatter/RowFormatter.js L22-24, L70-91: the transform gets the row and its result is
    // formatted; a falsy result writes no line.
    expect(await text(sheet(), { transform: (row: unknown[]) => row.map((v) => (v === 'a' ? 'A' : v)) })).toBe('A,"b,c",3\n\nx,y');
    expect(await text(sheet(), { transform: (row: unknown[]) => (row.length === 0 ? null : row) })).toBe('a,"b,c",3\nx,y');
    // A transform that does not take exactly one parameter calls back (RowFormatter.js L47-63;
    // types.js L5 `isSyncTransform`).
    expect(await text(sheet(), { transform: (row: unknown[], cb: (e: unknown, r: unknown) => void) => setTimeout(() => cb(null, [...row, 'z']), 0) })).toBe('a,"b,c",3,z\nz\nx,y,z');
    // An object row is written by the headers (RowFormatter.js L36-45, L126-131).
    expect(await text(sheet(), { headers: ['two', 'one'], transform: (row: unknown[]) => ({ one: row[0], two: row[1] }) })).toBe('two,one\n"b,c",a\n,\ny,x');
    expect(await text(sheet(), { transform: (row: unknown[]) => ({ one: row[0], two: row[1] }) })).toBe('a,"b,c"\n,\nx,y');
    // An error of the transform rejects, as in exceljs ("bad row"; for a transform that is not a
    // function exceljs rejects with a TypeError, RowFormatter.js L64-67).
    await expect(
      sheet().csv.writeBuffer({
        formatterOptions: {
          transform: (_row: unknown) => {
            throw new Error('bad row');
          },
        },
      }),
    ).rejects.toThrow('bad row');
    await expect(sheet().csv.writeBuffer({ formatterOptions: { transform: 'x' } })).rejects.toMatchObject({ code: 'XLV_E_INVALID_OPTION' });
  });

  it('a formula returned by transform is escaped like any other field (CSV injection)', async () => {
    // exceljs writes "=1+1,+cmd,@SUM(A1),-2+3" on each line and "=h", "=2+2" unescaped.
    const out = await text(sheet(), { transform: (_row: unknown) => ['=1+1', '+cmd', '@SUM(A1)', '-2+3'] });
    expect(out.split('\n')).toEqual(Array(3).fill('"\'=1+1","\'+cmd","\'@SUM(A1)","\'-2+3"'));
    const headers = await text(sheet(), { headers: ['=h'], transform: (_row: unknown) => ({ '=h': '=2+2' }) });
    expect(headers.split('\n')).toEqual(['"\'=h"', '"\'=2+2"', '"\'=2+2"', '"\'=2+2"']);
    // formulaInjection: "off" writes them as exceljs does.
    expect((await text(sheet(), { transform: (_row: unknown) => ['=1+1'] }, { formulaInjection: 'off' })).split('\n')[0]).toBe('=1+1');
  });

  it('quoteHeaders and quoteColumns, also as objects of header names', async () => {
    // FormatterOptions.js L12-13, L20-22: quoteHeaders defaults to quoteColumns;
    // formatter/FieldFormatter.js L24-36: an object is looked up by the column's header.
    const h = ['h1', 'h2'];
    // With headers, a missing row gives one empty field per header (RowFormatter.js L142-147).
    expect(await text(sheet(), { headers: h })).toBe('h1,h2\na,"b,c"\n,\nx,y');
    expect(await text(sheet(), { headers: h, quoteHeaders: true })).toBe('"h1","h2"\na,"b,c"\n,\nx,y');
    expect(await text(sheet(), { headers: h, quoteColumns: true })).toBe('"h1","h2"\n"a","b,c"\n"",""\n"x","y"');
    expect(await text(sheet(), { headers: h, quoteColumns: true, quoteHeaders: false })).toBe('h1,h2\n"a","b,c"\n"",""\n"x","y"');
    expect(await text(sheet(), { headers: ['h1', 'h2', 'h3'], quoteColumns: { h1: true } })).toBe('"h1",h2,h3\n"a","b,c",3\n"",,\n"x",y,');
    expect(await text(sheet(), { headers: h, quoteHeaders: { h2: true } })).toBe('h1,"h2"\na,"b,c"\n,\nx,y');
    expect(await text(sheet(), { headers: h, writeHeaders: false, quoteColumns: { h2: true } })).toBe('a,"b,c",3\n\nx,"y"');
    expect(await text(sheet(), { quoteColumns: [false, true] })).toBe('a,"b,c",3\n\nx,"y"');
    // Without headers, an object takes the header names from the first row written
    // (RowFormatter.js L109-124).
    expect(await text(sheet(), { quoteColumns: { a: true } })).toBe('"a","b,c",3\n\n"x",y');
    expect(await text(sheet(), { quoteColumns: { a: true } }, { includeEmptyRows: false })).toBe('"a","b,c",3\n"x",y');
  });

  it('headers: true, alwaysWriteHeaders, and a worksheet without rows', async () => {
    // headers: true writes the first row as the header line (RowFormatter.js L109-124), and later
    // rows get its length (L142-147).
    expect(await text(sheet(), { headers: true })).toBe('a,"b,c",3\n,,\nx,y,');
    expect(await text(sheet(), { headers: true }, { includeEmptyRows: false })).toBe('a,"b,c",3\nx,y,');
    expect(await text(sheet(), { headers: true, quoteHeaders: true }, { includeEmptyRows: false })).toBe('"a","b,c","3"\nx,y,');
    const empty = new ExcelJS.Workbook();
    empty.addWorksheet('S');
    // The header line is written with the first row only, or always with alwaysWriteHeaders
    // (RowFormatter.js L93-101); without header names that is an error (a plain Error in exceljs).
    expect(await text(empty, { headers: ['h1'] })).toBe('');
    expect(await text(empty, { headers: ['h1'], alwaysWriteHeaders: true })).toBe('h1');
    await expect(empty.csv.writeBuffer({ formatterOptions: { headers: true, alwaysWriteHeaders: true } })).rejects.toMatchObject({ code: 'XLV_E_INVALID_OPTION' });
    expect(await text(sheet(), { headers: ['h1'], alwaysWriteHeaders: true })).toBe('h1\na\n\nx');
  });

  it('a row of [header, value] pairs writes its values, escaped', async () => {
    // RowFormatter.js L26-31, L36-45, L133-141: when the first field of a row is an array of two,
    // the row is a list of pairs. exceljs writes "x\nx" for a map giving ['=1+1', 'x'].
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 'a';
    ws.getCell('A2').value = 'b';
    expect(new TextDecoder().decode(await wb.csv.writeBuffer({ map: () => ['=1+1', 'x'] }))).toBe('x\nx');
    expect(new TextDecoder().decode(await wb.csv.writeBuffer({ map: () => ['=1+1', '=2'] }))).toBe('"\'=2"\n"\'=2"');
  });

  it('headers: true takes the first row of the worksheet, also when it is not row 1', async () => {
    // Review of 1.3.1: the blank line written for a missing row 1 became the header line, so no
    // data was written. Checked with exceljs 4.4.0 (rows set with row.values): rows 2-3 give
    // "name,age\nann,5"; rows 3-4 give "\n\n" (exceljs writes one blank row, which becomes the
    // headers, and loses the data; xlvane writes the data); rows 1 and 3 give "name,age\n,\nann,5".
    const csv = async (rows: [number, unknown[]][]): Promise<string> => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('S');
      for (const [n, v] of rows) ws.getRow(n).values = v;
      return text(wb, { headers: true });
    };
    expect(await csv([[2, ['name', 'age']], [3, ['ann', 5]]])).toBe('name,age\nann,5');
    expect(await csv([[3, ['name', 'age']], [4, ['ann', 5]]])).toBe('name,age\nann,5');
    expect(await csv([[1, ['name', 'age']], [3, ['ann', 5]]])).toBe('name,age\n,\nann,5');
  });

  it('a BOM other than U+FEFF is not supported', async () => {
    // FormatterOptions.js L17-19: fast-csv takes `BOM` from the options too; exceljs writes "XX".
    await expect(sheet().csv.writeBuffer({ formatterOptions: { writeBOM: true, BOM: 'XX' } })).rejects.toMatchObject({ code: 'XLV_E_NOT_SUPPORTED' });
    expect(await text(sheet(), { writeBOM: true, BOM: '﻿' })).toBe('﻿a,"b,c",3\n\nx,y');
  });
});

describe('F12: column.eachCell and the column style setters', () => {
  // lib/doc/column.js L168-177: eachCell calls worksheet.eachRow (lib/doc/worksheet.js L551-568,
  // rows with values without includeEmpty) and gives the iteratee row.getCell(colNumber), which
  // creates the cell. The style setters go through eachCell (column.js L206-212 `_applyStyle`).
  // Checked with exceljs 4.4.0 on this worksheet: eachCell visits B1, B3, B5; values of A are
  // [, 1, , , , 'x']; a fill on column C reaches C1, C3 and C5 only.
  function sheet(): any {
    const ws = new ExcelJS.Workbook().addWorksheet('S');
    ws.getCell('A1').value = 1;
    ws.getCell('B3').value = 2;
    ws.getCell('A5').value = 'x';
    ws.getRow(4).height = 20;
    ws.getCell('B6').font = { bold: true };
    return ws;
  }

  it('visits the column cell of every row with values, creating it', () => {
    const ws = sheet();
    const seen: string[] = [];
    ws.getColumn(2).eachCell((cell: any, row: number) => seen.push(`${cell.address}:${String(row)}`));
    expect(seen).toEqual(['B1:1', 'B3:3', 'B5:5']);
    expect(ws.findCell(1, 2)).toBeDefined();
    const all: number[] = [];
    ws.getColumn(2).eachCell({ includeEmpty: true }, (_cell: any, row: number) => all.push(row));
    expect(all).toEqual([1, 2, 3, 4, 5, 6]);
  });

  it('values keeps the cells with values; style setters reach the rows with values', () => {
    const ws = sheet();
    expect(ws.getColumn(1).values).toEqual([undefined, 1, undefined, undefined, undefined, 'x']);
    expect(ws.getColumn(2).values).toEqual([undefined, undefined, undefined, 2]);
    ws.getColumn(3).fill = RED;
    const filled = [1, 2, 3, 4, 5, 6, 7].filter((r) => ws.findCell(r, 3)?.fill !== undefined);
    expect(filled).toEqual([1, 3, 5]);
    expect(ws.getColumn(3).fill).toEqual(RED);
    // A cell created later takes the column style (lib/doc/cell.js L27).
    expect(ws.getCell('C7').fill).toEqual(RED);
  });
});

describe('formulaType', () => {
  // lib/doc/cell.js L264-266: `this._value.formulaType`, which only the formula value defines
  // (L807-815); exceljs 4.4.0 gives undefined for a number, an empty cell and a merged cell.
  it('is undefined for a cell without a formula', () => {
    const ws = new ExcelJS.Workbook().addWorksheet('S');
    ws.getCell('A1').value = 1;
    ws.getCell('B1').value = { formula: 'A1*2', result: 2 };
    ws.getCell('C1').value = { sharedFormula: 'B1', result: 2 };
    ws.mergeCells('B2:C2');
    ws.getCell('B2').value = { formula: 'A1', result: 1 };
    expect(ws.getCell('A1').formulaType).toBeUndefined();
    expect(ws.getCell('Z9').formulaType).toBeUndefined();
    expect(ws.getCell('C2').formulaType).toBeUndefined();
    expect(ws.getCell('B1').formulaType).toBe(ExcelJS.FormulaType.Master);
    expect(ws.getCell('C1').formulaType).toBe(ExcelJS.FormulaType.Shared);
    expect(ws.getCell('B2').formulaType).toBe(ExcelJS.FormulaType.Master);
  });
});

describe('csv.read: an empty line', () => {
  // fast-csv 4.3.6 gives [] for a line of white space only (@fast-csv/parse/build/src/parser/
  // RowParser.js L15-48: the scanner skips white space, ParserOptions NEXT_TOKEN_REGEXP), and
  // exceljs adds `data.map(map)` (lib/csv/csv.js L86-89), so `map` is not called for it.
  // Checked with exceljs 4.4.0: for "a,b\n\nc\n \nd" map is called for a, b, c and d only, and
  // the empty lines are rows without values. A quoted empty field is a field ([""]).
  it('gives a row without fields, and map is not called for it', async () => {
    const calls: unknown[] = [];
    const ws = await new ExcelJS.Workbook().csv.read(['a,b\n\nc\n \t\nd\n""\n'], {
      map: (v: unknown) => {
        calls.push(v);
        return v;
      },
    });
    expect(calls).toEqual(['a', 'b', 'c', 'd', '']);
    expect(ws.getRow(2).values).toEqual([]);
    expect(ws.getRow(4).values).toEqual([]);
    expect(ws.getCell('A3').value).toBe('c');
    expect(ws.getCell('A6').value).toBe('');
    expect(ws.rowCount).toBe(6);
  });
});

describe('F3: row styles', () => {
  // exceljs writes a row's style as row@s with customFormat="1" (lib/xlsx/xform/sheet/row-xform.js
  // L19-23, L44-47) for the rows it writes: those with a height or cells (lib/doc/row.js L333-364,
  // `get model` is null otherwise). It reads row@s back as row.style (row-xform.js L80-82, L130),
  // and a cell created later in the row takes it (lib/doc/cell.js L27). Checked with exceljs
  // 4.4.0: of rows 2 (style only), 3 (style and a cell) and 4 (style and a height), rows 3 and 4
  // are written with s and customFormat="1"; read back, a new cell B3 is italic.
  async function written(wb: any): Promise<Uint8Array> {
    return (await wb.xlsx.writeBuffer()) as Uint8Array;
  }
  function styledRows(): any {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.getRow(2).font = { bold: true };
    ws.getRow(3).font = { italic: true };
    ws.getCell('A3').value = 1;
    ws.getRow(4).fill = RED;
    ws.getRow(4).height = 20;
    // xlsx.load looks at row settings up to the dimension or the last cell (docs/exceljs-compat.md,
    // "Writing workbooks that were read from files"), so a cell after row 4.
    ws.getCell('A5').value = 'end';
    return wb;
  }

  it('row.style is written as the row style of the rows exceljs writes', async () => {
    const bytes = await written(styledRows());
    const xml = strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']!);
    expect(xml).toMatch(/<row r="3"[^>]* s="\d+"[^>]* customFormat="1"/);
    expect(xml).toMatch(/<row r="4"[^>]* s="\d+"[^>]* customFormat="1"/);
    expect(xml).not.toContain('<row r="2"');
    const sheet = readWorkbook(bytes).sheets[0]!;
    expect(sheet.rowInfo(3)?.style?.font?.italic).toBe(true);
    expect(sheet.rowInfo(4)?.style?.fill).toMatchObject({ pattern: 'solid' });
    // Without styles, no row style is written.
    const plain = strFromU8(unzipSync((await styledRows().xlsx.writeBuffer({ useStyles: false })) as Uint8Array)['xl/worksheets/sheet1.xml']!);
    expect(plain).not.toContain('customFormat');
  });

  it('xlsx.load reads the row style, and new cells in the row take it', async () => {
    const back = (await load(await written(styledRows()))).getWorksheet('S');
    expect(back.getRow(3).font).toMatchObject({ italic: true });
    expect(back.getCell('B3').font).toMatchObject({ italic: true });
    expect(back.getRow(4).fill).toEqual(RED);
    expect(back.getCell('C4').fill).toEqual(RED);
    // A cell with a style of its own in the file keeps it.
    const file = buildXlsx({
      sheets: [{ name: 'S', data: '<row r="1" s="1" customFormat="1"><c r="A1" s="2"><v>1</v></c></row><row r="2" s="1"><c r="A2"><v>2</v></c></row>' }],
      styles: STYLES,
    });
    const ws = (await load(file)).getWorksheet('S');
    expect(ws.getRow(1).font?.bold).toBe(true);
    expect(ws.getCell('A1').font?.bold).toBeUndefined();
    expect(ws.getCell('B1').font?.bold).toBe(true);
    // Without customFormat="1" the row's s does not apply (ECMA-376 Part 1 §18.3.1.73), and the
    // core does not read it; exceljs reads it all the same (row-xform.js L80-82).
    expect(ws.getRow(2).style).toEqual({});
  });

  it('the streaming writer writes the row style, and the streaming reader reads it', async () => {
    await withTempDir(async (dir) => {
      const filename = join(dir, 'rows.xlsx');
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename, useStyles: true });
      const ws = wb.addWorksheet('S');
      const row = ws.getRow(1);
      row.font = { italic: true };
      row.getCell(1).value = 'a';
      row.commit();
      await wb.commit();
      const bytes = new Uint8Array(readFileSync(filename));
      expect(readWorkbook(bytes).sheets[0]!.rowInfo(1)?.style?.font?.italic).toBe(true);
      let style: unknown;
      for await (const sheet of new ExcelJS.stream.xlsx.WorkbookReader(bytes, { styles: 'cache' })) for await (const r of sheet as any) style = r.style;
      expect(style).toMatchObject({ font: { italic: true } });
    });
  });
});

describe('F4: the workbook default font', () => {
  // exceljs gives a style read from a file (s >= 1 for a cell, and row and column styles) the
  // font its fontId names, also font 0 (lib/xlsx/xform/style/styles-xform.js L296-345
  // `getStyleModel`, `addStyle` also for id 0); a cell with s="0" or no s has no style
  // (lib/xlsx/xform/sheet/cell-xform.js L428-432). Checked with exceljs 4.4.0 on this file: A1
  // (no s) and D1 (s="0") have no font, B1 (font 0 with format 0.000) is Arial 10, C1 bold
  // Arial 10, the row style and column C's style Arial 10.
  const FONT0_STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts>
<fonts count="2"><font><sz val="10"/><name val="Arial"/></font><font><b/><sz val="10"/><name val="Arial"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>
</styleSheet>`;
  const file = buildXlsx({
    sheets: [{ name: 'S', before: '<cols><col min="3" max="3" width="9" style="1"/></cols>', data: '<row r="1" s="1" customFormat="1"><c r="A1"><v>1</v></c><c r="B1" s="1"><v>2</v></c><c r="C1" s="2"><v>3</v></c><c r="D1" s="0"><v>4</v></c></row>' }],
    styles: FONT0_STYLES,
  });

  it('styled cells, rows and columns that use font 0 read it as their font', async () => {
    const ws = (await load(file)).getWorksheet('S');
    expect(ws.getCell('A1').font).toBeUndefined();
    expect(ws.getCell('D1').font).toBeUndefined();
    expect(ws.getCell('B1').font).toMatchObject({ name: 'Arial', size: 10 });
    expect(ws.getCell('B1').font.bold).toBeUndefined();
    expect(ws.getCell('C1').font).toMatchObject({ name: 'Arial', size: 10, bold: true });
    expect(ws.getRow(1).font).toMatchObject({ name: 'Arial', size: 10 });
    expect(ws.getColumn(3).font).toMatchObject({ name: 'Arial', size: 10 });
    expect(ws.getCell('E1').font).toMatchObject({ name: 'Arial', size: 10 });
  });

  it('a loaded workbook is written with its default font; one made in code with Calibri 11', async () => {
    // exceljs writes Calibri 11 as font 0 whatever the file had (styles-xform.js), so its cells
    // without a style become Calibri 11; xlvane keeps the file's font 0.
    const bytes = new Uint8Array(readFileSync(`${POI}47737.xlsx`));
    const wb = await load(bytes);
    const out = (await wb.xlsx.writeBuffer()) as Uint8Array;
    const read = readWorkbook(out);
    expect(read.defaultFont).toMatchObject({ name: 'Arial', size: 10 });
    // N1 uses font 0 with a style of its own: exceljs reads it as Arial 10.
    expect(wb.worksheets[0].getCell('N1').font).toMatchObject({ name: 'Arial', size: 10 });
    const n1 = [...read.sheets[0]!.cells(), ...read.sheets[0]!.styledCells()].find((c) => c.ref === 'N1');
    expect(n1?.style.font?.name ?? read.defaultFont?.name).toBe('Arial');
    const fresh = new ExcelJS.Workbook();
    fresh.addWorksheet('S').getCell('A1').value = 1;
    expect(readWorkbook((await fresh.xlsx.writeBuffer()) as Uint8Array).defaultFont).toMatchObject({ name: 'Calibri', size: 11 });
  });
});

describe('F4: a default font the writer does not accept', () => {
  // Review of 1.3.1: font 0 with a name of more than 31 characters, an empty name, or a size
  // outside 1 to 409.55 points made xlsx.writeBuffer throw XLV_E_INVALID_STYLE, because it was
  // passed to the core as the default font. 1.3.0 and exceljs 4.4.0 write these files (exceljs always writes
  // its own Calibri 11 as font 0). Such parts of font 0 are now left out, so saving never fails
  // because of font 0; a cell that uses font 0 gets the same accepted font.
  const fontFile = (font0: string): Uint8Array =>
    buildXlsx({
      sheets: [{ name: 'S', data: '<row r="1"><c r="A1"><v>1</v></c><c r="B1" s="1"><v>2</v></c></row>' }],
      styles: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1">${font0}</fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>
</styleSheet>`,
    });
  const cases: [string, string, Record<string, unknown>][] = [
    ['a 60-character name', `<font><b/><sz val="11"/><name val="${'N'.repeat(60)}"/></font>`, { size: 11, bold: true }],
    ['an empty name', '<font><sz val="11"/><name val=""/></font>', { size: 11 }],
    ['size 5000', '<font><sz val="5000"/><name val="Arial"/></font>', { name: 'Arial' }],
    ['size 0.5', '<font><sz val="0.5"/><name val="Arial"/></font>', { name: 'Arial' }],
  ];
  for (const [name, font0, kept] of cases) {
    it(`a file whose font 0 has ${name} is written`, async () => {
      const wb = await load(fontFile(font0));
      const out = (await wb.xlsx.writeBuffer()) as Uint8Array;
      const read = readWorkbook(out);
      expect(read.defaultFont).toMatchObject(kept);
      expect(wb.getWorksheet('S').getCell('B1').font).toMatchObject(kept);
      expect((await load(out)).getWorksheet('S').getCell('B1').value).toBe(2);
    });
  }
});

describe('F3, F4: style index 0 on a column or row', () => {
  // exceljs ignores style id 0 of a column or row (lib/xlsx/xform/sheet/col-xform.js L80
  // `if (model.styleId)`, row-xform.js L130 `model.styleId ? ... : {}`). LibreOffice writes
  // <col style="0"> on every column. Review of 1.3.1: such a column or row got xf 0 with font 0
  // as its style, so new cells took it and the row was written back with s="1". Checked with
  // exceljs 4.4.0 on this file: column A and row 2 have the style {}, and so has a new cell B5;
  // the written file has no column or row style.
  const file = buildXlsx({
    sheets: [
      {
        name: 'S',
        before: '<cols><col min="1" max="3" width="10" style="0"/></cols>',
        data: '<row r="1"><c r="A1"><v>1</v></c></row><row r="2" s="0" customFormat="1"><c r="A2"><v>2</v></c></row>',
      },
    ],
    styles: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="10"/><name val="Arial"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="bottom"/></xf></cellXfs>
</styleSheet>`,
  });

  it('gives the column and row no style', async () => {
    const wb = await load(file);
    const ws = wb.getWorksheet('S');
    expect(ws.getColumn(1).style).toEqual({});
    expect(ws.getRow(2).style).toEqual({});
    expect(ws.getCell('B5').style).toEqual({});
    const xml = strFromU8(unzipSync((await wb.xlsx.writeBuffer()) as Uint8Array)['xl/worksheets/sheet1.xml']!);
    expect(xml).not.toContain('customFormat');
    expect(xml).not.toMatch(/<col [^>]*style=/);
  });
});

describe('F5: six-digit rgb colours', () => {
  // POI corpus 52348.xlsx stores colours as rgb="000000", "808080", "C0C0C0", "FF0000". exceljs
  // 4.4.0 reads them as { argb: '000000' } and so on (lib/xlsx/xform/style/color-xform.js L39-40)
  // and writes the 6 digits back (L19-20). xlvane reads them as opaque RGB, FF and the 6 digits,
  // and writes those 8 digits.
  it('survive xlsx.load and writeBuffer as FF followed by the six digits', async () => {
    const wb = await load(new Uint8Array(readFileSync(`${POI}52348.xlsx`)));
    const ws = wb.worksheets[0];
    expect(ws.getCell('C2').font.color).toEqual({ argb: 'FF000000' });
    const out = (await wb.xlsx.writeBuffer()) as Uint8Array;
    const styles = strFromU8(unzipSync(out)['xl/styles.xml']!);
    for (const rgb of ['FF000000', 'FF808080', 'FFC0C0C0', 'FFFF0000']) expect(styles).toContain(`rgb="${rgb}"`);
    const back = (await load(out)).worksheets[0];
    expect(back.getCell('C2').font.color).toEqual({ argb: 'FF000000' });
  });
});
