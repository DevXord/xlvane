// Row default style (`row@s` with `customFormat="1"`; CT_Row, sml.xsd L2256–2273, W2-04; E-24:
// `s` applies only when `customFormat` is 1). Written from `RowOptions.style` and
// `StreamRowInput.rowStyle`, read back as `RowInfo.style` and `StreamRow.style`. The core never
// styles cells from their row (as for column styles, D-A04).
import { readFileSync } from 'node:fs';
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { openWorkbookStream, type StreamRow } from '../../packages/xlvane/src/stream-read';
import { writeWorkbookStream } from '../../packages/xlvane/src/stream-write';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError } from '../helpers/expect-error';
import { buildXlsx } from '../helpers/xlsx-builder';

function part(bytes: Uint8Array, name: string): string {
  const data = unzipSync(bytes)[name];
  if (!data) throw new Error(`missing ${name}`);
  return strFromU8(data);
}

async function collect(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

async function streamRows(bytes: Uint8Array, sheet: string): Promise<StreamRow[]> {
  const book = await openWorkbookStream(bytes);
  const out: StreamRow[] = [];
  for await (const row of book.rows(sheet)) out.push(row);
  return out;
}

const BOLD = { name: 'Calibri', size: 11, bold: true };
const corpus = (name: string): Uint8Array => new Uint8Array(readFileSync(new URL(`../corpus/poi/${name}`, import.meta.url)));

/** A styles part with cellXfs 0 (plain) and 1 (bold font). */
const STYLES =
  '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
  '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/></cellXfs>';

describe('row default style, writing', () => {
  it('is written as row@s with customFormat="1" and read back as RowInfo.style', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A2', 'x').row(2, { style: { font: { bold: true }, numFmt: '0.00' } });
    const bytes = writeWorkbook(wb);
    expect(part(bytes, 'xl/worksheets/sheet1.xml')).toMatch(/<row r="2" s="[1-9]\d*" customFormat="1">/);
    expect(readWorkbook(bytes).getSheet('S').rowInfo(2)).toEqual({ height: null, hidden: false, style: { font: BOLD, numFmt: '0.00' } });
  });

  it('writes a row that has only a style', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', 1).row(3, { style: { fill: { color: 'FFFF00' } } });
    const bytes = writeWorkbook(wb);
    expect(part(bytes, 'xl/worksheets/sheet1.xml')).toMatch(/<row r="3" s="[1-9]\d*" customFormat="1"\/>/);
    expect(readWorkbook(bytes).getSheet('S').rowInfo(3)).toEqual({ height: null, hidden: false, style: { fill: { pattern: 'solid', fgColor: 'FFFFFF00' } } });
  });

  it('keeps height, hidden and outline next to the style, in schema attribute order', () => {
    const wb = createWorkbook();
    wb.addSheet('S').row(1, { height: 30, hidden: true, outlineLevel: 1 }).row(1, { style: { font: { bold: true } } });
    const bytes = writeWorkbook(wb);
    expect(part(bytes, 'xl/worksheets/sheet1.xml')).toMatch(/<row r="1" s="1" customFormat="1" ht="30" customHeight="1" hidden="1" outlineLevel="1"\/>/);
    expect(readWorkbook(bytes).getSheet('S').rowInfo(1)).toEqual({ height: 30, hidden: true, outlineLevel: 1, style: { font: BOLD } });
  });

  it('does not style cells of the row (the core is explicit)', () => {
    const wb = createWorkbook();
    wb.addSheet('S').row(1, { style: { font: { italic: true } } }).set('A1', 1);
    expect(readWorkbook(writeWorkbook(wb)).getSheet('S').cell('A1')?.style).toEqual({});
  });

  it('style {} writes no style attribute and later calls replace the style', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', 1).row(1, { height: 20, style: { font: { bold: true } } }).row(1, { style: {} });
    const bytes = writeWorkbook(wb);
    expect(part(bytes, 'xl/worksheets/sheet1.xml')).toContain('<row r="1" ht="20" customHeight="1">');
    expect(readWorkbook(bytes).getSheet('S').rowInfo(1)).toEqual({ height: 20, hidden: false });
  });

  it('shares the cell format with an equal cell style', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', 1, { font: { bold: true } }).row(2, { style: { font: { bold: true } } });
    const xml = part(writeWorkbook(wb), 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('<c r="A1" s="1">');
    expect(xml).toContain('<row r="2" s="1" customFormat="1"/>');
  });

  it('invalid styles throw XLV_E_INVALID_STYLE located at the row', () => {
    const sheet = createWorkbook().addSheet('S');
    const error = expectXlvError(() => sheet.row(4, { style: { font: { size: 0 } } }), 'XLV_E_INVALID_STYLE', { property: 'font.size' });
    expect(error.location).toMatchObject({ sheet: 'S', cell: '4:4' });
    expectXlvError(() => sheet.row(4, { style: 'bold' as never }), 'XLV_E_INVALID_STYLE');
  });
});

describe('row default style, streaming writer', () => {
  it('writes rowStyle as the row format, apart from the cell style', async () => {
    const bytes = await collect(
      writeWorkbookStream({
        sheets: [
          {
            name: 'S',
            rows: [
              { values: ['a'], rowStyle: { font: { bold: true } } },
              { values: [], rowStyle: { font: { italic: true } } },
              { values: ['c'], style: { font: { bold: true } } },
            ],
          },
        ],
      }),
    );
    const xml = part(bytes, 'xl/worksheets/sheet1.xml');
    expect(xml).toMatch(/<row r="1" s="\d+" customFormat="1"><c r="A1" t="inlineStr">/);
    expect(xml).toMatch(/<row r="2" s="\d+" customFormat="1"\/>/);
    expect(xml).toMatch(/<row r="3"><c r="A3" s="\d+"/);
    const s = readWorkbook(bytes).getSheet('S');
    expect(s.rowInfo(1)?.style).toEqual({ font: BOLD });
    expect(s.rowInfo(2)?.style).toEqual({ font: { name: 'Calibri', size: 11, italic: true } });
    expect(s.rowInfo(3)).toBeUndefined();
    expect(s.cell('A1')?.style).toEqual({});
  });

  it('rejects an invalid rowStyle with XLV_E_INVALID_STYLE', async () => {
    const stream = writeWorkbookStream({ sheets: [{ name: 'S', rows: [{ values: [1], rowStyle: { font: { size: 0 } } }] }] });
    await expect(collect(stream)).rejects.toMatchObject({ code: 'XLV_E_INVALID_STYLE' });
  });
});

describe('row default style, reading', () => {
  const read = (rows: string) => readWorkbook(buildXlsx({ styles: STYLES, sheets: [{ name: 'S', data: rows }] }));

  it('reads s only with customFormat="1" (E-24)', () => {
    const wb = read('<row r="1" s="1" customFormat="1"/><row r="2" s="1"><c r="A2"><v>1</v></c></row><row r="3" s="1" customFormat="0" ht="12"/><row r="4" customFormat="true"/>');
    const s = wb.getSheet('S');
    expect(s.rowInfo(1)).toEqual({ height: null, hidden: false, style: { font: BOLD } });
    expect(s.rowInfo(2)).toBeUndefined();
    expect(s.rowInfo(3)).toEqual({ height: 12, hidden: false });
    // `s` defaults to 0 (sml.xsd L2263): the row format is cell format 0.
    expect(s.rowInfo(4)).toEqual({ height: null, hidden: false, style: {} });
    expect(wb.warnings).toEqual([]);
  });

  it('warns about a style index that is not a number and ignores it', () => {
    const wb = read('<row r="1" s="x1" customFormat="1" ht="15"/>');
    expect(wb.getSheet('S').rowInfo(1)).toEqual({ height: 15, hidden: false });
    expect(wb.warnings).toMatchObject([{ code: 'XLV_W_INVALID_VALUE', count: 1 }]);
  });

  it('the streaming reader gives the row style', async () => {
    const bytes = buildXlsx({ styles: STYLES, sheets: [{ name: 'S', data: '<row r="1" s="1" customFormat="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row>' }] });
    const rows = await streamRows(bytes, 'S');
    expect(rows[0]?.style).toEqual({ font: BOLD });
    expect(rows[1]).not.toHaveProperty('style');
  });

  it('reads the row styles of files written by Excel (Apache POI corpus)', () => {
    // 47813.xlsx (Microsoft Macintosh Excel 12): <row r="1" spans="1:3" s="1" customFormat="1">,
    // cellXfs 1 uses font 1 (bold Verdana 10).
    const a = readWorkbook(corpus('47813.xlsx')).getSheet('Numbers');
    expect(a.rowInfo(1)?.style).toEqual({ font: { name: 'Verdana', size: 10, bold: true } });
    expect(a.rowInfo(2)?.style).toBeUndefined();
    // 48962.xlsx: <row r="7" spans="1:7" s="7" customFormat="1" ht="20">, cellXfs 7 uses font 7
    // (strike Verdana 10).
    const b = readWorkbook(corpus('48962.xlsx')).getSheet('Sheet1');
    expect(b.rowInfo(7)).toEqual({ height: 20, hidden: false, style: { font: { name: 'Verdana', size: 10, strike: true } } });
    expect(b.rowInfo(6)).toEqual({ height: 20, hidden: false });
  });
});
