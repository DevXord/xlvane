// Column default style (`col@style`, docs/decisions.md D-A04): written as the index of a cell
// format (CT_Col, sml.xsd L2238–2252, W2-03; [MS-OI29500] §2.1.598 d: 0 … 65429), read back as
// `ColumnInfo.style`. The core never styles cells from their column.
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { editCells } from '../../packages/xlvane/src/edit';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { writeWorkbookStream } from '../../packages/xlvane/src/stream-write';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError } from '../helpers/expect-error';
import { buildXlsx } from '../helpers/xlsx-builder';

function part(bytes: Uint8Array, name: string): string {
  const data = unzipSync(bytes)[name];
  if (!data) throw new Error(`missing ${name}`);
  return strFromU8(data);
}

const BOLD = { name: 'Calibri', size: 11, bold: true };

async function collect(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

describe('column default style', () => {
  it('is written as col@style and read back as ColumnInfo.style', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', 'x').column('B', { style: { font: { bold: true }, numFmt: '0.00' } }).column('C', { width: 12, style: { fill: { color: 'FFFF00' } } });
    const bytes = writeWorkbook(wb);
    const xml = part(bytes, 'xl/worksheets/sheet1.xml');
    expect(xml).toMatch(/<col min="2" max="2" style="\d+"\/>/);
    expect(xml).toMatch(/<col min="3" max="3" width="12" customWidth="1" style="\d+"\/>/);
    const s = readWorkbook(bytes).getSheet('S');
    expect(s.columns).toEqual([
      { min: 2, max: 2, width: null, hidden: false, style: { font: { name: 'Calibri', size: 11, bold: true }, numFmt: '0.00' } },
      { min: 3, max: 3, width: 12, hidden: false, style: { fill: { pattern: 'solid', fgColor: 'FFFFFF00' } } },
    ]);
  });

  it('does not style cells of the column (the core is explicit)', () => {
    const wb = createWorkbook();
    wb.addSheet('S').column('A', { style: { font: { italic: true } } }).set('A1', 1);
    const s = readWorkbook(writeWorkbook(wb)).getSheet('S');
    expect(s.cell('A1')?.style).toEqual({});
  });

  it('style {} or the default style writes no style attribute', () => {
    const wb = createWorkbook();
    wb.addSheet('S').column('A', { width: 10, style: {} });
    expect(part(writeWorkbook(wb), 'xl/worksheets/sheet1.xml')).toContain('<col min="1" max="1" width="10" customWidth="1"/>');
  });

  it('later calls replace the style and keep width and hidden', () => {
    const wb = createWorkbook();
    wb.addSheet('S').column('A', { width: 10, hidden: true, style: { font: { bold: true } } }).column('A', { style: { font: { italic: true } } });
    const c = readWorkbook(writeWorkbook(wb)).getSheet('S').columns[0];
    expect(c).toEqual({ min: 1, max: 1, width: 10, hidden: true, style: { font: { name: 'Calibri', size: 11, italic: true } } });
  });

  it('invalid styles throw XLV_E_INVALID_STYLE', () => {
    const sheet = createWorkbook().addSheet('S');
    expectXlvError(() => sheet.column('A', { style: { font: { size: 0 } } }), 'XLV_E_INVALID_STYLE', { property: 'font.size' });
    expectXlvError(() => sheet.column('A', { style: 'bold' as never }), 'XLV_E_INVALID_STYLE');
  });

  it('the streaming writer writes column styles', async () => {
    const bytes = await collect(
      writeWorkbookStream({ sheets: [{ name: 'S', columns: [undefined, { style: { font: { bold: true } } }], rows: [[1, 2]] }] }),
    );
    expect(readWorkbook(bytes).getSheet('S').columns).toEqual([{ min: 2, max: 2, width: null, hidden: false, style: { font: BOLD } }]);
  });

  it('reads a style index of a file; out-of-range or invalid indexes give the default style', () => {
    const wb = createWorkbook();
    wb.addSheet('S').column('A', { style: { font: { bold: true } } }).set('B1', 1);
    const bytes = writeWorkbook(wb);
    const s = readWorkbook(bytes).getSheet('S');
    expect(s.columns[0]?.style).toEqual({ font: BOLD });
    // A column without the attribute has no style property.
    const plain = createWorkbook();
    plain.addSheet('S').column('A', { width: 5 });
    expect(readWorkbook(writeWorkbook(plain)).getSheet('S').columns[0]).not.toHaveProperty('style');
    // Hand-edited indexes: out of range gives the default (empty) style; not a number is ignored
    // with a warning.
    const withStyle = (value: string): Uint8Array => {
      const files = unzipSync(bytes);
      files['xl/worksheets/sheet1.xml'] = strToU8(part(bytes, 'xl/worksheets/sheet1.xml').replace(/style="\d+"/, `style="${value}"`));
      return zipSync(files);
    };
    expect(readWorkbook(withStyle('999')).getSheet('S').columns[0]?.style).toEqual({});
    const bad = readWorkbook(withStyle('x'));
    expect(bad.getSheet('S').columns[0]).not.toHaveProperty('style');
    expect(bad.warnings.map((w) => w.code)).toContain('XLV_W_INVALID_VALUE');
  });

  it('editCells keeps col@style and starts new cells from it (Part 1 §18.3.1.13)', () => {
    const wb = createWorkbook();
    wb.addSheet('S').column('A', { style: { font: { bold: true } } }).set('A1', 'old').set('B1', 1);
    const edited = editCells(writeWorkbook(wb), [
      { sheet: 'S', cell: 'A2', value: 'new' },
      { sheet: 'S', cell: 'A3', value: 'red', style: { font: { color: 'FF0000' } } },
      { sheet: 'S', cell: 'A1', value: 'changed' },
      { sheet: 'S', cell: 'B2', value: 2 },
    ]);
    const s = readWorkbook(edited.bytes).getSheet('S');
    expect(s.columns[0]?.style).toEqual({ font: BOLD });
    expect(s.cell('A2')?.style).toEqual({ font: BOLD });
    expect(s.cell('A3')?.style).toEqual({ font: { ...BOLD, color: 'FFFF0000' } });
    // An existing cell keeps its own format; other columns are unchanged.
    expect(s.cell('A1')?.style).toEqual({});
    expect(s.cell('B2')?.style).toEqual({});
  });

  it('editCells does not use the column style in a row with its own format, or an index the file lacks', () => {
    const sheet =
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols><col min="1" max="2" style="1"/><col min="3" max="3" style="99"/></cols>' +
      '<sheetData><row r="1" s="0" customFormat="1"><c r="D1"><v>1</v></c></row></sheetData></worksheet>';
    const styles =
      '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
      '<borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>';
    const bytes = buildXlsx({ sheets: [{ name: 'S', data: sheet }], styles });
    const edited = editCells(bytes, [
      { sheet: 'S', cell: 'B1', value: 'row format' },
      { sheet: 'S', cell: 'B2', value: 'column format' },
      { sheet: 'S', cell: 'C2', value: 'missing format' },
    ]);
    const s = readWorkbook(edited.bytes).getSheet('S');
    expect(s.cell('B1')?.style).toEqual({});
    expect(s.cell('B2')?.style.font?.bold).toBe(true);
    expect(s.cell('C2')?.style).toEqual({});
  });
});
