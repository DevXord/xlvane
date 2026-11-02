// insertRows/deleteRows move the conditional formats and data validations of the [MS-XLSX]
// extension (x14:conditionalFormatting, x14:dataValidation with xm:sqref and xm:f; SR-40 … SR-44)
// and the cell-reference thresholds of classic rules (cfvo@val, SR-12), as LibreOffice 24.2.7.2
// does (observed 2026-11-02, R-18 … R-20 in docs/facts/ref-adjust.md; D-F04).
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import type { Workbook } from '../../packages/xlvane/src/model';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { deleteRows, insertRows } from '../../packages/xlvane/src/rows';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError } from '../helpers/expect-error';
import { buildXlsx } from '../helpers/xlsx-builder';

const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const XM = 'http://schemas.microsoft.com/office/excel/2006/main';

/** The workbook observed in LibreOffice: a colour scale, an extended data bar and a 3Stars icon set with cell thresholds. */
function source(): Uint8Array {
  const wb = createWorkbook();
  const main = wb.addSheet('Main');
  wb.addSheet('Lists').set('A1', 5);
  for (let r = 1; r <= 20; r++) main.set(`B${String(r)}`, r).set(`D${String(r)}`, r).set(`E${String(r)}`, r).set(`G${String(r)}`, r);
  main.set('C12', 8).set('C15', 4);
  main.conditionalFormat('B2:B20', [{ type: 'colorScale', colorScale: { cfvo: [{ type: 'min' }, { type: 'formula', value: '$C$12' }], colors: ['F8696B', '63BE7B'] } }]);
  main.conditionalFormat('D2:D20', [{ type: 'dataBar', dataBar: { cfvo: [{ type: 'min' }, { type: 'formula', value: '$C$15' }], color: '638EC6', gradient: false, borderColor: '0000FF', axisColor: '000000' } }]);
  main.conditionalFormat('E5:E10 G3:G6', [{ type: 'iconSet', iconSet: { iconSet: '3Stars', cfvo: [{ type: 'percent', value: 0 }, { type: 'formula', value: '$C$12' }, { type: 'percent', value: 67 }] } }]);
  return writeWorkbook(wb);
}

/** Ranges, rule types and thresholds of the conditional formats of a sheet. */
function summary(wb: Workbook, sheet = 'Main'): string[] {
  return wb.getSheet(sheet).conditionalFormats.map((f) => {
    const r = f.rules[0]!;
    const cfvo = (r.colorScale ?? r.dataBar ?? r.iconSet)?.cfvo ?? [];
    return `${f.sqref} ${r.type} ${cfvo.map((v) => `${v.type}:${String(v.value)}`).join(',')}`;
  });
}

const sheetXml = (bytes: Uint8Array, n = 1): string => strFromU8(unzipSync(bytes)[`xl/worksheets/sheet${String(n)}.xml`]!);

describe('extension conditional formats and cell thresholds', () => {
  it('move with inserted rows (as LibreOffice: insert 2 rows at 5)', () => {
    const out = readWorkbook(insertRows(source(), { sheet: 'Main', at: 5, count: 2 }).bytes);
    expect(summary(out)).toEqual(['B2:B22 colorScale min:null,formula:$C$14', 'D2:D22 dataBar min:null,formula:$C$17', 'E7:E12 G3:G8 iconSet percent:0,formula:$C$14,percent:67']);
  });

  it('shrink with deleted rows (as LibreOffice: delete 2 rows at 5)', () => {
    const out = readWorkbook(deleteRows(source(), { sheet: 'Main', at: 5, count: 2 }).bytes);
    expect(summary(out)).toEqual(['B2:B18 colorScale min:null,formula:$C$10', 'D2:D18 dataBar min:null,formula:$C$13', 'E5:E8 G3:G4 iconSet percent:0,formula:$C$10,percent:67']);
  });

  it('become #REF! when the threshold cell is deleted (as LibreOffice: delete row 12)', () => {
    const out = readWorkbook(deleteRows(source(), { sheet: 'Main', at: 12, count: 1 }).bytes);
    expect(summary(out)).toEqual(['B2:B19 colorScale min:null,formula:#REF!', 'D2:D19 dataBar min:null,formula:$C$14', 'E5:E10 G3:G6 iconSet percent:0,formula:#REF!,percent:67']);
  });

  it('are removed with their extension when all their rows are deleted', () => {
    const wb = createWorkbook();
    const s = wb.addSheet('S');
    for (let r = 1; r <= 9; r++) s.set(`A${String(r)}`, r);
    s.conditionalFormat('A3:A4', [{ type: 'dataBar', dataBar: { cfvo: [{ type: 'min' }, { type: 'max' }], color: '638EC6', gradient: false, axisColor: '000000' } }]);
    s.conditionalFormat('A3:A4', [{ type: 'iconSet', iconSet: { iconSet: '3Triangles', cfvo: [{ type: 'percent', value: 0 }, { type: 'percent', value: 33 }, { type: 'percent', value: 67 }] } }]);
    s.conditionalFormat('A8', [{ type: 'iconSet', iconSet: { iconSet: '5Boxes', cfvo: [0, 20, 40, 60, 80].map((value) => ({ type: 'percent' as const, value })) } }]);
    const gone = deleteRows(writeWorkbook(wb), { sheet: 'S', at: 3, count: 2 }).bytes;
    expect(summary(readWorkbook(gone), 'S')).toEqual(['A6 iconSet percent:0,percent:20,percent:40,percent:60,percent:80']);
    expect(sheetXml(gone)).not.toContain('x14:dataBar');
    expect(sheetXml(gone)).not.toContain('3Triangles');
    const all = deleteRows(gone, { sheet: 'S', at: 6, count: 1 }).bytes;
    expect(readWorkbook(all).getSheet('S').conditionalFormats).toEqual([]);
    expect(sheetXml(all)).not.toContain('extLst');
  });

  it('are adjusted on other sheets, also when the changed sheet is named "Main" (in the namespace URIs)', () => {
    const wb = createWorkbook();
    wb.addSheet('Main').set('A1', 1).set('A5', 5);
    const other = wb.addSheet('Report');
    other.set('B1', 1);
    other.conditionalFormat('B1:B3', [{ type: 'iconSet', iconSet: { iconSet: '3Stars', cfvo: [{ type: 'percent', value: 0 }, { type: 'formula', value: 'Main!$A$5' }, { type: 'percent', value: 67 }] } }]);
    other.conditionalFormat('C1', [{ type: 'colorScale', colorScale: { cfvo: [{ type: 'num', value: 1 }, { type: 'formula', value: 'Main!$A$5' }], colors: ['FFFFFF', '000000'] } }]);
    const out = readWorkbook(insertRows(writeWorkbook(wb), { sheet: 'Main', at: 2, count: 3 }).bytes);
    // The reader lists rules that exist only in the extension after the classic ones.
    expect(summary(out, 'Report')).toEqual(['C1 colorScale num:1,formula:Main!$A$8', 'B1:B3 iconSet percent:0,formula:Main!$A$8,percent:67']);
  });
});

describe('extension data validations', () => {
  const file = (sqref: string): Uint8Array =>
    buildXlsx({
      sheets: [
        {
          name: 'S',
          data: '<row r="1"><c r="A1"><v>1</v></c></row>',
          after: `<extLst><ext uri="{CCE6A557-97BC-4B89-ADB6-D9C93CAAB3DF}" xmlns:x14="${X14}" xmlns:xm="${XM}"><x14:dataValidations count="2"><x14:dataValidation type="list" allowBlank="1"><x14:formula1><xm:f>Lists!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>${sqref}</xm:sqref></x14:dataValidation><x14:dataValidation type="list"><x14:formula1><xm:f>Lists!$B$1:$B$3</xm:f></x14:formula1><xm:sqref>C1:C2</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>`,
        },
        { name: 'Lists', data: '<row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c></row>' },
      ],
    });

  it('move with rows of their sheet, also with a line break in xm:sqref', () => {
    const out = readWorkbook(insertRows(file('E1:E20\n G5'), { sheet: 'S', at: 3, count: 1 }).bytes);
    expect(out.getSheet('S').dataValidations.map((v) => [v.sqref, v.formula1])).toEqual([
      ['E1:E21 G6', 'Lists!$A$1:$A$3'],
      ['C1:C2', 'Lists!$B$1:$B$3'],
    ]);
  });

  it('follow the rows of the list sheet', () => {
    const out = readWorkbook(insertRows(file('E1:E20'), { sheet: 'Lists', at: 2, count: 2 }).bytes);
    expect(out.getSheet('S').dataValidations.map((v) => v.formula1)).toEqual(['Lists!$A$1:$A$5', 'Lists!$B$1:$B$5']);
  });

  it('are removed when all their rows are deleted, and the count is updated', () => {
    const bytes = deleteRows(file('E5'), { sheet: 'S', at: 5, count: 1 }).bytes;
    expect(readWorkbook(bytes).getSheet('S').dataValidations.map((v) => v.sqref)).toEqual(['C1:C2']);
    expect(sheetXml(bytes)).toContain('<x14:dataValidations count="1">');
  });
});

describe('attribute values with "$"', () => {
  it('are written as they are (a cell threshold and a hyperlink location such as "Main!$A$14")', () => {
    const wb = createWorkbook();
    const main = wb.addSheet('Main');
    main.set('A1', 'Go').link('A1', { location: 'Main!$A$12' }).set('A12', 1);
    const out = readWorkbook(insertRows(writeWorkbook(wb), { sheet: 'Main', at: 5, count: 2 }).bytes).getSheet('Main');
    expect(out.hyperlinks.map((h) => h.location)).toEqual(['Main!$A$14']);
  });
});

describe('unusual extension layouts (found by review)', () => {
  const sheet = (ext: string): Uint8Array => buildXlsx({ sheets: [{ name: 'S', data: '<row r="1"><c r="A1"><v>1</v></c></row>', after: `<extLst>${ext}</extLst>` }] });
  const cf = (sqref: string) => `<x14:conditionalFormatting><x14:cfRule type="expression" priority="1"><xm:f>A1&gt;0</xm:f></x14:cfRule><xm:sqref>${sqref}</xm:sqref></x14:conditionalFormatting>`;
  const dv = (sqref: string) => `<x14:dataValidation type="list"><x14:formula1><xm:f>$Z$1:$Z$3</xm:f></x14:formula1><xm:sqref>${sqref}</xm:sqref></x14:dataValidation>`;

  it('an empty xm:sqref is left as it is', () => {
    const bytes = insertRows(sheet(`<ext uri="{78C0D931-6437-407D-A8EE-F0AAD7539E65}" xmlns:x14="${X14}" xmlns:xm="${XM}"><x14:conditionalFormattings>${cf('')}${cf('B2')}</x14:conditionalFormattings></ext>`), { sheet: 'S', at: 1, count: 1 }).bytes;
    expect(sheetXml(bytes)).toContain('<xm:sqref></xm:sqref>');
    expect(sheetXml(bytes)).toContain('<xm:sqref>B3</xm:sqref>');
  });

  it('a container whose items are all removed is removed, also when another container of the ext stays', () => {
    const ext = `<ext uri="{CCE6A557-97BC-4B89-ADB6-D9C93CAAB3DF}" xmlns:x14="${X14}" xmlns:xm="${XM}"><x14:dataValidations count="1">${dv('E5')}</x14:dataValidations><x14:dataValidations count="1">${dv('E9')}</x14:dataValidations></ext>`;
    const xml = sheetXml(deleteRows(sheet(ext), { sheet: 'S', at: 5, count: 1 }).bytes);
    expect(xml.match(/<x14:dataValidations /g)).toHaveLength(1);
    expect(xml).toContain('<xm:sqref>E8</xm:sqref>');
  });

  it('a known extension in an unexpected layout is refused instead of left unmoved', () => {
    const XM2 = 'urn:example:not-xm';
    const other = sheet(`<ext uri="{78C0D931-6437-407D-A8EE-F0AAD7539E65}" xmlns:x14="${X14}" xmlns:xm="${XM2}"><x14:conditionalFormattings>${cf('B2')}</x14:conditionalFormattings></ext>`);
    expectXlvError(() => insertRows(other, { sheet: 'S', at: 1, count: 1 }), 'XLV_E_NOT_SUPPORTED');
    const wrapped = sheet(`<ext uri="{78C0D931-6437-407D-A8EE-F0AAD7539E65}" xmlns:x14="${X14}" xmlns:xm="${XM}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:AlternateContent><mc:Choice Requires="x14"><x14:conditionalFormattings>${cf('B2')}</x14:conditionalFormattings></mc:Choice></mc:AlternateContent></ext>`);
    expectXlvError(() => insertRows(wrapped, { sheet: 'S', at: 1, count: 1 }), 'XLV_E_NOT_SUPPORTED');
  });
});

describe('other extensions that refer to cells', () => {
  it('are still refused on the changed sheet', () => {
    const bytes = buildXlsx({
      sheets: [
        {
          name: 'S',
          data: '<row r="1"><c r="A1"><v>1</v></c></row>',
          after: `<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="${X14}" xmlns:xm="${XM}"><x14:sparklineGroups><x14:sparklineGroup><x14:sparklines><x14:sparkline><xm:f>S!A1:A5</xm:f><xm:sqref>B1</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst>`,
        },
      ],
    });
    expectXlvError(() => insertRows(bytes, { sheet: 'S', at: 2, count: 1 }), 'XLV_E_NOT_SUPPORTED');
  });
});
