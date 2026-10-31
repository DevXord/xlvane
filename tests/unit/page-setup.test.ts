// Page setup, print options, margins, headers and footers, page breaks, print area and print
// titles (ECMA-376 Part 1 §18.3.1.3, .14, .46, .62–.65, .70, .74, §18.2.5; [MS-OI29500]
// §2.1.589, §2.1.599, §2.1.618–§2.1.638, §2.1.645; docs/facts/sheet-rules.md SR-23 … SR-31;
// docs/decisions.md D-C04).
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { insertRows } from '../../packages/xlvane/src/rows';
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
const sheetXml = (bytes: Uint8Array): string => part(bytes, 'xl/worksheets/sheet1.xml');

describe('writing', () => {
  it('writes nothing for a sheet without print settings', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', 1);
    const xml = sheetXml(writeWorkbook(wb));
    expect(xml).not.toMatch(/pageMargins|pageSetup|printOptions|headerFooter|Breaks|sheetProtection/);
  });

  it('writes the print elements in schema order after hyperlinks and before legacyDrawing', () => {
    const wb = createWorkbook();
    const s = wb.addSheet('S').set('A1', 'x').link('A1', { url: 'https://example.com' }).note('B2', 'n');
    s.pageSetup({ orientation: 'landscape', paperSize: 9, margins: { left: 0.25 }, printOptions: { gridLines: true, horizontalCentered: true } });
    s.headerFooter({ oddHeader: '&CReport & more', oddFooter: '&RPage &P of &N' });
    s.pageBreaks({ rows: [25], columns: ['F'] });
    const xml = sheetXml(writeWorkbook(wb));
    const order = ['</hyperlinks>', '<printOptions', '<pageMargins', '<pageSetup', '<headerFooter', '<rowBreaks', '<colBreaks', '<legacyDrawing'];
    const at = order.map((tag) => xml.indexOf(tag));
    expect(at.every((i) => i > 0)).toBe(true);
    expect([...at].sort((a, b) => a - b)).toEqual(at);
    expect(xml).toContain('<printOptions horizontalCentered="1" gridLines="1"/>');
    expect(xml).toContain('<pageMargins left="0.25" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>');
    expect(xml).toContain('<pageSetup paperSize="9" orientation="landscape"/>');
    expect(xml).toContain('<headerFooter><oddHeader>&amp;CReport &amp; more</oddHeader><oddFooter>&amp;RPage &amp;P of &amp;N</oddFooter></headerFooter>');
    // Zero-based ids, max across the other axis, as in the Part 1 examples.
    expect(xml).toContain('<rowBreaks count="1" manualBreakCount="1"><brk id="24" max="16383" man="1"/></rowBreaks>');
    expect(xml).toContain('<colBreaks count="1" manualBreakCount="1"><brk id="5" max="1048575" man="1"/></colBreaks>');
  });

  it('writes fitToPage in sheetPr after the tab colour', () => {
    const wb = createWorkbook();
    wb.addSheet('S').tabColor('FF0000').pageSetup({ fitToPage: true, fitToWidth: 1, fitToHeight: 0 });
    const xml = sheetXml(writeWorkbook(wb));
    expect(xml).toContain('<sheetPr><tabColor rgb="FFFF0000"/><pageSetUpPr fitToPage="1"/></sheetPr>');
    expect(xml).toContain('<pageSetup fitToHeight="0"/>');
  });

  it('writes the remaining page setup attributes only when they differ from the defaults', () => {
    const wb = createWorkbook();
    wb.addSheet('S').pageSetup({ scale: 80, firstPageNumber: 3, pageOrder: 'overThenDown', blackAndWhite: true, draft: true, cellComments: 'atEnd', errors: 'NA', horizontalDpi: 300, verticalDpi: 300, copies: 2 });
    expect(sheetXml(writeWorkbook(wb))).toContain(
      '<pageSetup scale="80" firstPageNumber="3" pageOrder="overThenDown" blackAndWhite="1" draft="1" cellComments="atEnd" useFirstPageNumber="1" errors="NA" horizontalDpi="300" verticalDpi="300" copies="2"/>',
    );
  });

  it('writes the print area and titles as local built-in names, quoting the sheet name', () => {
    const wb = createWorkbook();
    wb.addSheet("Bob's Sheet").printArea(['A1:G20', '$I$1:$J$5']).printTitles({ rows: '1:3', columns: 'A:B' });
    wb.addSheet('Other');
    const bytes = writeWorkbook(wb);
    const workbook = part(bytes, 'xl/workbook.xml');
    expect(workbook).toContain(`<definedName name="_xlnm.Print_Area" localSheetId="0">'Bob''s Sheet'!$A$1:$G$20,'Bob''s Sheet'!$I$1:$J$5</definedName>`);
    expect(workbook).toContain(`<definedName name="_xlnm.Print_Titles" localSheetId="0">'Bob''s Sheet'!$A:$B,'Bob''s Sheet'!$1:$3</definedName>`);
    // Listed in app.xml without the prefix, as Excel does (SR-31).
    expect(part(bytes, 'docProps/app.xml')).toContain(`<vt:lpstr>'Bob''s Sheet'!Print_Area</vt:lpstr>`);
    const read = readWorkbook(bytes).getSheet("Bob's Sheet");
    expect(read.printArea).toEqual(['A1:G20', 'I1:J5']);
    expect(read.printTitles).toEqual({ rows: '1:3', columns: 'A:B' });
    expect(readWorkbook(bytes).getSheet('Other').printArea).toBeNull();
  });

  it('removes settings with null and {}', () => {
    const wb = createWorkbook();
    wb.addSheet('S').printArea('A1:B2').printArea(null).printTitles({ rows: '1' }).printTitles(null).headerFooter({ oddHeader: 'x' }).headerFooter(null).pageBreaks({ rows: [3] }).pageBreaks({});
    const bytes = writeWorkbook(wb);
    expect(sheetXml(bytes)).not.toMatch(/headerFooter|Breaks/);
    expect(part(bytes, 'xl/workbook.xml')).not.toContain('definedName');
  });

  it('streams the same print elements', async () => {
    const stream = writeWorkbookStream({
      sheets: [
        {
          name: 'S',
          pageSetup: { orientation: 'portrait', margins: { top: 1 } },
          headerFooter: { oddFooter: '&P' },
          pageBreaks: { rows: [10] },
          printArea: 'A1:C9',
          printTitles: { rows: '1' },
          protect: { sort: false },
          rows: [['a'], ['b']],
        },
      ],
    });
    const bytes = new Uint8Array(await new Response(stream).arrayBuffer());
    const sheet = readWorkbook(bytes).getSheet('S');
    expect(sheet.pageSetup.orientation).toBe('portrait');
    expect(sheet.pageSetup.margins?.top).toBe(1);
    expect(sheet.headerFooter?.oddFooter).toBe('&P');
    expect(sheet.pageBreaks).toEqual({ rows: [10], columns: [] });
    expect(sheet.printArea).toEqual(['A1:C9']);
    expect(sheet.printTitles).toEqual({ rows: '1:1', columns: null });
    expect(sheet.protection?.sheet).toBe(true);
    expect(sheet.protection?.sort).toBe(false);
  });
});

describe('limits and validation ([MS-OI29500])', () => {
  const s = () => createWorkbook().addSheet('S');
  it('margins from 0 to below 49 inches (§2.1.637)', () => {
    expect(s().pageSetup({ margins: { left: 0, right: 48.9 } })).toBeDefined();
    expectXlvError(() => s().pageSetup({ margins: { left: 49 } }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s().pageSetup({ margins: { top: -0.1 } }), 'XLV_E_INVALID_OPTION');
  });
  it('scale 10 to 400 or 0, copies 1 to 32767, fitTo up to 32767 (§2.1.638)', () => {
    expect(s().pageSetup({ scale: 0 })).toBeDefined();
    expectXlvError(() => s().pageSetup({ scale: 5 }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s().pageSetup({ scale: 401 }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s().pageSetup({ copies: 0 }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s().pageSetup({ copies: 32768 }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s().pageSetup({ fitToWidth: 32768 }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s().pageSetup({ orientation: 'sideways' as never }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s().pageSetup({ papersize: 9 } as never), 'XLV_E_INVALID_OPTION');
  });
  it('header and footer texts of at most 255 characters (§2.1.618 b … §2.1.634 a)', () => {
    expect(s().headerFooter({ oddHeader: 'x'.repeat(255) })).toBeDefined();
    const e = expectXlvError(() => s().headerFooter({ oddHeader: 'x'.repeat(256) }), 'XLV_E_INVALID_OPTION');
    expect(e.hint).toContain('Shorten');
  });
  it('page breaks from row 2 and column B, at most 1022 row and 1023 column breaks (§2.1.589, §2.1.599, §2.1.645)', () => {
    expectXlvError(() => s().pageBreaks({ rows: [1] }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s().pageBreaks({ columns: ['A'] }), 'XLV_E_INVALID_OPTION');
    expect(s().pageBreaks({ rows: Array.from({ length: 1022 }, (_, i) => i + 2) })).toBeDefined();
    expectXlvError(() => s().pageBreaks({ rows: Array.from({ length: 1023 }, (_, i) => i + 2) }), 'XLV_E_INVALID_OPTION');
    expect(s().pageBreaks({ columns: Array.from({ length: 1023 }, (_, i) => i + 2) })).toBeDefined();
  });
  it('print areas and titles must be ranges', () => {
    expectXlvError(() => s().printArea('A1:'), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s().printArea([]), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s().printTitles({ rows: 'A' }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s().printTitles({ columns: '1:2' }), 'XLV_E_INVALID_OPTION');
  });
});

describe('reading', () => {
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const sheet = (inner: string, before = '', names = '') =>
    readWorkbook(
      buildXlsx({
        sheets: [{ name: 'Sheet 1', data: `<worksheet xmlns="${NS}">${before}<sheetData/>${inner}</worksheet>` }],
        workbookAfterSheets: names,
      }),
    );

  it('reads the defaults when the sheet has no print elements', () => {
    const s = sheet('').sheets[0]!;
    expect(s.pageSetup).toMatchObject({ margins: null, orientation: 'default', paperSize: 1, scale: 100, fitToPage: false, copies: 1, firstPageNumber: null });
    expect(s.headerFooter).toBeNull();
    expect(s.pageBreaks).toEqual({ rows: [], columns: [] });
    expect(s.printArea).toBeNull();
  });

  it('reads Excel-style elements', () => {
    const s = sheet(
      '<printOptions headings="1" gridLines="1" gridLinesSet="0"/><pageMargins left="0.25" right="0.25" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
        '<pageSetup paperSize="9" orientation="landscape" firstPageNumber="5" useFirstPageNumber="1" horizontalDpi="200" verticalDpi="200"/>' +
        '<headerFooter differentFirst="1" alignWithMargins="0"><oddFooter>&amp;L&amp;F_x0041_</oddFooter><firstHeader>First</firstHeader></headerFooter>' +
        '<rowBreaks count="1" manualBreakCount="1"><brk id="51" max="9" man="1"/></rowBreaks><colBreaks count="1" manualBreakCount="1"><brk id="11" max="1048575" man="1"/></colBreaks>',
      '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>',
    ).sheets[0]!;
    // gridLines prints only with gridLinesSet (Part 1 §18.3.1.70).
    expect(s.pageSetup.printOptions).toEqual({ gridLines: false, headings: true, horizontalCentered: false, verticalCentered: false });
    expect(s.pageSetup).toMatchObject({ paperSize: 9, orientation: 'landscape', firstPageNumber: 5, horizontalDpi: 200, fitToPage: true, margins: { left: 0.25, footer: 0.3 } });
    expect(s.headerFooter).toMatchObject({ oddFooter: '&L&FA', firstHeader: 'First', differentFirst: true, alignWithMargins: false, scaleWithDoc: true });
    expect(s.pageBreaks).toEqual({ rows: [52], columns: [12] });
  });

  it('reads print areas and titles of several forms (POI corpus forms)', () => {
    const names =
      `<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">'Sheet 1'!$A$1:$J$78,'Sheet 1'!$L$2</definedName>` +
      `<definedName name="_xlnm.Print_Titles" localSheetId="0">'Sheet 1'!$A:$B,'Sheet 1'!$2:$3</definedName></definedNames>`;
    const s = sheet('', '', names).sheets[0]!;
    expect(s.printArea).toEqual(['A1:J78', 'L2']);
    expect(s.printTitles).toEqual({ rows: '2:3', columns: 'A:B' });
    const other = sheet('', '', `<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">OFFSET(A1,0,0)</definedName></definedNames>`);
    expect(other.sheets[0]!.printArea).toBeNull();
    expect(other.definedNames[0]?.refersTo).toBe('OFFSET(A1,0,0)');
  });

  it('warns about invalid values and keeps the defaults', () => {
    const wb = sheet('<pageMargins left="x" right="0" top="0" bottom="0" header="0" footer="0"/><pageSetup scale="big" orientation="up"/>');
    expect(wb.sheets[0]!.pageSetup).toMatchObject({ margins: null, scale: 100, orientation: 'default' });
    expect(wb.warnings.map((w) => w.code)).toContain('XLV_W_INVALID_VALUE');
  });
});

describe('row operations keep the print names (phase B)', () => {
  it('insertRows moves the print area and titles of the sheet', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', 1).printArea('A1:C10').printTitles({ rows: '2:3' });
    const out = insertRows(writeWorkbook(wb), { sheet: 'S', at: 2, count: 2 });
    const read = readWorkbook(out.bytes).getSheet('S');
    expect(read.printArea).toEqual(['A1:C12']);
    expect(read.printTitles).toEqual({ rows: '4:5', columns: null });
  });
});
