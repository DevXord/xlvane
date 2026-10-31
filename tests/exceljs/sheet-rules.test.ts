// xlvane-specific tests of phase C1 in @xlvane/compat: data validations, conditional
// formatting, sheet protection and printing (D-C06; docs/facts/exceljs.md X4-16 … X4-19). Where
// exceljs writes data the specification or Excel rejects, xlvane writes valid data; the
// differences are listed in docs/exceljs-compat.md.
import { join } from 'node:path';
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { hashPassword, verifyPassword } from '../../packages/xlvane/src/password';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { ExcelJS, expectNotSupported, roundTrip, withTempDir } from './helpers';

async function bytesOf(wb: any): Promise<Uint8Array> {
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

async function sheetXml(wb: any, n = 1): Promise<string> {
  return strFromU8(unzipSync(await bytesOf(wb))[`xl/worksheets/sheet${String(n)}.xml`]!);
}

describe('data validations', () => {
  it('remove() deletes the validation (exceljs leaves an empty entry and writes a validation without a type)', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form');
    ws.dataValidations.add('A1', { type: 'whole', operator: 'greaterThan', formulae: [0] });
    ws.dataValidations.add('G1', { type: 'whole', operator: 'greaterThan', formulae: [0] });
    ws.dataValidations.remove('G1');
    expect(Object.keys(ws.dataValidations.model)).toEqual(['A1']);
    const xml = await sheetXml(wb);
    expect(xml.match(/<dataValidation /g)).toHaveLength(1);
    expect(xml).toContain('sqref="A1"');
  });

  it('finds the validation of a cell inside a range rule, also after reading', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form');
    ws.dataValidations.add('B2:B20', { type: 'list', allowBlank: true, formulae: ['"Yes,No"'] });
    expect(ws.getCell('B7').dataValidation).toEqual({ type: 'list', allowBlank: true, formulae: ['"Yes,No"'] });
    expect(ws.getCell('C7').dataValidation).toBeUndefined();
    const ws2 = (await roundTrip(wb)).getWorksheet('Form');
    expect(ws2.getCell('B20').dataValidation).toEqual({ type: 'list', allowBlank: true, formulae: ['"Yes,No"'] });
  });

  it('writes equal validations of single cells as ranges that do not overlap', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form');
    const rule = { type: 'whole', operator: 'greaterThan', formulae: [0] };
    // exceljs writes A2:B2 and B1:B2 for these (X4-16).
    for (const a of ['A2', 'B1', 'B2']) ws.getCell(a).dataValidation = { ...rule };
    const xml = await sheetXml(wb);
    const sqref = /<dataValidation [^>]*sqref="([^"]+)"/.exec(xml)![1];
    expect(sqref).toBe('B1:B2 A2');
  });

  it('keeps whole-column rules as one range when reading', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Form');
    ws.dataValidations.add('C:C', { type: 'textLength', operator: 'lessThanOrEqual', formulae: [10] });
    const ws2 = (await roundTrip(wb)).getWorksheet('Form');
    expect(Object.keys(ws2.dataValidations.model)).toEqual(['C1:C1048576']);
    expect(ws2.getCell('C500000').dataValidation).toEqual({ type: 'textLength', operator: 'lessThanOrEqual', formulae: [10] });
  });

  it('follows rows that are inserted above it', () => {
    const ws = new ExcelJS.Workbook().addWorksheet('Form');
    ws.getCell('A1').value = 'x';
    ws.dataValidations.add('B2:B4', { type: 'list', formulae: ['"a,b"'] });
    ws.insertRow(1, ['new']);
    expect(Object.keys(ws.dataValidations.model)).toEqual(['B3:B5']);
  });
});

describe('conditional formatting', () => {
  it('escapes quotes in the text of a containsText rule (exceljs writes them as they are)', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('CF');
    ws.addConditionalFormatting({ ref: 'B2:B9', rules: [{ type: 'containsText', operator: 'containsText', text: 'say "hi"', style: { font: { bold: true } } }] });
    const xml = await sheetXml(wb);
    expect(xml).toContain('<formula>NOT(ISERROR(SEARCH("say ""hi""",B2)))</formula>');
    expect(xml).toMatch(/type="containsText"[^>]*operator="containsText"[^>]*text="say &quot;hi&quot;"/);
    const [cf] = (await roundTrip(wb)).getWorksheet('CF').conditionalFormattings;
    expect(cf.rules[0]).toMatchObject({ type: 'containsText', operator: 'containsText', text: 'say "hi"', style: { font: { bold: true } } });
  });

  it('generates the formulas of blank, error and timePeriod rules as exceljs does', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('CF');
    ws.addConditionalFormatting({
      ref: 'C3:C9',
      rules: [
        { type: 'containsText', operator: 'containsBlanks', style: { font: { italic: true } } },
        { type: 'containsText', operator: 'notContainsErrors', style: { font: { italic: true } } },
        { type: 'timePeriod', timePeriod: 'today', style: { font: { bold: true } } },
      ],
    });
    const xml = await sheetXml(wb);
    expect(xml).toContain('<formula>LEN(TRIM(C3))=0</formula>');
    expect(xml).toContain('<formula>NOT(ISERROR(C3))</formula>');
    expect(xml).toContain('<formula>FLOOR(C3,1)=TODAY()</formula>');
  });

  it('gives rules without a priority unique priorities (exceljs repeats the highest one)', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('CF');
    ws.addConditionalFormatting({ ref: 'A1:A9', rules: [{ type: 'expression', priority: 3, formulae: ['A1>1'], style: { font: { bold: true } } }] });
    ws.addConditionalFormatting({ ref: 'B1:B9', rules: [{ type: 'expression', formulae: ['B1>1'], style: { font: { bold: true } } }] });
    const xml = await sheetXml(wb);
    const priorities = [...xml.matchAll(/priority="(\d+)"/g)].map((m) => m[1]);
    expect(new Set(priorities).size).toBe(priorities.length);
  });

  it('reads colour scales and icon sets back as exceljs gives them', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('CF');
    ws.addConditionalFormatting({
      ref: 'A1:A10',
      rules: [
        { type: 'colorScale', cfvo: [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }], color: [{ argb: 'FFF8696B' }, { argb: 'FFFFEB84' }, { argb: 'FF63BE7B' }] },
        { type: 'iconSet', iconSet: '3Arrows', cfvo: [{ type: 'percent', value: 0 }, { type: 'percent', value: 33 }, { type: 'percent', value: 67 }] },
      ],
    });
    const [cf] = (await roundTrip(wb)).getWorksheet('CF').conditionalFormattings;
    expect(cf.ref).toBe('A1:A10');
    expect(cf.rules[0]).toMatchObject({ type: 'colorScale', cfvo: [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }], color: [{ argb: 'FFF8696B' }, { argb: 'FFFFEB84' }, { argb: 'FF63BE7B' }] });
    expect(cf.rules[1]).toMatchObject({ type: 'iconSet', iconSet: '3Arrows', reverse: false, showValue: true });
  });

  const cfSheet = (rules: any[]) => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('CF');
    for (let i = 1; i <= 9; i++) ws.getCell(`A${String(i)}`).value = i - 5;
    ws.addConditionalFormatting({ ref: 'A1:A9', rules });
    return wb;
  };

  it('writes a data bar without gradient: true as an x14 data bar (X4-20, D-C07)', async () => {
    const wb = cfSheet([{ type: 'dataBar', cfvo: [{ type: 'min' }, { type: 'max' }], color: { argb: 'FF638EC6' } }]);
    const xml = await sheetXml(wb);
    expect(xml).toContain('<x14:dataBar');
    expect(xml).toMatch(/<x14:dataBar[^>]*minLength="0"[^>]*maxLength="100"/);
    // exceljs writes no axis colour; [MS-XLSX] §2.6.30 requires one for an axis, so there is none.
    expect(xml).toContain('axisPosition="none"');
    const bar = readWorkbook(await bytesOf(wb)).getSheet('CF').conditionalFormats[0]!.rules[0]!.dataBar!;
    expect(bar).toMatchObject({ extended: true, gradient: true, minLength: 0, maxLength: 100, axisPosition: 'none' });
    const [cf] = (await roundTrip(wb)).getWorksheet('CF').conditionalFormattings;
    expect(cf.rules[0]).toMatchObject({ type: 'dataBar', cfvo: [{ type: 'min' }, { type: 'max' }], color: { argb: 'FF638EC6' }, minLength: 0, maxLength: 100, gradient: true, border: false, axisPosition: 'none' });
  });

  it('maps the x14 data bar properties and reads them back in exceljs shape', async () => {
    const rule = {
      type: 'dataBar',
      gradient: false,
      minLength: 5,
      maxLength: 80,
      border: true,
      borderColor: { argb: 'FF0000FF' },
      negativeBarColorSameAsPositive: false,
      negativeFillColor: { argb: 'FFFF0000' },
      negativeBarBorderColorSameAsPositive: false,
      negativeBorderColor: { argb: 'FF990000' },
      axisPosition: 'middle',
      axisColor: { argb: 'FF000000' },
      direction: 'rightToLeft',
      cfvo: [{ type: 'autoMin' }, { type: 'autoMax' }],
      color: { argb: 'FF00FF00' },
    };
    const [cf] = (await roundTrip(cfSheet([rule]))).getWorksheet('CF').conditionalFormattings;
    expect(cf.rules[0]).toMatchObject({
      ...rule,
      priority: 1,
    });
  });

  it('refuses border: true without a borderColor, which [MS-XLSX] requires', async () => {
    let code = '';
    try {
      await bytesOf(cfSheet([{ type: 'dataBar', border: true, cfvo: [{ type: 'min' }, { type: 'max' }], color: { argb: 'FF638EC6' } }]));
    } catch (error) {
      code = (error as { code?: string }).code ?? '';
    }
    expect(code).toBe('XLV_E_INVALID_OPTION');
  });

  it('writes the x14 icon sets and custom icons, keeping the given icon ids (exceljs renumbers them)', async () => {
    const cfvo3 = [{ type: 'percent', value: 0 }, { type: 'percent', value: 33 }, { type: 'percent', value: 67 }];
    const cfvo4 = [{ type: 'percent', value: 0 }, { type: 'percent', value: 25 }, { type: 'percent', value: 50 }, { type: 'percent', value: 75 }];
    const icons = [{ iconSet: '3Stars', iconId: 2 }, { iconSet: '3Flags', iconId: 1 }, { iconSet: 'NoIcons', iconId: 0 }, { iconSet: '4Arrows', iconId: 3 }];
    const wb = cfSheet([
      { type: 'iconSet', iconSet: '3Stars', cfvo: cfvo3 },
      { type: 'iconSet', iconSet: '5Boxes', cfvo: [...cfvo4, { type: 'percent', value: 90 }] },
      { type: 'iconSet', iconSet: '4Arrows', custom: true, reverse: true, showValue: false, cfvo: cfvo4, icons },
    ]);
    const xml = await sheetXml(wb);
    expect(xml).toContain('<x14:cfIcon iconSet="3Stars" iconId="2"/>');
    const [cf] = (await roundTrip(wb)).getWorksheet('CF').conditionalFormattings;
    expect(cf.rules[0]).toMatchObject({ type: 'iconSet', iconSet: '3Stars', cfvo: cfvo3 });
    expect(cf.rules[1]).toMatchObject({ type: 'iconSet', iconSet: '5Boxes' });
    // [MS-XLSX] §2.6.28: reverse is ignored with custom icons, so it is not written (exceljs writes it).
    expect(cf.rules[2]).toMatchObject({ type: 'iconSet', iconSet: '4Arrows', custom: true, reverse: false, showValue: false, icons });
  });

  it('keeps custom icons through a read and a second write', async () => {
    const cfvo3 = [{ type: 'percent', value: 0 }, { type: 'percent', value: 33 }, { type: 'percent', value: 67 }];
    const icons = [{ iconSet: '3Flags', iconId: 0 }, { iconSet: '3Flags', iconId: 1 }, { iconSet: '3Stars', iconId: 2 }];
    const once = await roundTrip(cfSheet([{ type: 'iconSet', iconSet: '3Arrows', custom: true, cfvo: cfvo3, icons }]));
    const twice = await roundTrip(once);
    expect(twice.getWorksheet('CF').conditionalFormattings[0].rules[0]).toMatchObject({ custom: true, icons });
  });

  it('removeConditionalFormatting takes an index, a filter or nothing (lib/doc/worksheet.js L815)', () => {
    const ws = new ExcelJS.Workbook().addWorksheet('CF');
    for (const ref of ['A1', 'B1', 'C1']) ws.addConditionalFormatting({ ref, rules: [{ type: 'expression', formulae: ['TRUE'] }] });
    ws.removeConditionalFormatting(0);
    expect(ws.conditionalFormattings.map((c: any) => c.ref)).toEqual(['B1', 'C1']);
    ws.removeConditionalFormatting((c: any) => c.ref !== 'C1');
    expect(ws.conditionalFormattings.map((c: any) => c.ref)).toEqual(['B1']);
    ws.removeConditionalFormatting();
    expect(ws.conditionalFormattings).toEqual([]);
  });
});

describe('sheet protection', () => {
  it('maps exceljs options (what users may do) to the attributes (what is locked)', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Locked');
    await ws.protect('', { formatCells: true, sort: true, selectLockedCells: false, objects: false });
    const xml = await sheetXml(wb);
    const tag = /<sheetProtection [^>]*\/>/.exec(xml)![0];
    expect(tag).toContain('sheet="1"');
    expect(tag).toContain('formatCells="0"');
    expect(tag).toContain('sort="0"');
    expect(tag).toContain('selectLockedCells="1"');
    expect(tag).toContain('objects="1"');
    const ws2 = (await roundTrip(wb)).getWorksheet('Locked');
    expect(ws2.sheetProtection).toEqual({ sheet: true, formatCells: true, sort: true, selectLockedCells: false, objects: false });
  });

  it('stores a password as exceljs does (SHA-512, random salt, spinCount), readable back and checkable', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Locked');
    const options = { spinCount: 1000.4, sort: true };
    await ws.protect('secret', options);
    expect(options.spinCount).toBe(1000.4);
    expect(ws.sheetProtection).toMatchObject({ sheet: true, algorithmName: 'SHA-512', spinCount: 1000, sort: true });
    expect(atob(ws.sheetProtection.saltValue)).toHaveLength(16);
    // The hash is the one exceljs computes for that salt (Encryptor.convertPasswordToHash).
    expect(ws.sheetProtection.hashValue).toBe(hashPassword('secret', { salt: ws.sheetProtection.saltValue, spinCount: 1000 }).hashValue);
    const ws2 = (await roundTrip(wb)).getWorksheet('Locked');
    expect(ws2.sheetProtection).toEqual(ws.sheetProtection);
    const read = readWorkbook(await bytesOf(wb)).getSheet('Locked');
    expect(verifyPassword(read.protection!, 'secret')).toBe(true);
    expect(verifyPassword(read.protection!, 'Secret')).toBe(false);
    // Saving a loaded file keeps the verifier.
    const again = readWorkbook(await bytesOf((await roundTrip(wb)) as any)).getSheet('Locked');
    expect(verifyPassword(again.protection!, 'secret')).toBe(true);
  });

  it('defaults spinCount to 100000 and drops it without a password; unprotect() removes the protection', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Locked');
    await ws.protect('x');
    expect(ws.sheetProtection.spinCount).toBe(100000);
    await ws.protect('', { spinCount: 5 });
    expect(ws.sheetProtection).toEqual({ sheet: true });
    ws.unprotect();
    expect(await sheetXml(wb)).not.toContain('<sheetProtection');
  });
});

describe('page setup and printing', () => {
  it('writes the exceljs default margins and page setup', async () => {
    const xml = await sheetXml((() => {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('Print').getCell('A1').value = 1;
      return wb;
    })());
    expect(xml).toContain('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>');
    expect(xml).toMatch(/<pageSetup [^>]*orientation="portrait"/);
  });

  it('round-trips orientation, paper, fit to page, comments ("asDisplyed", exceljs spelling) and print area', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Bob's Sheet", { pageSetup: { orientation: 'landscape', paperSize: 9, fitToWidth: 1, fitToHeight: 0 } });
    ws.getCell('A1').value = 1;
    ws.pageSetup.cellComments = 'asDisplyed';
    ws.pageSetup.printArea = 'A1:G20&&I1:J5';
    ws.pageSetup.showGridLines = true;
    const bytes = await bytesOf(wb);
    const xml = strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']!);
    expect(xml).toMatch(/<pageSetup [^>]*cellComments="asDisplayed"/);
    expect(xml).toContain('<printOptions gridLines="1"/>');
    // An apostrophe in the sheet name is doubled (exceljs writes the name as it is, X4-19).
    expect(strFromU8(unzipSync(bytes)['xl/workbook.xml']!)).toContain("'Bob''s Sheet'!$A$1:$G$20,'Bob''s Sheet'!$I$1:$J$5");
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(bytes);
    const ps = wb2.getWorksheet("Bob's Sheet").pageSetup;
    expect(ps).toMatchObject({ orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, cellComments: 'asDisplayed', showGridLines: true, printArea: 'A1:G20&&I1:J5' });
    expect(wb2.definedNames.model).toEqual([]);
  });

  it('round-trips headers, footers and row breaks (exceljs drops row breaks when reading)', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Print', { headerFooter: { oddHeader: '&CReport', oddFooter: '&RPage &P of &N' } });
    for (let i = 1; i <= 40; i++) ws.addRow([i]);
    ws.getRow(20).addPageBreak();
    const xml = await sheetXml(wb);
    expect(xml).toContain('<brk id="20" max="16383" man="1"/>');
    const ws2 = (await roundTrip(wb)).getWorksheet('Print');
    expect(ws2.headerFooter).toEqual({ oddHeader: '&CReport', oddFooter: '&RPage &P of &N' });
    expect(ws2.rowBreaks).toEqual([{ id: 20, max: 16383, man: 1 }]);
  });
});

describe('streamed worksheets', () => {
  it('write validations, conditional formats, protection and page setup set before the first committed row', async () => {
    await withTempDir(async (dir) => {
      const filename = join(dir, 'wb.test.xlsx');
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename, useStyles: true });
      const ws = wb.addWorksheet('Data', { pageSetup: { orientation: 'landscape' } });
      ws.dataValidations.add('A2:A100', { type: 'whole', operator: 'greaterThan', formulae: [0] });
      ws.addConditionalFormatting({ ref: 'A2:A100', rules: [{ type: 'cellIs', operator: 'greaterThan', formulae: [100], style: { font: { bold: true } } }] });
      await ws.protect('', { sort: true });
      ws.addRow(['n']).commit();
      ws.addRow([5]).commit();
      ws.commit();
      await wb.commit();
      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.readFile(filename);
      const ws2 = wb2.getWorksheet('Data');
      expect(ws2.getCell('A50').dataValidation).toMatchObject({ type: 'whole', operator: 'greaterThan', formulae: [0] });
      expect(ws2.conditionalFormattings[0].rules[0]).toMatchObject({ type: 'cellIs', operator: 'greaterThan', formulae: ['100'] });
      expect(ws2.sheetProtection).toEqual({ sheet: true, sort: true });
      expect(ws2.pageSetup.orientation).toBe('landscape');
    });
  });

  it('refuse changes made after rows were committed', async () => {
    await withTempDir(async (dir) => {
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: join(dir, 'wb.test.xlsx') });
      const ws = wb.addWorksheet('Data');
      ws.addRow([1]).commit();
      ws.dataValidations.add('A1', { type: 'whole', operator: 'greaterThan', formulae: [0] });
      await expectNotSupported(() => ws.commit());
    });
  });
});
