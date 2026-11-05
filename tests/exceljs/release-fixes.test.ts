// Regression tests for the @xlvane/exceljs findings of the pre-release audits (docs/decisions.md
// D-R01): file-read values that must be written back, error locations, exceljs value shapes and
// own-property lookups. Test files are built here (tests/helpers/xlsx-builder.ts).
import { strFromU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { buildXlsx, c, row } from '../helpers/xlsx-builder';
import { ExcelJS } from './helpers';

const enc = new TextEncoder();

async function load(bytes: Uint8Array): Promise<any> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes);
  return wb;
}

async function reload(wb: any): Promise<any> {
  return load(await wb.xlsx.writeBuffer());
}

async function part(wb: any, name: string): Promise<string> {
  return strFromU8(unzipSync(await wb.xlsx.writeBuffer())[name]!);
}

/** Replaces text in one part of a package. */
function patch(bytes: Uint8Array, name: string, from: string | RegExp, to: string): Uint8Array {
  const files = unzipSync(bytes);
  const text = strFromU8(files[name]!);
  const next = text.replace(from, to);
  expect(next).not.toBe(text);
  files[name] = enc.encode(next);
  return zipSync(files);
}

async function caught(run: () => unknown): Promise<any> {
  try {
    await run();
  } catch (error) {
    return error;
  }
  throw new Error('expected an error');
}

describe('values read from files are written back', () => {
  it('page setup values of 0 (paperSize, copies, dpi) are left out instead of refused', async () => {
    const bytes = buildXlsx({
      sheets: [{ name: 'S', data: row(1, c('A1', 1)), after: '<pageSetup paperSize="0" copies="0" horizontalDpi="0" verticalDpi="0" orientation="landscape"/>' }],
    });
    const wb = await load(bytes);
    const xml = await part(wb, 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('orientation="landscape"');
    expect(xml).not.toMatch(/paperSize|copies|Dpi/);
  });

  it('shared formula groups the core cannot write as shared are written as ordinary formulas', async () => {
    // Master E1 whose ref does not start at the master, and two overlapping groups in column A.
    const data =
      row(1, '<c r="A1"><f t="shared" ref="A1:A4" si="0">B1*2</f><v>2</v></c>', '<c r="E1"><f t="shared" ref="D1:E2" si="1">B1+1</f><v>2</v></c>') +
      row(2, '<c r="A2"><f t="shared" si="0"/><v>4</v></c>', '<c r="E2"><f t="shared" si="1"/><v>3</v></c>') +
      row(3, '<c r="A3"><f t="shared" ref="A3:A4" si="2">B3*3</f><v>9</v></c>') +
      row(4, '<c r="A4"><f t="shared" si="2"/><v>12</v></c>');
    const wb = await load(buildXlsx({ sheets: [{ name: 'S', data }] }));
    const out = readWorkbook(await wb.xlsx.writeBuffer()).sheets[0]!;
    expect(out.cell('A2')?.formula?.text).toBe('B2*2');
    expect(out.cell('A2')?.value).toBe(4);
    expect(out.cell('A4')?.formula?.text).toBe('B4*3');
    expect(out.cell('E2')?.formula?.text).toBe('B2+1');
    expect(out.cell('E2')?.value).toBe(3);
  });

  it('a data validation formula that refers to a cell stays text (exceljs gives NaN)', async () => {
    const after =
      '<dataValidations count="2"><dataValidation type="whole" operator="greaterThan" sqref="A1"><formula1>$C$1</formula1></dataValidation>' +
      '<dataValidation type="decimal" operator="between" sqref="A2"><formula1>0.5</formula1><formula2>Limits!B2</formula2></dataValidation></dataValidations>';
    const wb = await load(buildXlsx({ sheets: [{ name: 'S', data: row(1, c('A1', 1)), after }] }));
    const ws = wb.getWorksheet('S');
    expect(ws.getCell('A1').dataValidation.formulae).toEqual(['$C$1']);
    expect(ws.getCell('A2').dataValidation.formulae).toEqual([0.5, 'Limits!B2']);
    const xml = await part(wb, 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('<formula1>$C$1</formula1>');
    expect(xml).toContain('<formula2>Limits!B2</formula2>');
  });

  it('EMF and WMF pictures are kept', async () => {
    const emf = new Uint8Array(64);
    emf.set([0x01, 0, 0, 0]);
    emf.set([0x20, 0x45, 0x4d, 0x46], 40);
    const wmf = new Uint8Array([0xd7, 0xcd, 0xc6, 0x9a, 0, 0, 1, 2, 3, 4]);
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.addImage(wb.addImage({ buffer: emf, extension: 'emf' }), 'B2:C3');
    ws.addImage(wb.addImage({ buffer: wmf, extension: 'wmf' }), 'E2:F3');
    const files = unzipSync(await wb.xlsx.writeBuffer());
    expect(Object.keys(files).filter((n) => n.startsWith('xl/media/')).sort()).toEqual(['xl/media/image1.emf', 'xl/media/image2.wmf']);
    expect(strFromU8(files['[Content_Types].xml']!)).toContain('image/x-emf');
    expect(strFromU8(files['[Content_Types].xml']!)).toContain('image/x-wmf');
    const back = await load(zipSync(files));
    expect(back.getWorksheet('S').getImages()).toHaveLength(2);
  });

  it('a hyperlink with neither a target nor a location is not turned into "#"', async () => {
    const wb = await load(buildXlsx({ sheets: [{ name: 'S', data: row(1, c('A1', 7)), after: '<hyperlinks><hyperlink ref="A1"/></hyperlinks>' }] }));
    expect(wb.getWorksheet('S').getCell('A1').value).toBe(7);
    await expect(wb.xlsx.writeBuffer()).resolves.toBeInstanceOf(Uint8Array);
  });

  it('a row height of 409.5 points (written by Excel) is accepted', async () => {
    const wb = await load(buildXlsx({ sheets: [{ name: 'S', data: '<row r="1" ht="409.5" customHeight="1"><c r="A1"><v>1</v></c></row>' }] }));
    expect(wb.getWorksheet('S').getRow(1).height).toBe(409.5);
    expect(await part(wb, 'xl/worksheets/sheet1.xml')).toContain('ht="409.5"');
  });

  it('a table with a custom style name of its file is written without the name', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('S').addTable({ name: 'T', ref: 'A1', columns: [{ name: 'A' }], rows: [[1]] });
    const bytes = patch(await wb.xlsx.writeBuffer(), 'xl/tables/table1.xml', /name="TableStyleMedium2"/, 'name="MyCompanyStyle"');
    const back = await load(bytes);
    expect(back.getWorksheet('S').getTable('T').table.style.theme).toBe('MyCompanyStyle');
    const xml = await part(back, 'xl/tables/table1.xml');
    expect(xml).not.toContain('MyCompanyStyle');
    // A name given in code is still checked.
    const own = new ExcelJS.Workbook();
    const err = await caught(() => own.addWorksheet('Data').addTable({ name: 'U', ref: 'A1', style: { theme: 'MyCompanyStyle' }, columns: [{ name: 'A' }], rows: [[1]] }));
    expect(err.code).toBe('XLV_E_INVALID_TABLE');
  });

  it('a table error names the worksheet, not a scratch sheet', async () => {
    const wb = new ExcelJS.Workbook();
    const err = await caught(() => wb.addWorksheet('Sales').addTable({ name: 'T', ref: 'A1', style: { theme: 'Nope' }, columns: [{ name: 'A' }], rows: [[1]] }));
    expect(err.location.sheet).toBe('Sales');
  });

  it('table totals keep their cached results when written again', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('S').addTable({ name: 'T', ref: 'A1', totalsRow: true, columns: [{ name: 'A', totalsRowLabel: 'Sum' }, { name: 'B', totalsRowFunction: 'sum' }], rows: [['x', 1], ['y', 2]] });
    const bytes = patch(await wb.xlsx.writeBuffer(), 'xl/worksheets/sheet1.xml', '<f>SUBTOTAL(109,T[B])</f>', '<f>SUBTOTAL(109,T[B])</f><v>3</v>');
    const back = await load(bytes);
    expect(back.getWorksheet('S').getCell('B4').value).toEqual({ formula: 'SUBTOTAL(109,T[B])', result: 3 });
    expect(await part(back, 'xl/worksheets/sheet1.xml')).toContain('<f>SUBTOTAL(109,T[B])</f><v>3</v>');
  });

  it('activeTab pointing at a chart sheet is mapped to a loaded worksheet', async () => {
    const bytes = buildXlsx({ sheets: [{ name: 'Data', data: row(1, c('A1', 1)) }, { name: 'Chart', kind: 'chartsheet' }] });
    const withView = patch(bytes, 'xl/workbook.xml', '<sheets>', '<bookViews><workbookView activeTab="1"/></bookViews><sheets>');
    const wb = await load(withView);
    expect(wb.views[0].activeTab).toBe(0);
    await expect(wb.xlsx.writeBuffer()).resolves.toBeInstanceOf(Uint8Array);
  });
});

describe('exceljs behaviour', () => {
  it('a hyperlink target that is itself a place in the workbook is not joined with its location', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('S').getCell('A1').value = { text: 'go', hyperlink: 'https://example.com/a' };
    let bytes = await wb.xlsx.writeBuffer();
    bytes = patch(bytes, 'xl/worksheets/_rels/sheet1.xml.rels', 'Target="https://example.com/a"', 'Target="#Other!B2"');
    bytes = patch(bytes, 'xl/worksheets/sheet1.xml', /<hyperlink ref="A1" r:id="(\w+)"/, '<hyperlink ref="A1" r:id="$1" location="Other!B2"');
    const back = await load(bytes);
    expect(back.getWorksheet('S').getCell('A1').value).toEqual({ text: 'go', hyperlink: '#Other!B2' });
  });

  it('row height and visibility of rows without cells are read', async () => {
    const data = '<row r="1"><c r="A1"><v>1</v></c></row><row r="3" ht="30" customHeight="1"/><row r="5" hidden="1"/>';
    const wb = await load(buildXlsx({ sheets: [{ name: 'S', data, before: '<dimension ref="A1:A5"/>' }] }));
    const ws = wb.getWorksheet('S');
    expect(ws.getRow(3).height).toBe(30);
    expect(ws.getRow(5).hidden).toBe(true);
    const out = readWorkbook(await wb.xlsx.writeBuffer()).sheets[0]!;
    expect(out.rowInfo(3)?.height).toBe(30);
    expect(out.rowInfo(5)?.hidden).toBe(true);
  });

  it('a style error at writeBuffer names the worksheet and the cell', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('Report').getCell('C7').font = { color: { argb: 'XYZ' } };
    const err = await caught(() => wb.xlsx.writeBuffer());
    expect(err.code).toBe('XLV_E_INVALID_STYLE');
    expect(err.location).toMatchObject({ sheet: 'Report', cell: 'C7' });
  });

  it('a tab colour error at writeBuffer names the worksheet', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('Report', { properties: { tabColor: { argb: 'nothex' } } });
    const err = await caught(() => wb.xlsx.writeBuffer());
    expect(err.location.sheet).toBe('Report');
  });

  it('the 7-digit tab colour of the exceljs README is written with a final 0', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('My Sheet', { properties: { tabColor: { argb: 'FFC0000' } } });
    expect(await part(wb, 'xl/worksheets/sheet1.xml')).toContain('<tabColor rgb="FFC00000"/>');
  });

  it('worksheet.dimensions is a Range as in exceljs', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    const empty = ws.dimensions;
    // Values from exceljs 4.4.0 (lib/doc/range.js) for an empty sheet and for B2 + D5.
    expect(empty.model).toEqual({ top: 0, left: 0, bottom: 0, right: 0 });
    expect(String(empty)).toBe('A1:A1');
    expect(empty.count).toBe(1);
    ws.getCell('B2').value = 1;
    ws.getCell('D5').value = 2;
    const d = ws.dimensions;
    expect(d.model).toEqual({ top: 2, left: 2, bottom: 5, right: 4 });
    expect([d.top, d.left, d.bottom, d.right]).toEqual([2, 2, 5, 4]);
    expect([d.tl, d.br, d.range, d.$range, d.shortRange, d.count, `${d}`]).toEqual(['B2', 'D5', 'B2:D5', '$B$2:$D$5', 'B2:D5', 12, 'B2:D5']);
  });

  it('a solid conditional-format fill set with bgColor reads back with bgColor', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 1;
    const style = { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFF0000' } } };
    ws.addConditionalFormatting({ ref: 'A1:A5', rules: [{ type: 'expression', priority: 1, formulae: ['A1>0'], style }] });
    const back = (await reload(wb)).getWorksheet('S');
    expect(back.conditionalFormattings[0].rules[0].style).toEqual(style);
  });

  it('Row.destroy and Cell.destroy exist and change nothing', () => {
    const ws = new ExcelJS.Workbook().addWorksheet('S');
    ws.getCell('A1').value = 1;
    ws.getRow(1).destroy();
    ws.getCell('A1').destroy();
    expect(ws.getCell('A1').value).toBe(1);
  });

  it('dataValidation errorStyle "error" (used in exceljs samples) is written as the default, stop', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').dataValidation = { type: 'whole', operator: 'notEqual', showErrorMessage: true, formulae: [5], errorStyle: 'error', error: 'Not 5' };
    const xml = await part(wb, 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('<dataValidation type="whole" operator="notEqual"');
    expect(xml).not.toContain('errorStyle');
    expect(readWorkbook(await wb.xlsx.writeBuffer()).sheets[0]!.dataValidations[0]!.errorStyle).toBe('stop');
  });

  it('the default export has PaperSize, as exceljs types it', () => {
    expect(ExcelJS.PaperSize.A4).toBe(9);
  });
});

describe('names that are Object.prototype keys (security L1)', () => {
  it('tables named "constructor" and "__proto__" are plain entries', () => {
    const ws = new ExcelJS.Workbook().addWorksheet('S');
    expect(ws.getTable('constructor')).toBeUndefined();
    expect(ws.getTable('toString')).toBeUndefined();
    ws.addTable({ name: 'constructor', ref: 'A1', columns: [{ name: 'A' }], rows: [[1]] });
    expect(ws.getTable('constructor').name).toBe('constructor');
    expect(Object.keys(ws.tables)).toEqual(['constructor']);
    expect(({} as Record<string, unknown>)['polluted']).toBeUndefined();
  });

  it('dataValidations.find("constructor") and an "__proto__" key do not reach Object.prototype', () => {
    const ws = new ExcelJS.Workbook().addWorksheet('S');
    expect(ws.dataValidations.find('constructor')).toBeUndefined();
    expect(ws.dataValidations.find('__proto__')).toBeUndefined();
    ws.dataValidations.add('__proto__', { type: 'list', formulae: ['"a"'], polluted: true } as never);
    expect(Object.keys(ws.dataValidations.model)).toEqual(['__proto__']);
    expect(({} as Record<string, unknown>)['polluted']).toBeUndefined();
    ws.dataValidations.model = JSON.parse('{"__proto__": {"type": "any"}, "A1": {"type": "any"}}') as never;
    expect(Object.keys(ws.dataValidations.model).sort()).toEqual(['A1', '__proto__']);
    expect(ws.getCell('B9').dataValidation).toBeUndefined();
  });
});

describe('CSV and splice hardening (security M3, M4, M5, L5)', () => {
  async function csvText(values: unknown[][], formatterOptions?: Record<string, unknown>): Promise<string> {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    for (const v of values) ws.addRow(v);
    return new TextDecoder().decode(await wb.csv.writeBuffer(formatterOptions ? { formatterOptions } : {}));
  }

  it('M3: a field starting with NUL and then "=" is escaped', async () => {
    const text = await csvText([['\0=cmd|calc', '\0\0+1']]);
    expect(text).toBe(`"'=cmd|calc","'+1"`);
  });

  it('M4: fields named like Object.prototype keys stay text when read', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = await wb.csv.read(['constructor,toString,valueOf,__proto__,hasOwnProperty\ntrue,#N/A,x,y,z\n']);
    expect(ws.getRow(1).values.slice(1)).toEqual(['constructor', 'toString', 'valueOf', '__proto__', 'hasOwnProperty']);
    expect(ws.getRow(2).values.slice(1)).toEqual([true, { error: '#N/A' }, 'x', 'y', 'z']);
  });

  it('M5: the 3-D reference check is linear in the formula length', () => {
    const ws = new ExcelJS.Workbook().addWorksheet('S');
    ws.getCell('A1').value = { formula: `SUM(${':'.repeat(200_000)})` };
    const start = performance.now();
    ws.spliceRows(1, 0, [1]);
    expect(performance.now() - start).toBeLessThan(2_000);
  });

  it('L5: quote: false refuses fields with a line break or the delimiter', async () => {
    const err = await caught(() => csvText([['a\nb']], { quote: false }));
    expect(err.code).toBe('XLV_E_INVALID_OPTION');
    await expect(caught(() => csvText([['a,b']], { quote: false }))).resolves.toMatchObject({ code: 'XLV_E_INVALID_OPTION' });
    expect(await csvText([['plain', '=1']], { quote: false })).toBe(`plain,'=1`);
  });
});
