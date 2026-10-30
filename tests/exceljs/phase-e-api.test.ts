// xlvane tests for the exceljs API completed in phase E (docs/decisions.md D-E04; X4-27 … X4-30):
// gradient fills, diagonal borders, font and alignment extras, protection, shared and array
// formulas, outline levels, sheet properties, more view settings, workbook windows and
// calculation, write options, model snapshots and the smaller API members.
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { ExcelJS } from './helpers';

const part = (bytes: Uint8Array, name: string): string => strFromU8(unzipSync(bytes)[name]!);

async function roundTrip(wb: InstanceType<typeof ExcelJS.Workbook>): Promise<InstanceType<typeof ExcelJS.Workbook>> {
  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.load(await wb.xlsx.writeBuffer());
  return wb2;
}

describe('styles', () => {
  it('round-trips gradient fills as exceljs reads them', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    const stops = [
      { position: 0, color: { argb: 'FFFFFFFF' } },
      { position: 1, color: { argb: 'FF4472C4' } },
    ];
    ws.getCell('A1').fill = { type: 'gradient', gradient: 'angle', degree: 45, stops };
    ws.getCell('A2').fill = { type: 'gradient', gradient: 'path', center: { left: 0.5, top: 0.5 }, stops };
    ws.getCell('A1').value = 1;
    ws.getCell('A2').value = 2;
    const bytes = new Uint8Array(await wb.xlsx.writeBuffer());
    expect(part(bytes, 'xl/styles.xml')).toContain('<gradientFill type="path" left="0.5" right="0.5" top="0.5" bottom="0.5">');
    const ws2 = (await roundTrip(wb)).getWorksheet('S')!;
    expect(ws2.getCell('A1').fill).toEqual({ type: 'gradient', gradient: 'angle', degree: 45, stops });
    expect(ws2.getCell('A2').fill).toEqual({ type: 'gradient', gradient: 'path', center: { left: 0.5, top: 0.5 }, stops });
  });

  it('round-trips diagonal borders, font extras, alignment extras and protection', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    const cell = ws.getCell('A1');
    cell.value = 'x';
    cell.border = { diagonal: { up: true, down: false, style: 'thin', color: { argb: 'FFFF0000' } } };
    cell.font = { name: 'Arial', size: 10, family: 2, scheme: 'minor', charset: 1, vertAlign: 'superscript', outline: true };
    cell.alignment = { shrinkToFit: true, readingOrder: 'rtl', textRotation: -45 };
    cell.protection = { locked: false, hidden: true };
    const bytes = new Uint8Array(await wb.xlsx.writeBuffer());
    expect(part(bytes, 'xl/styles.xml')).toContain('textRotation="135"');
    const c2 = (await roundTrip(wb)).getWorksheet('S')!.getCell('A1');
    expect(c2.border?.diagonal).toEqual({ style: 'thin', color: { argb: 'FFFF0000' }, up: true, down: false });
    expect(c2.font).toEqual({ name: 'Arial', size: 10, family: 2, scheme: 'minor', charset: 1, vertAlign: 'superscript', outline: true });
    expect(c2.alignment).toEqual({ shrinkToFit: true, readingOrder: 'rtl', textRotation: -45 });
    expect(c2.protection).toEqual({ locked: false, hidden: true });
  });

  it('converts textRotation "vertical" to 255 and back', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 'v';
    ws.getCell('A1').alignment = { textRotation: 'vertical' };
    expect((await roundTrip(wb)).getWorksheet('S')!.getCell('A1').alignment).toEqual({ textRotation: 'vertical' });
  });

  it('sets protection on rows and columns', () => {
    const ws = new ExcelJS.Workbook().addWorksheet('S');
    ws.getCell('A1').value = 1;
    ws.getRow(1).protection = { locked: false };
    expect(ws.getCell('A1').protection).toEqual({ locked: false });
    ws.getColumn(2).protection = { hidden: true };
    expect(ws.getColumn(2).protection).toEqual({ hidden: true });
  });
});

describe('shared and array formulas', () => {
  it('fillFormula writes a shared formula; dependents read back as { sharedFormula } with a translated formula', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 1;
    ws.getCell('A2').value = 2;
    ws.getCell('A3').value = 3;
    ws.fillFormula('B1:B3', 'A1*2', [2, 4, 6]);
    expect(ws.getCell('B2').value).toEqual({ sharedFormula: 'B1', result: 4 });
    expect(ws.getCell('B2').formula).toBe('A2*2');
    expect(ws.getCell('B1').formulaType).toBe(ExcelJS.FormulaType.Master);
    expect(ws.getCell('B2').formulaType).toBe(ExcelJS.FormulaType.Shared);
    expect(ws.getCell('A1').formulaType).toBe(ExcelJS.FormulaType.None);
    const bytes = new Uint8Array(await wb.xlsx.writeBuffer());
    expect(part(bytes, 'xl/worksheets/sheet1.xml')).toContain('<f t="shared" ref="B1:B3" si="0">A1*2</f>');
    const ws2 = (await roundTrip(wb)).getWorksheet('S')!;
    expect(ws2.getCell('B1').value).toEqual({ formula: 'A1*2', result: 2, ref: 'B1:B3', shareType: 'shared' });
    expect(ws2.getCell('B3').value).toEqual({ sharedFormula: 'B1', result: 6 });
    expect(ws2.getCell('B3').formula).toBe('A3*2');
  });

  it('grows the master range to { sharedFormula } dependents set by hand, as exceljs', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('C1').value = { formula: 'ROW()', result: 1 };
    ws.getCell('C2').value = { sharedFormula: 'C1', result: 2 };
    ws.getCell('D2').value = { sharedFormula: 'C1', result: 2 };
    const xml = part(new Uint8Array(await wb.xlsx.writeBuffer()), 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('<f t="shared" ref="C1:D2" si="0">ROW()</f>');
  });

  it('writes array formulas from fillFormula and { shareType: "array" }', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.fillFormula('A1:A2', 'ROW(A1:A2)', [1, 2], 'array');
    expect(ws.getCell('A2').value).toBe(2);
    const bytes = new Uint8Array(await wb.xlsx.writeBuffer());
    expect(part(bytes, 'xl/worksheets/sheet1.xml')).toContain('<f t="array" ref="A1:A2">ROW(A1:A2)</f>');
    expect((await roundTrip(wb)).getWorksheet('S')!.getCell('A1').value).toEqual({ formula: 'ROW(A1:A2)', result: 1, ref: 'A1:A2', shareType: 'array' });
  });

  it('accepts results as a function or a 2-D array', () => {
    const ws = new ExcelJS.Workbook().addWorksheet('S');
    ws.fillFormula('A1:B2', 'ROW()*COLUMN()', (r: number, c: number) => r * c);
    expect(ws.getCell('B2').value).toEqual({ sharedFormula: 'A1', result: 4 });
    ws.fillFormula('C1:D2', 'ROW()', [
      [1, 1],
      [2, 2],
    ]);
    expect(ws.getCell('D2').value).toEqual({ sharedFormula: 'C1', result: 2 });
  });

  it('the streaming writer writes dependents as translated formulas', async () => {
    const chunks: Uint8Array[] = [];
    const stream = {
      write: (c: Uint8Array, cb: () => void) => {
        chunks.push(new Uint8Array(c));
        cb();
      },
      end: (cb: () => void) => cb(),
    };
    const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ stream: stream as never });
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 1;
    ws.getCell('A2').value = 2;
    ws.fillFormula('B1:B2', 'A1+1', [2, 3]);
    ws.commit();
    await wb.commit();
    const bytes = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
    let at = 0;
    for (const c of chunks) {
      bytes.set(c, at);
      at += c.length;
    }
    const sheet = readWorkbook(bytes).getSheet('S');
    expect(sheet.cell('B2')?.formula?.text).toBe('A2+1');
  });
});

describe('outline levels and sheet properties', () => {
  it('writes row and column outline levels and reads them back', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S', { properties: { outlineLevelRow: 1, outlineProperties: { summaryBelow: false, summaryRight: true } } });
    ws.addRow(['h']);
    ws.addRow([1]).outlineLevel = 1;
    ws.getColumn(2).outlineLevel = 2;
    ws.properties['defaultRowHeight'] = 18;
    const bytes = new Uint8Array(await wb.xlsx.writeBuffer());
    const xml = part(bytes, 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('<outlinePr summaryBelow="0"/>');
    expect(xml).toContain('<sheetFormatPr defaultRowHeight="18" outlineLevelRow="1" outlineLevelCol="2"/>');
    expect(xml).toContain('<row r="2" outlineLevel="1" collapsed="1">');
    const ws2 = (await roundTrip(wb)).getWorksheet('S')!;
    expect(ws2.getRow(2).outlineLevel).toBe(1);
    expect(ws2.getColumn(2).outlineLevel).toBe(2);
    expect(ws2.properties).toMatchObject({ defaultRowHeight: 18, outlineLevelRow: 1, outlineLevelCol: 2, outlineProperties: { summaryBelow: false, summaryRight: true } });
  });

  it('computes collapsed and the column helpers as exceljs', () => {
    const ws = new ExcelJS.Workbook().addWorksheet('S');
    const col = ws.getColumn(1);
    expect(col.isDefault).toBe(true);
    col.outlineLevel = 1;
    expect(col.isDefault).toBe(false);
    expect(col.collapsed).toBe(true);
    ws.properties['outlineLevelCol'] = 2;
    expect(col.collapsed).toBe(false);
    col.header = ['A', 'B'];
    expect(col.headerCount).toBe(2);
    expect(JSON.parse(col.toString())).toEqual({ headers: ['A', 'B'] });
    expect(col.equivalentTo(ws.getColumn(2))).toBe(false);
  });
});

describe('views, workbook windows and calculation', () => {
  it('writes topLeftCell, activeCell, zoomScaleNormal and showRuler', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S', { views: [{ state: 'normal', topLeftCell: 'B5', activeCell: 'C6', zoomScaleNormal: 80, showRuler: false }] });
    ws.getCell('A1').value = 1;
    const xml = part(new Uint8Array(await wb.xlsx.writeBuffer()), 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('showRuler="0" topLeftCell="B5" zoomScaleNormal="80"');
    expect(xml).toContain('<selection activeCell="C6" sqref="C6"/>');
    // exceljs reads no topLeftCell for a normal view (sheet-view-xform.js L175–191).
    const view = (await roundTrip(wb)).getWorksheet('S')!.views[0]!;
    expect(view).toMatchObject({ state: 'normal', activeCell: 'C6', zoomScaleNormal: 80, showRuler: false });
  });

  it('writes workbook views and calcProperties.fullCalcOnLoad', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('A').getCell('A1').value = 1;
    wb.addWorksheet('B').getCell('A1').value = 2;
    wb.views = [{ x: 0, y: 0, width: 10000, height: 20000, firstSheet: 0, activeTab: 1, visibility: 'visible' }];
    wb.calcProperties.fullCalcOnLoad = true;
    const xml = part(new Uint8Array(await wb.xlsx.writeBuffer()), 'xl/workbook.xml');
    expect(xml).toContain('<workbookView xWindow="0" yWindow="0" windowWidth="10000" windowHeight="20000" firstSheet="0" activeTab="1"/>');
    expect(xml).toContain('<calcPr fullCalcOnLoad="1"/>');
    const wb2 = await roundTrip(wb);
    expect(wb2.views).toEqual([{ x: 0, y: 0, width: 10000, height: 20000, visibility: 'visible', firstSheet: 0, activeTab: 1 }]);
    expect(wb2.calcProperties).toEqual({ fullCalcOnLoad: true });
  });
});

describe('write options', () => {
  it('useSharedStrings: false writes inline strings; zip STORE writes stored parts', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('S').getCell('A1').value = 'hello';
    const inline = new Uint8Array(await wb.xlsx.writeBuffer({ useSharedStrings: false }));
    expect(unzipSync(inline)['xl/sharedStrings.xml']).toBeUndefined();
    const stored = new Uint8Array(await wb.xlsx.writeBuffer({ zip: { compression: 'STORE' } }));
    const deflated = new Uint8Array(await wb.xlsx.writeBuffer({ zip: { compressionOptions: { level: 9 } } }));
    expect(stored.length).toBeGreaterThan(deflated.length);
  });
});

describe('model snapshots and smaller members', () => {
  it('copies a workbook through model', () => {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Me';
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 'x';
    ws.getCell('B2').value = { formula: 'A1', result: 'x' };
    ws.mergeCells('C1:D1');
    const copy = new ExcelJS.Workbook();
    copy.model = wb.model;
    const ws2 = copy.getWorksheet('S')!;
    expect(copy.creator).toBe('Me');
    expect(ws2.getCell('A1').value).toBe('x');
    expect(ws2.getCell('B2').value).toEqual({ formula: 'A1', result: 'x' });
    expect(ws2.getCell('D1').isMergedTo(ws2.getCell('C1'))).toBe(true);
  });

  it('gives CSV strings as exceljs cell.toCsvString', () => {
    const ws = new ExcelJS.Workbook().addWorksheet('S');
    const v = (value: unknown): string => {
      const c = ws.getCell('A1');
      c.value = value as never;
      return c.toCsvString();
    };
    expect(v('a"b')).toBe('"a""b"');
    expect(v(1.5)).toBe('1.5');
    expect(v(true)).toBe('1');
    expect(v(new Date(Date.UTC(2024, 0, 2)))).toBe('2024-01-02T00:00:00.000Z');
    expect(v({ formula: 'A2', result: 7 })).toBe('7');
    expect(v({ text: 'site', hyperlink: 'https://example.com' })).toBe('https://example.com');
    expect(v(null)).toBe('');
  });

  it('has findRows, lastColumn, mergeCellsWithoutStyle, getCellEx, removeWorksheetEx, clearThemes and nextId', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.addRow([1]);
    expect(ws.findRows(1, 2).map((r: any) => r?.number)).toEqual([1, undefined]);
    ws.getColumn(3).width = 20;
    expect(ws.lastColumn.number).toBe(3);
    ws.getCell('A5').value = 'm';
    ws.getCell('A5').font = { bold: true };
    ws.getCell('B5').font = { italic: true };
    ws.mergeCellsWithoutStyle('A5:B5');
    expect(ws.getCell('B5').font).toEqual({ italic: true });
    expect(ws.getRow(1).getCellEx({ col: 2 }).address).toBe('B1');
    expect(wb.nextId).toBe(2);
    const other = wb.addWorksheet('T');
    wb.removeWorksheetEx(other);
    expect(wb.getWorksheet('T')).toBeUndefined();
    wb.clearThemes();
    expect(ws.getCell('A1').workbook).toBe(wb);
  });
});
