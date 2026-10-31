// Outline levels, sheet properties, more view settings, workbook windows, calculation settings
// and writer options (docs/decisions.md D-E04; facts W2-65 … W2-67 in docs/facts/sml-write.md).
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { writeWorkbookStream } from '../../packages/xlvane/src/stream-write';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError } from '../helpers/expect-error';
import { buildXlsx } from '../helpers/xlsx-builder';

const part = (bytes: Uint8Array, name: string): string => strFromU8(unzipSync(bytes)[name]!);

async function collect(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  for await (const c of stream as unknown as AsyncIterable<Uint8Array>) chunks.push(c);
  const out = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
  let at = 0;
  for (const c of chunks) {
    out.set(c, at);
    at += c.length;
  }
  return out;
}

describe('outline levels', () => {
  it('writes and reads row and column levels, collapsed flags and the sheet maxima', () => {
    const wb = createWorkbook();
    const s = wb.addSheet('S');
    s.set('A1', 'x').row(2, { outlineLevel: 1 }).row(3, { outlineLevel: 2, hidden: true }).row(4, { collapsed: true });
    s.column('B', { outlineLevel: 3, collapsed: true });
    const bytes = writeWorkbook(wb);
    const xml = part(bytes, 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('<sheetFormatPr defaultRowHeight="15" outlineLevelRow="2" outlineLevelCol="3"/>');
    expect(xml).toContain('<col min="2" max="2" outlineLevel="3" collapsed="1"/>');
    expect(xml).toContain('<row r="3" hidden="1" outlineLevel="2"/>');
    expect(xml).toContain('<row r="4" collapsed="1"/>');
    const sheet = readWorkbook(bytes).getSheet('S');
    expect(sheet.rowInfo(2)).toEqual({ height: null, hidden: false, outlineLevel: 1 });
    expect(sheet.rowInfo(4)).toEqual({ height: null, hidden: false, collapsed: true });
    expect(sheet.columns[0]).toMatchObject({ min: 2, max: 2, outlineLevel: 3, collapsed: true });
    expect(sheet.properties).toEqual({ defaultRowHeight: 15, outlineLevelRow: 2, outlineLevelCol: 3, summaryBelow: true, summaryRight: true });
  });

  it('refuses levels outside 0–7 ([MS-OI29500] §2.1.644 b, §2.1.598 e)', () => {
    const s = createWorkbook().addSheet('S');
    expectXlvError(() => s.row(1, { outlineLevel: 8 }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s.column('A', { outlineLevel: -1 }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s.row(1, { collapsed: 'yes' as never }), 'XLV_E_INVALID_OPTION');
  });
});

describe('sheet properties', () => {
  it('writes the default row height and summary positions, and reads them back', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', 1).properties({ defaultRowHeight: 18, summaryBelow: false, summaryRight: false, outlineLevelRow: 1 });
    const bytes = writeWorkbook(wb);
    const xml = part(bytes, 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('<sheetPr><outlinePr summaryBelow="0" summaryRight="0"/></sheetPr>');
    expect(xml).toContain('<sheetFormatPr defaultRowHeight="18" outlineLevelRow="1"/>');
    expect(readWorkbook(bytes).getSheet('S').properties).toEqual({ defaultRowHeight: 18, outlineLevelRow: 1, outlineLevelCol: 0, summaryBelow: false, summaryRight: false });
  });

  it('refuses invalid values', () => {
    const s = createWorkbook().addSheet('S');
    expectXlvError(() => s.properties({ defaultRowHeight: -1 }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s.properties({ summaryBelow: 1 as never }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s.properties({ outlineLevelCol: 9 }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s.properties({ other: 1 } as never), 'XLV_E_INVALID_OPTION');
  });
});

describe('more view settings', () => {
  it('writes topLeftCell, activeCell, zoomScaleNormal and showRuler, and reads them back', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', 1).view({ topLeftCell: 'c10', activeCell: 'D12', zoomScaleNormal: 90, showRuler: false });
    const bytes = writeWorkbook(wb);
    const xml = part(bytes, 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('<sheetView tabSelected="1" showRuler="0" topLeftCell="C10" zoomScaleNormal="90" workbookViewId="0"><selection activeCell="D12" sqref="D12"/></sheetView>');
    expect(readWorkbook(bytes).getSheet('S').view).toMatchObject({ topLeftCell: 'C10', activeCell: 'D12', zoomScaleNormal: 90, showRuler: false });
  });

  it('puts the active cell into the selection of frozen panes', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', 1).freeze({ rows: 1 }).view({ activeCell: 'B5', topLeftCell: 'A20' });
    const xml = part(writeWorkbook(wb), 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="B5" sqref="B5"/>');
    expect(xml).not.toContain('topLeftCell="A20"');
  });

  it('refuses invalid values', () => {
    const s = createWorkbook().addSheet('S');
    expectXlvError(() => s.view({ zoomScaleNormal: 5 }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s.view({ activeCell: 'not a cell' }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s.view({ showRuler: 0 as never }), 'XLV_E_INVALID_OPTION');
  });
});

describe('workbook windows and calculation', () => {
  it('writes bookViews and calcPr in schema order and reads them back', () => {
    const wb = createWorkbook().views([{ x: 10, y: 20, width: 12000, height: 8000, firstSheet: 1, activeTab: 1 }]).calculation({ fullCalcOnLoad: true });
    wb.addSheet('A').set('A1', 1);
    wb.addSheet('B').set('A1', 2);
    wb.defineName('Total', 'A!$A$1');
    const bytes = writeWorkbook(wb);
    const xml = part(bytes, 'xl/workbook.xml');
    expect(xml).toContain('<bookViews><workbookView xWindow="10" yWindow="20" windowWidth="12000" windowHeight="8000" firstSheet="1" activeTab="1"/></bookViews>');
    expect(xml).toMatch(/<\/definedNames><calcPr fullCalcOnLoad="1"\/><\/workbook>$/);
    const read = readWorkbook(bytes);
    expect(read.views).toEqual([{ x: 10, y: 20, width: 12000, height: 8000, firstSheet: 1, activeTab: 1, visibility: 'visible' }]);
    expect(read.calculation).toEqual({ fullCalcOnLoad: true });
  });

  it('keeps the default window and no calcPr when nothing is set', () => {
    const wb = createWorkbook();
    wb.addSheet('A').set('A1', 1);
    const bytes = writeWorkbook(wb);
    expect(part(bytes, 'xl/workbook.xml')).not.toContain('calcPr');
    const read = readWorkbook(bytes);
    expect(read.views).toEqual([{ x: null, y: null, width: null, height: null, firstSheet: null, activeTab: null, visibility: 'visible' }]);
    expect(read.calculation).toBeNull();
  });

  it('refuses invalid windows', () => {
    const wb = createWorkbook();
    expectXlvError(() => wb.views([{ width: -1 }]), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => wb.views([{ visibility: 'gone' as never }]), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => wb.views({} as never), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => wb.calculation({ fullCalcOnLoad: 1 as never }), 'XLV_E_INVALID_OPTION');
  });
});

describe('writer options', () => {
  it('writes inline strings without a shared string table when sharedStrings is false', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', 'hello');
    const bytes = writeWorkbook(wb, { sharedStrings: false });
    const files = unzipSync(bytes);
    expect(files['xl/sharedStrings.xml']).toBeUndefined();
    expect(strFromU8(files['xl/worksheets/sheet1.xml']!)).toContain('t="inlineStr"');
    expect(strFromU8(files['[Content_Types].xml']!)).not.toContain('sharedStrings');
    expect(strFromU8(files['xl/_rels/workbook.xml.rels']!)).not.toContain('sharedStrings');
    expect(readWorkbook(bytes).getSheet('S').cell('A1')?.value).toBe('hello');
  });

  it('stores every part with compression level 0 and keeps the output deterministic', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', 'x'.repeat(1000));
    const stored = writeWorkbook(wb, { compressionLevel: 0 });
    const best = writeWorkbook(wb, { compressionLevel: 9 });
    expect(stored.length).toBeGreaterThan(best.length);
    expect(writeWorkbook(wb, { compressionLevel: 0 })).toEqual(stored);
    expect(readWorkbook(stored).getSheet('S').cell('A1')?.value).toBe('x'.repeat(1000));
    expectXlvError(() => writeWorkbook(wb, { compressionLevel: 10 }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => writeWorkbook(wb, { sharedStrings: 'no' as never }), 'XLV_E_INVALID_OPTION');
  });
});

describe('streaming writer', () => {
  it('takes row outline levels, sheet properties, windows, calculation and a compression level', async () => {
    const bytes = await collect(
      writeWorkbookStream({
        views: [{ activeTab: 0, width: 9000 }],
        calculation: { fullCalcOnLoad: true },
        compressionLevel: 1,
        sheets: [{ name: 'S', properties: { outlineLevelRow: 1, summaryBelow: false }, view: { activeCell: 'A2' }, rows: [['h'], { values: [1], outlineLevel: 1, collapsed: true }] }],
      }),
    );
    const xml = part(bytes, 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('<sheetFormatPr defaultRowHeight="15" outlineLevelRow="1"/>');
    expect(xml).toContain('<row r="2" outlineLevel="1" collapsed="1">');
    const read = readWorkbook(bytes);
    expect(read.views[0]).toMatchObject({ width: 9000 });
    expect(read.calculation).toEqual({ fullCalcOnLoad: true });
    expect(read.getSheet('S').view.activeCell).toBe('A2');
    expect(read.getSheet('S').properties.summaryBelow).toBe(false);
  });
});

describe('review fixes (D-E06): active cell of a view with several panes', () => {
  it('is the selection of the active pane, as Excel writes frozen views', () => {
    const view = (inner: string): string | null | undefined =>
      readWorkbook(buildXlsx({ sheets: [{ name: 'S', before: `<sheetViews><sheetView workbookViewId="0">${inner}</sheetView></sheetViews>`, data: '' }] })).getSheet('S').view.activeCell;
    expect(
      view(
        '<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>' +
          '<selection pane="topRight" activeCell="B1" sqref="B1"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/><selection pane="bottomRight" activeCell="D5" sqref="D5"/>',
      ),
    ).toBe('D5');
    expect(view('<selection activeCell="C3" sqref="C3"/>')).toBe('C3');
  });
});

describe('streamed workbook windows (D-R03)', () => {
  it('refuses activeTab beyond the streamed sheets', async () => {
    const { writeWorkbookStream } = await import('../../packages/xlvane/src/stream-write');
    const stream = writeWorkbookStream({ views: [{ activeTab: 5 }], sheets: [{ name: 'S', rows: [[1]] }] });
    let code = '';
    try {
      const reader = stream.getReader();
      for (;;) if ((await reader.read()).done) break;
    } catch (error) {
      code = (error as { code?: string }).code ?? 'CRASH';
    }
    expect(code).toBe('XLV_E_INVALID_OPTION');
  });
});
