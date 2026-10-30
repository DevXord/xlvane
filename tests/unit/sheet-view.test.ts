// Sheet tab colour and view settings (sheetPr/tabColor, sheetView, pane state="split"; ECMA-376
// Part 1 §18.3.1.82, §18.3.1.93, §18.3.1.87, §18.3.1.66; [MS-OI29500] §2.1.640, §2.1.659;
// docs/facts/sml-write.md W2-50, W2-51; docs/decisions.md D-A09).
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { editCells } from '../../packages/xlvane/src/edit';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { writeWorkbookStream } from '../../packages/xlvane/src/stream-write';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError } from '../helpers/expect-error';
import { buildXlsx } from '../helpers/xlsx-builder';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

function sheetXml(bytes: Uint8Array): string {
  const data = unzipSync(bytes)['xl/worksheets/sheet1.xml'];
  if (!data) throw new Error('missing sheet');
  return strFromU8(data);
}

function withSheet(inner: string): Uint8Array {
  return buildXlsx({ sheets: [{ name: 'S', data: `<worksheet xmlns="${NS}">${inner}<sheetData/></worksheet>` }] });
}

describe('writing', () => {
  it('writes the tab colour first, in sheetPr', () => {
    const wb = createWorkbook();
    wb.addSheet('S').tabColor('00b050').set('A1', 1);
    const xml = sheetXml(writeWorkbook(wb));
    expect(xml).toMatch(/<worksheet[^>]*><sheetPr><tabColor rgb="FF00B050"\/><\/sheetPr><dimension /);
    const theme = createWorkbook();
    theme.addSheet('S').tabColor({ theme: 4, tint: 0.4 });
    expect(sheetXml(writeWorkbook(theme))).toContain('<sheetPr><tabColor theme="4" tint="0.4"/></sheetPr>');
    const none = createWorkbook();
    none.addSheet('S').tabColor('FF0000').tabColor(null);
    expect(sheetXml(writeWorkbook(none))).not.toContain('sheetPr');
  });

  it('writes only the view attributes that differ from the defaults', () => {
    const plain = createWorkbook();
    plain.addSheet('S');
    expect(sheetXml(writeWorkbook(plain))).toContain('<sheetView tabSelected="1" workbookViewId="0"/>');
    const wb = createWorkbook();
    wb.addSheet('S').view({ zoom: 80, showGridLines: false, showRowColHeaders: false, rightToLeft: true, layout: 'pageLayout' });
    expect(sheetXml(writeWorkbook(wb))).toContain('<sheetView tabSelected="1" showGridLines="0" showRowColHeaders="0" rightToLeft="1" view="pageLayout" zoomScale="80" workbookViewId="0"/>');
  });

  it('writes a split window in 1/20 point, xSplit for the vertical bar', () => {
    const wb = createWorkbook();
    wb.addSheet('S').view({ split: { x: 100, y: 37.5 } });
    expect(sheetXml(writeWorkbook(wb))).toContain('<pane xSplit="2000" ySplit="750" activePane="bottomRight" state="split"/>');
    const rows = createWorkbook();
    rows.addSheet('S').view({ split: { x: 0, y: 30 } });
    expect(sheetXml(writeWorkbook(rows))).toContain('<pane ySplit="600" activePane="bottomLeft" state="split"/>');
  });

  it('a split replaces frozen panes and freezing replaces a split', () => {
    const wb = createWorkbook();
    const s = wb.addSheet('S').freeze({ rows: 1 }).view({ split: { y: 30 } });
    expect(s.data().freeze).toBeNull();
    s.freeze({ cols: 1 });
    expect(s.data().view?.split).toBeNull();
    expect(sheetXml(writeWorkbook(wb))).toContain('state="frozen"');
  });

  it('rejects invalid options with a hint', () => {
    const s = createWorkbook().addSheet('S');
    expectXlvError(() => s.view({ zoom: 9 }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s.view({ zoom: 401 }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s.view({ zoom: 50.5 }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s.view({ showGridLines: 'no' as never }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s.view({ layout: 'print' as never }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s.view({ split: { x: 0, y: 0 } }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s.view({ split: { x: -1, y: 5 } }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s.view({ split: { x: Number.NaN, y: 5 } }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s.view({ zom: 50 } as never), 'XLV_E_INVALID_OPTION', { option: 'zom' });
    expectXlvError(() => s.tabColor('red' as never), 'XLV_E_INVALID_STYLE');
  });

  it('the streaming writer takes the same settings', async () => {
    const stream = writeWorkbookStream({ sheets: [{ name: 'S', tabColor: 'FF0000', view: { zoom: 150, split: { x: 50, y: 0 } }, rows: [[1]] }] });
    const bytes = new Uint8Array(await new Response(stream).arrayBuffer());
    const sheet = readWorkbook(bytes).getSheet('S');
    expect(sheet.tabColor).toBe('FFFF0000');
    expect(sheet.view).toEqual({ zoom: 150, showGridLines: true, showRowColHeaders: true, rightToLeft: false, layout: 'normal', split: { x: 50, y: 0 }, topLeftCell: null, activeCell: null, zoomScaleNormal: 0, showRuler: true });
  });
});

describe('reading', () => {
  it('round-trips tab colour and view', () => {
    const wb = createWorkbook();
    wb.addSheet('S').tabColor({ indexed: 10 }).view({ zoom: 75, showGridLines: false, rightToLeft: true, layout: 'pageBreakPreview', split: { x: 20, y: 40 } });
    const sheet = readWorkbook(writeWorkbook(wb)).getSheet('S');
    expect(sheet.tabColor).toEqual({ indexed: 10 });
    expect(sheet.view).toEqual({ zoom: 75, showGridLines: false, showRowColHeaders: true, rightToLeft: true, layout: 'pageBreakPreview', split: { x: 20, y: 40 }, topLeftCell: null, activeCell: null, zoomScaleNormal: 0, showRuler: true });
    expect(sheet.freeze).toBeNull();
  });

  it('gives the defaults when the file sets nothing', () => {
    const sheet = readWorkbook(withSheet('')).getSheet('S');
    expect(sheet.tabColor).toBeNull();
    expect(sheet.view).toEqual({ zoom: 100, showGridLines: true, showRowColHeaders: true, rightToLeft: false, layout: 'normal', split: null, topLeftCell: null, activeCell: null, zoomScaleNormal: 0, showRuler: true });
  });

  it('reads a pane without state as a split, and frozen panes not as one', () => {
    const split = readWorkbook(withSheet('<sheetViews><sheetView workbookViewId="0"><pane xSplit="2310" ySplit="2070" topLeftCell="C1" activePane="bottomRight"/></sheetView></sheetViews>')).getSheet('S');
    expect(split.view.split).toEqual({ x: 115.5, y: 103.5 });
    expect(split.freeze).toBeNull();
    const frozen = readWorkbook(withSheet('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" state="frozen"/></sheetView></sheetViews>')).getSheet('S');
    expect(frozen.view.split).toBeNull();
    expect(frozen.freeze).toEqual({ rows: 1, cols: 0 });
  });

  it('uses the first sheet view, as for frozen panes', () => {
    const sheet = readWorkbook(withSheet('<sheetViews><sheetView zoomScale="50" workbookViewId="0"/><sheetView zoomScale="200" workbookViewId="1"/></sheetViews>')).getSheet('S');
    expect(sheet.view.zoom).toBe(50);
  });

  it('warns about invalid values and uses the defaults', () => {
    const wb = readWorkbook(withSheet('<sheetPr><tabColor rgb="red"/></sheetPr><sheetViews><sheetView zoomScale="5" showGridLines="maybe" view="print" workbookViewId="0"><pane xSplit="-3" state="split"/></sheetView></sheetViews>'));
    const sheet = wb.getSheet('S');
    expect(sheet.tabColor).toBeNull();
    expect(sheet.view).toEqual({ zoom: 100, showGridLines: true, showRowColHeaders: true, rightToLeft: false, layout: 'normal', split: null, topLeftCell: null, activeCell: null, zoomScaleNormal: 0, showRuler: true });
    expect(wb.warnings.filter((w) => w.code === 'XLV_W_INVALID_VALUE').reduce((n, w) => n + w.count, 0)).toBeGreaterThanOrEqual(4);
  });

  it('editCells keeps the tab colour and view', () => {
    const wb = createWorkbook();
    wb.addSheet('S').tabColor('00B0F0').view({ zoom: 90 }).set('A1', 1);
    const edited = editCells(writeWorkbook(wb), [{ sheet: 'S', cell: 'B2', value: 2 }]);
    const sheet = readWorkbook(edited.bytes).getSheet('S');
    expect(sheet.tabColor).toBe('FF00B0F0');
    expect(sheet.view.zoom).toBe(90);
  });
});
