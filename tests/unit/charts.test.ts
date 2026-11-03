// Charts on worksheets: chart parts (ECMA-376 Part 1 §21.2) in the sheet's drawing;
// docs/facts/charts.md CH-01 … CH-35, docs/decisions.md D-K01.
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { CHART_PALETTE } from '../../packages/xlvane/src/charts';
import { editCells } from '../../packages/xlvane/src/edit';
import { inspect } from '../../packages/xlvane/src/inspect';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { insertRows } from '../../packages/xlvane/src/rows';
import { writeWorkbookStream } from '../../packages/xlvane/src/stream-write';
import { createWorkbook, writeWorkbook, type NewSheet } from '../../packages/xlvane/src/write';
import type { ChartInput } from '../../packages/xlvane/src/charts';
import { expectXlvError, expectXlvErrorAsync } from '../helpers/expect-error';

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0]);

function parts(bytes: Uint8Array): Record<string, string> {
  return Object.fromEntries(Object.entries(unzipSync(bytes)).map(([k, v]) => [k, strFromU8(v)]));
}

function patch(bytes: Uint8Array, edit: (name: string, text: string) => string | null): Uint8Array {
  const out: Record<string, Uint8Array> = {};
  for (const [name, data] of Object.entries(unzipSync(bytes))) {
    const t = edit(name, strFromU8(data));
    if (t !== null) out[name] = strToU8(t);
  }
  return zipSync(out);
}

async function bytesOf(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

/** Sheet "Data": A1:C4 = Month/Sales/Costs, Jan…Mar, 10/20/30, 5/6/7. */
function dataSheet(sheet: NewSheet): NewSheet {
  sheet.set('A1', 'Month').set('B1', 'Sales').set('C1', 'Costs');
  ['Jan', 'Feb', 'Mar'].forEach((m, i) => sheet.set(`A${String(i + 2)}`, m).set(`B${String(i + 2)}`, (i + 1) * 10).set(`C${String(i + 2)}`, i + 5));
  return sheet;
}

function written(chart: Omit<ChartInput, 'anchor'> & { anchor?: ChartInput['anchor'] }): Record<string, string> {
  const wb = createWorkbook();
  dataSheet(wb.addSheet('Data')).addChart({ anchor: 'E2:L18', ...chart });
  return parts(writeWorkbook(wb));
}

const SALES = { name: 'Data!B1', categories: 'Data!A2:A4', values: 'Data!B2:B4' };
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:roundedCorners val="0"/><c:chart>';
const TAIL = '<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>';
const NAME = '<c:tx><c:strRef><c:f>Data!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Sales</c:v></c:pt></c:strCache></c:strRef></c:tx>';
const strPts = (v: string[]): string => v.map((s, i) => `<c:pt idx="${String(i)}"><c:v>${s}</c:v></c:pt>`).join('');
const CAT = `<c:strRef><c:f>Data!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/>${strPts(['Jan', 'Feb', 'Mar'])}</c:strCache></c:strRef>`;
const VAL = `<c:numRef><c:f>Data!$B$2:$B$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/>${strPts(['10', '20', '30'])}</c:numCache></c:numRef>`;
const FILL = (c: string): string => `<a:solidFill><a:srgbClr val="${c}"/></a:solidFill>`;
const AX = '<c:axId val="500000001"/><c:axId val="500000002"/>';
const title = (t: string, rot = false): string => `<c:title><c:tx><c:rich><a:bodyPr${rot ? ' rot="-5400000" vert="horz"' : ''}/><a:lstStyle/><a:p><a:r><a:t>${t}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
const catAx = (t = ''): string =>
  `<c:catAx><c:axId val="500000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>${t}<c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>`;
const valAx = (between: string, t = '', scale = '', id = '500000002', cross = '500000001', pos = 'l', grid = true): string =>
  `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/>${scale}</c:scaling><c:delete val="0"/><c:axPos val="${pos}"/>${grid ? '<c:majorGridlines/>' : ''}${t}<c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="${cross}"/><c:crosses val="autoZero"/><c:crossBetween val="${between}"/></c:valAx>`;
const chart = (plot: string, head = '<c:autoTitleDeleted val="1"/>', tail = TAIL): string => `${HEAD}${head}<c:plotArea><c:layout/>${plot}</c:plotArea>${tail}`;

describe('writing charts: exact XML (CH-04 … CH-28)', () => {
  it('column: barChart col, clustered, filled series with caches, catAx and valAx', () => {
    const p = written({ type: 'column', title: 'Sales', series: [SALES] });
    expect(p['xl/charts/chart1.xml']).toBe(
      chart(
        `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/>${NAME}<c:spPr>${FILL('5B9BD5')}</c:spPr><c:cat>${CAT}</c:cat><c:val>${VAL}</c:val></c:ser><c:gapWidth val="150"/>${AX}</c:barChart>${catAx()}${valAx('between')}`,
        `${title('Sales')}<c:autoTitleDeleted val="0"/>`,
      ),
    );
  });

  it('bar: barDir bar, stacked with overlap 100, rotated category axis title', () => {
    const p = written({ type: 'bar', stacked: true, series: [SALES], axes: { x: { title: 'Month' }, y: { title: 'EUR' } } });
    expect(p['xl/charts/chart1.xml']).toBe(
      chart(`<c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/>${NAME}<c:spPr>${FILL('5B9BD5')}</c:spPr><c:cat>${CAT}</c:cat><c:val>${VAL}</c:val></c:ser><c:gapWidth val="150"/><c:overlap val="100"/>${AX}</c:barChart>${catAx(title('Month', true))}${valAx('between', title('EUR'))}`),
    );
  });

  it('line: lineChart, coloured line without markers, smooth 0, marker 1, scale max before min', () => {
    const p = written({ type: 'line', stacked: 'percent', series: [{ ...SALES, color: '#c00000' }], legend: 'bottom', axes: { y: { min: 0, max: 50 } } });
    expect(p['xl/charts/chart1.xml']).toBe(
      chart(
        `<c:lineChart><c:grouping val="percentStacked"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/>${NAME}<c:spPr><a:ln>${FILL('C00000')}</a:ln></c:spPr><c:marker><c:symbol val="none"/></c:marker><c:cat>${CAT}</c:cat><c:val>${VAL}</c:val><c:smooth val="0"/></c:ser><c:marker val="1"/>${AX}</c:lineChart>${catAx()}${valAx('between', '', '<c:max val="50"/><c:min val="0"/>')}`,
        '<c:autoTitleDeleted val="1"/>',
        TAIL.replace('val="r"', 'val="b"'),
      ),
    );
  });

  it('pie: varyColors 1, one dPt per value with the palette, no axes', () => {
    const p = written({ type: 'pie', series: [SALES], legend: false });
    const points = [0, 1, 2].map((k) => `<c:dPt><c:idx val="${String(k)}"/><c:spPr>${FILL(CHART_PALETTE[k]!)}</c:spPr></c:dPt>`).join('');
    expect(p['xl/charts/chart1.xml']).toBe(
      chart(
        `<c:pieChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/>${NAME}${points}<c:cat>${CAT}</c:cat><c:val>${VAL}</c:val></c:ser><c:firstSliceAng val="0"/></c:pieChart>`,
        '<c:autoTitleDeleted val="1"/>',
        '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>',
      ),
    );
  });

  it('scatter: xVal and yVal as numRef, markers only, two value axes with midCat', () => {
    const p = written({ type: 'scatter', series: [{ name: 'Sales', categories: 'Data!C2:C4', values: 'Data!B2:B4' }], legend: 'left' });
    const x = `<c:numRef><c:f>Data!$C$2:$C$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/>${strPts(['5', '6', '7'])}</c:numCache></c:numRef>`;
    expect(p['xl/charts/chart1.xml']).toBe(
      chart(
        `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Sales</c:v></c:tx><c:spPr><a:ln><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr>${FILL('5B9BD5')}<a:ln>${FILL('5B9BD5')}</a:ln></c:spPr></c:marker><c:xVal>${x}</c:xVal><c:yVal>${VAL}</c:yVal><c:smooth val="0"/></c:ser>${AX}</c:scatterChart>` +
          valAx('midCat', '', '', '500000001', '500000002', 'b', false) +
          valAx('midCat'),
        '<c:autoTitleDeleted val="1"/>',
        TAIL.replace('val="r"', 'val="l"'),
      ),
    );
  });

  it('area: areaChart standard, filled series, valAx midCat', () => {
    const p = written({ type: 'area', series: [SALES], legend: 'top' });
    expect(p['xl/charts/chart1.xml']).toBe(
      chart(`<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/>${NAME}<c:spPr>${FILL('5B9BD5')}</c:spPr><c:cat>${CAT}</c:cat><c:val>${VAL}</c:val></c:ser>${AX}</c:areaChart>${catAx()}${valAx('midCat')}`, '<c:autoTitleDeleted val="1"/>', TAIL.replace('val="r"', 'val="t"')),
    );
  });

  it('writes the drawing graphicFrame as Excel does (CH-04), the chart relationship and the content type (CH-01, CH-02)', () => {
    const wb = createWorkbook();
    const img = wb.addImage(PNG);
    dataSheet(wb.addSheet('Data'))
      .addImage(img, 'A10:B12')
      .addChart({ type: 'column', anchor: 'E2:L18', series: [SALES], name: 'Sales chart', description: 'Sales by month' })
      .addChart({ type: 'line', anchor: { from: 'E20', size: { width: 100, height: 50 } }, series: [SALES] });
    const p = parts(writeWorkbook(wb));
    const d = p['xl/drawings/drawing1.xml']!;
    expect(d).toContain(
      '<xdr:twoCellAnchor><xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>12</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>18</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
        '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Sales chart" descr="Sales by month"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId2"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>',
    );
    expect(d).toContain('<xdr:oneCellAnchor><xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>19</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="952500" cy="476250"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="3" name="Chart 2"/>');
    expect(d.indexOf('<xdr:pic>')).toBeLessThan(d.indexOf('<xdr:graphicFrame'));
    const rels = p['xl/drawings/_rels/drawing1.xml.rels']!;
    expect(rels).toContain('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>');
    expect(rels).toContain('<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>');
    expect(rels).toContain('<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>');
    expect(p['[Content_Types].xml']).toContain('<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>');
    expect(p['[Content_Types].xml']).toContain('<Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>');
    // No Office extensions, style or colour parts (CH-03, CH-32).
    expect(Object.keys(p).filter((n) => /style|colors|theme/.test(n))).toEqual(['xl/styles.xml']);
    expect(Object.values(p).join('')).not.toMatch(/c14|c16|a16|mc:AlternateContent/);
  });

  it('numbers chart parts across the workbook and writes a drawing for a sheet with charts only', () => {
    const wb = createWorkbook();
    dataSheet(wb.addSheet('Data')).addChart({ type: 'pie', anchor: 'E2:H9', series: [SALES] });
    wb.addSheet('Report').addChart({ type: 'column', anchor: 'A1:F10', series: [SALES] }).addChart({ type: 'area', anchor: 'A12:F20', series: [SALES] });
    const p = parts(writeWorkbook(wb));
    expect(Object.keys(p).filter((n) => n.startsWith('xl/charts/')).sort()).toEqual(['xl/charts/chart1.xml', 'xl/charts/chart2.xml', 'xl/charts/chart3.xml']);
    expect(p['xl/drawings/_rels/drawing2.xml.rels']).toContain('Target="../charts/chart2.xml"');
    expect(p['xl/drawings/_rels/drawing2.xml.rels']).toContain('Target="../charts/chart3.xml"');
    expect(p['xl/worksheets/sheet2.xml']).toMatch(/<drawing r:id="rId1"\/><\/worksheet>$/);
    expect(p['xl/charts/chart2.xml']).toContain('<c:f>Data!$B$2:$B$4</c:f>');
  });
});

describe('writing charts: references, palette and caches', () => {
  it('quotes sheet names only when needed and writes absolute references (CH-18, F2-11)', () => {
    const wb = createWorkbook();
    dataSheet(wb.addSheet('Sales Data'));
    dataSheet(wb.addSheet('A1'));
    dataSheet(wb.addSheet("O'Brien"));
    dataSheet(wb.addSheet('TRUE'));
    wb.addSheet('Charts').addChart({
      type: 'line',
      anchor: 'A1:H10',
      series: [
        { name: "'sales data'!B1", categories: "'Sales Data'!A2:A4", values: "='Sales Data'!$B$2:B4" },
        { values: "'A1'!C4:C2" },
        { values: "'O''Brien'!B2:B4" },
        { values: "'TRUE'!B2:B4" },
      ],
    });
    const xml = parts(writeWorkbook(wb))['xl/charts/chart1.xml']!;
    expect([...xml.matchAll(/<c:f>([^<]*)<\/c:f>/g)].map((m) => m[1])).toEqual([
      '&apos;Sales Data&apos;!$B$1'.replace(/&apos;/g, "'"),
      "'Sales Data'!$A$2:$A$4",
      "'Sales Data'!$B$2:$B$4",
      "'A1'!$C$2:$C$4",
      "'O''Brien'!$B$2:$B$4",
      "'TRUE'!$B$2:$B$4",
    ]);
  });

  it('gives series the palette in order, starting again after six (CH-35), and keeps color', () => {
    const series = Array.from({ length: 8 }, (_, i) => ({ values: 'Data!B2:B4', ...(i === 2 ? { color: '123abc' } : {}) }));
    const xml = written({ type: 'column', series })['xl/charts/chart1.xml']!;
    expect([...xml.matchAll(/<c:ser>.*?<a:srgbClr val="([0-9A-F]{6})"/g)].map((m) => m[1])).toEqual(['5B9BD5', 'ED7D31', '123ABC', 'FFC000', '4472C4', '70AD47', '5B9BD5', 'ED7D31']);
  });

  it('fills caches from the cell values: number formats, empty cells, formulas, numeric categories (CH-19)', () => {
    const wb = createWorkbook();
    const s = wb.addSheet('Data');
    s.set('A1', 2024).set('A2', 2025).set('A4', 2027);
    s.set('B1', 0.25, { numFmt: '0%' }).set('B2', { formula: 'B1*2', result: 0.5 }).set('B3', 'n/a').set('B4', 1e21);
    s.set('C1', 'Label & <more>').set('D1', true);
    s.addChart({ type: 'column', anchor: 'E2:L18', series: [{ name: 'Data!C1', categories: 'Data!A1:A4', values: 'Data!B1:B4' }, { name: 'Data!D1', values: 'Data!B1:B2' }, { name: 'Data!E1', values: 'Data!B1:B2' }] });
    const xml = parts(writeWorkbook(wb))['xl/charts/chart1.xml']!;
    expect(xml).toContain('<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Label &amp; &lt;more&gt;</c:v></c:pt></c:strCache>');
    expect(xml).toContain('<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>TRUE</c:v></c:pt></c:strCache>');
    expect(xml).toContain('<c:f>Data!$E$1</c:f><c:strCache><c:ptCount val="1"/></c:strCache>');
    expect(xml).toContain('<c:cat><c:numRef><c:f>Data!$A$1:$A$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="4"/><c:pt idx="0"><c:v>2024</c:v></c:pt><c:pt idx="1"><c:v>2025</c:v></c:pt><c:pt idx="3"><c:v>2027</c:v></c:pt></c:numCache></c:numRef></c:cat>');
    expect(xml).toContain('<c:numCache><c:formatCode>0%</c:formatCode><c:ptCount val="4"/><c:pt idx="0"><c:v>0.25</c:v></c:pt><c:pt idx="1"><c:v>0.5</c:v></c:pt><c:pt idx="3"><c:v>1E21</c:v></c:pt></c:numCache>');
  });

  it('streams charts without caches and checks sheets at the end of the stream (CH-20)', async () => {
    const rows = [['Month', 'Sales'], ['Jan', 10], ['Feb', 20]];
    const bytes = await bytesOf(writeWorkbookStream({ sheets: [{ name: 'Data', rows, charts: [{ type: 'pie', anchor: 'D2:H10', series: [{ name: 'Data!B1', categories: 'Data!A2:A3', values: 'Data!B2:B3' }] }] }] }));
    const p = parts(bytes);
    const xml = p['xl/charts/chart1.xml']!;
    expect(xml).toContain('<c:tx><c:strRef><c:f>Data!$B$1</c:f></c:strRef></c:tx>');
    expect(xml).toContain('<c:cat><c:strRef><c:f>Data!$A$2:$A$3</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>Data!$B$2:$B$3</c:f></c:numRef></c:val>');
    expect(xml).not.toContain('Cache');
    expect(p['[Content_Types].xml']).toContain('/xl/charts/chart1.xml');
    expect(readWorkbook(bytes).getSheet('Data').charts[0]?.types).toEqual(['pie']);
    // A chart over a sheet that comes later in an iterable of sheets.
    function* sheets(): Generator<{ name: string; rows: unknown[][]; charts?: ChartInput[] }> {
      yield { name: 'Charts', rows: [], charts: [{ type: 'column', anchor: 'A1:F10', series: [{ values: "'Later'!A1:A2" }] }] };
      yield { name: 'Later', rows: [[1], [2]] };
    }
    const later = readWorkbook(await bytesOf(writeWorkbookStream({ sheets: sheets() as never })));
    expect(later.getSheet('Charts').charts[0]?.series[0]?.values).toBe('Later!$A$1:$A$2');
    function* missing(): Generator<{ name: string; rows: unknown[][]; charts?: ChartInput[] }> {
      yield { name: 'Charts', rows: [], charts: [{ type: 'column', anchor: 'A1:F10', series: [{ values: 'Nowhere!A1:A2' }] }] };
    }
    await expectXlvErrorAsync(() => bytesOf(writeWorkbookStream({ sheets: missing() as never })), 'XLV_E_INVALID_CHART');
    expectXlvError(() => writeWorkbookStream({ sheets: [{ name: 'Data', rows, charts: [{ type: 'pie', anchor: 'D2:H10', series: [{ values: 'Other!B2:B3' }] }] }] }), 'XLV_E_INVALID_CHART');
  });
});

describe('chart validation (XLV_E_INVALID_CHART)', () => {
  const sheet = (): NewSheet => dataSheet(createWorkbook().addSheet('Data'));
  const bad = (chart: unknown, option?: string): void => {
    const error = expectXlvError(() => sheet().addChart(chart as ChartInput), 'XLV_E_INVALID_CHART');
    expect(error.location.sheet).toBe('Data');
    expect(error.hint).not.toBe('');
    if (option !== undefined) expect(error.details['option']).toBe(option);
  };
  const ok = { type: 'column', anchor: 'E2:L18', series: [SALES] };

  it('accepts a valid chart', () => {
    expect(() => sheet().addChart(ok as ChartInput)).not.toThrow();
  });

  it('refuses invalid types, series, colours, titles, legends, stacking and axes', () => {
    bad(null, 'chart');
    bad({ ...ok, kind: 'x' }, 'kind');
    bad({ ...ok, type: 'radar' }, 'type');
    bad({ ...ok, type: '__proto__' }, 'type');
    bad({ ...ok, series: [] }, 'series');
    bad({ ...ok, series: Array.from({ length: 256 }, () => SALES) }, 'series');
    bad({ ...ok, series: [{ values: 'B2:B4' }] }, 'values');
    bad({ ...ok, series: [{ values: 'Data!B2:C4' }] }, 'values');
    bad({ ...ok, series: [{ values: '[1]Data!B2:B4' }] }, 'values');
    bad({ ...ok, series: [{ values: 'Jan:Mar!B2:B4' }] }, 'values');
    bad({ ...ok, series: [{ values: 'Data!B2:B4,Data!C2' }] }, 'values');
    bad({ ...ok, series: [{ values: 'Data!B:B' }] }, 'values');
    bad({ ...ok, series: [{ values: 42 }] }, 'values');
    bad({ ...ok, series: [{ ...SALES, categories: 'Data!A2:B4' }] }, 'categories');
    bad({ ...ok, series: [{ ...SALES, name: 'Data!B1:C1' }] }, 'name');
    bad({ ...ok, series: [{ ...SALES, name: 'x'.repeat(256) }] }, 'name');
    bad({ ...ok, series: [{ ...SALES, name: 'a\u0001b' }] }, 'name');
    bad({ ...ok, series: [{ ...SALES, color: 'red' }] }, 'color');
    bad({ ...ok, series: [{ ...SALES, colour: 'FF0000' }] }, 'colour');
    bad({ ...ok, title: '' }, 'title');
    bad({ ...ok, title: 'x'.repeat(256) }, 'title');
    bad({ ...ok, title: 'a\nb' }, 'title');
    bad({ ...ok, legend: 'middle' }, 'legend');
    bad({ ...ok, stacked: 'yes' }, 'stacked');
    bad({ ...ok, type: 'scatter', stacked: true }, 'stacked');
    bad({ ...ok, type: 'pie', stacked: 'percent' }, 'stacked');
    bad({ ...ok, type: 'pie', series: [SALES, SALES] }, 'series');
    bad({ ...ok, type: 'pie', series: [{ ...SALES, color: 'FF0000' }] }, 'color');
    bad({ ...ok, type: 'pie', series: [{ values: 'Data!B1:B1001' }] }, 'values');
    bad({ ...ok, type: 'pie', axes: {} }, 'axes');
    bad({ ...ok, axes: { z: {} } }, 'z');
    bad({ ...ok, axes: { x: { min: 1 } } }, 'min');
    bad({ ...ok, axes: { y: { min: Number.NaN } } }, 'min');
    bad({ ...ok, axes: { y: { min: 5, max: 5 } } }, 'min');
    bad({ ...ok, axes: { y: { title: 7 } } }, 'axis title');
    bad({ ...ok, name: '' }, 'name');
    bad({ ...ok, description: 'a\u0000' }, 'description');
  });

  it('refuses invalid anchors with the chart code', () => {
    bad({ ...ok, anchor: undefined }, 'anchor');
    bad({ ...ok, anchor: 'nowhere' }, 'anchor');
    bad({ ...ok, anchor: { from: 'E8', to: 'B2' } }, 'to');
    const error = expectXlvError(() => sheet().addChart({ ...ok, anchor: 'X' } as ChartInput), 'XLV_E_INVALID_CHART');
    expect(error.message).toContain('chart');
  });

  it('refuses references to sheets the workbook does not have when it is written', () => {
    const wb = createWorkbook();
    dataSheet(wb.addSheet('Data')).addChart({ type: 'line', anchor: 'E2:L18', series: [{ values: 'Missing!B2:B4' }] });
    const error = expectXlvError(() => writeWorkbook(wb), 'XLV_E_INVALID_CHART');
    expect(error.message).toContain('"Missing"');
    expect(error.hint).toContain('"Data"');
    // Added later: the chart is valid once the sheet exists.
    wb.addSheet('Missing');
    expect(() => writeWorkbook(wb)).not.toThrow();
  });
});

describe('reading charts', () => {
  const cases: [string, ChartInput][] = [
    ['column', { type: 'column', anchor: 'E2:L18', title: 'Sales', series: [SALES, { name: 'Costs', categories: 'Data!A2:A4', values: 'Data!C2:C4' }], axes: { x: { title: 'Month' }, y: { title: 'EUR', min: 0, max: 40 } } }],
    ['column stacked', { type: 'column', anchor: 'E2:L18', stacked: true, series: [SALES], legend: 'bottom' }],
    ['bar percent', { type: 'bar', anchor: 'E2:L18', stacked: 'percent', series: [SALES], legend: false }],
    ['bar', { type: 'bar', anchor: 'E2:L18', series: [SALES], axes: { x: { title: 'Month' } } }],
    ['line', { type: 'line', anchor: { from: 'E2', to: 'L18', editAs: 'oneCell' }, series: [SALES], legend: 'top', axes: { y: { max: 100 } } }],
    ['line stacked', { type: 'line', anchor: 'E2:L18', stacked: true, series: [SALES] }],
    ['pie', { type: 'pie', anchor: { from: 'E2', size: { width: 300, height: 200 } }, title: 'Share', series: [SALES], legend: 'left' }],
    ['scatter', { type: 'scatter', anchor: 'E2:L18', series: [{ name: 'Data!B1', categories: 'Data!C2:C4', values: 'Data!B2:B4' }], axes: { x: { title: 'Costs' }, y: { title: 'Sales', min: -5 } } }],
    ['area', { type: 'area', anchor: 'E2:L18', series: [SALES] }],
    ['area percent', { type: 'area', anchor: 'E2:L18', stacked: 'percent', series: [SALES, { values: 'Data!C2:C4' }] }],
  ];

  it.each(cases)('reads back what was written: %s', (_, input) => {
    const wb = createWorkbook();
    dataSheet(wb.addSheet('Data')).addChart({ ...input, name: 'My chart', description: 'Alt text' });
    const back = readWorkbook(writeWorkbook(wb));
    expect(back.warnings).toEqual([]);
    const [c] = back.getSheet('Data').charts;
    const ref = (t: string | undefined): string | null => (t === undefined ? null : t.replace(/([A-Z]+)(\d+)/g, '$$$1$$$2'));
    expect(c).toMatchObject({ name: 'My chart', description: 'Alt text', types: [input.type] });
    expect(c?.stacked).toBe(input.type === 'pie' || input.type === 'scatter' ? null : (input.stacked ?? false));
    expect(c?.title).toEqual(input.title === undefined ? null : { text: input.title });
    expect(c?.legend).toBe(input.legend === undefined ? 'right' : input.legend === false ? null : input.legend);
    expect(c?.series).toEqual(
      input.series.map((s) => ({
        type: input.type,
        name: s.name === undefined ? null : s.name.includes('!') ? { ref: ref(s.name) } : { text: s.name },
        categories: ref(s.categories),
        values: ref(s.values),
      })),
    );
    const axis = (a: { title?: string; min?: number; max?: number } | undefined) => ({ title: a?.title === undefined ? null : { text: a.title }, min: a?.min ?? null, max: a?.max ?? null });
    expect(c?.axes).toEqual(input.type === 'pie' ? { x: null, y: null } : { x: axis(input.axes?.x), y: axis(input.axes?.y) });
  });

  it('reads the anchor as for pictures', () => {
    const wb = createWorkbook();
    dataSheet(wb.addSheet('Data')).addChart({ type: 'pie', anchor: { from: 'E2', size: { width: 100, height: 50 } }, series: [SALES] });
    const c = readWorkbook(writeWorkbook(wb)).getSheet('Data').charts[0]!;
    expect(c.anchor).toEqual({ type: 'oneCell', from: { cell: 'E2', col: 4, row: 1, colOffset: 0, rowOffset: 0 }, size: { width: 952500, height: 476250 } });
    expect(c.name).toBe('Chart 1');
    expect(c.description).toBeNull();
  });

  it('reads combination charts, other types, reference titles, rich titles and order-independent children (CH-09, CH-22, CH-29)', () => {
    const wb = createWorkbook();
    dataSheet(wb.addSheet('Data')).addChart({ type: 'column', anchor: 'E2:L18', series: [SALES] });
    const bytes = patch(writeWorkbook(wb), (name, text) => {
      if (name !== 'xl/charts/chart1.xml') return text;
      return text
        .replace(
          '</c:barChart>',
          '</c:barChart><c:lineChart><c:grouping val="standard"/><c:ser><c:idx val="1"/><c:order val="1"/><c:tx><c:v>Trend_x0021_</c:v></c:tx><c:val><c:numLit><c:ptCount val="1"/></c:numLit></c:val></c:ser><c:axId val="500000001"/><c:axId val="500000002"/></c:lineChart><c:radarChart><c:radarStyle val="marker"/><c:ser><c:idx val="2"/><c:order val="2"/><c:cat><c:multiLvlStrRef><c:f>Data!$A$2:$B$4</c:f></c:multiLvlStrRef></c:cat></c:ser><c:axId val="1"/><c:axId val="2"/></c:radarChart>',
        )
        .replace('<c:autoTitleDeleted val="1"/>', '<c:autoTitleDeleted val="0"/><c:title><c:tx><c:strRef><c:f>Data!$B$1</c:f></c:strRef></c:tx></c:title>')
        .replace('<c:legendPos val="r"/>', '<c:legendPos val="tr"/>');
    });
    const c = readWorkbook(bytes).getSheet('Data').charts[0]!;
    expect(c.types).toEqual(['column', 'line', 'other:radarChart']);
    expect(c.title).toEqual({ ref: 'Data!$B$1' });
    expect(c.legend).toBe('topRight');
    expect(c.stacked).toBe(false);
    expect(c.series.map((s) => [s.type, s.name, s.categories, s.values])).toEqual([
      ['column', { ref: 'Data!$B$1' }, 'Data!$A$2:$A$4', 'Data!$B$2:$B$4'],
      ['line', { text: 'Trend!' }, null, null],
      ['other:radarChart', null, 'Data!$A$2:$B$4', null],
    ]);
    const rich = patch(writeWorkbook(wb), (name, text) =>
      name === 'xl/charts/chart1.xml'
        ? text.replace('<c:autoTitleDeleted val="1"/>', '<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Two</a:t></a:r><a:r><a:t> runs</a:t></a:r><a:br/><a:fld id="x" type="y"><a:t>field</a:t></a:fld></a:p><a:p><a:r><a:t>second</a:t></a:r></a:p></c:rich></c:tx></c:title>').replace('<c:legend><c:legendPos val="r"/>', '<c:legend>')
        : text,
    );
    const r = readWorkbook(rich).getSheet('Data').charts[0]!;
    expect(r.title).toEqual({ text: 'Two runs\nfield\nsecond' });
    expect(r.legend).toBe('right');
    // A title without text (automatic title) and an absent legend.
    const auto = patch(writeWorkbook(wb), (name, text) => (name === 'xl/charts/chart1.xml' ? text.replace('<c:autoTitleDeleted val="1"/>', '<c:title><c:overlay val="0"/></c:title>').replace(/<c:legend>.*<\/c:legend>/, '') : text));
    const a = readWorkbook(auto).getSheet('Data').charts[0]!;
    expect([a.title, a.legend]).toEqual([null, null]);
  });

  it('reports shapes but not charts with XLV_W_DRAWING_CONTENT_NOT_READ', () => {
    const wb = createWorkbook();
    dataSheet(wb.addSheet('Data')).addChart({ type: 'column', anchor: 'E2:L18', series: [SALES] });
    const bytes = patch(writeWorkbook(wb), (name, text) =>
      name === 'xl/drawings/drawing1.xml' ? text.replace('</xdr:wsDr>', '<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp/><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>') : text,
    );
    const back = readWorkbook(bytes);
    expect(back.getSheet('Data').charts).toHaveLength(1);
    expect(back.warnings.map((w) => [w.code, w.message])).toEqual([['XLV_W_DRAWING_CONTENT_NOT_READ', expect.stringContaining('1 object other than pictures and charts')]]);
    expect(readWorkbook(writeWorkbook(wb)).warnings).toEqual([]);
  });

  it('skips broken chart frames with a warning', () => {
    const wb = createWorkbook();
    dataSheet(wb.addSheet('Data')).addChart({ type: 'column', anchor: 'E2:L18', series: [SALES] });
    const noRel = patch(writeWorkbook(wb), (name, text) => (name === 'xl/drawings/drawing1.xml' ? text.replace('r:id="rId1"', 'r:id="rId9"') : text));
    expect(readWorkbook(noRel).getSheet('Data').charts).toEqual([]);
    expect(readWorkbook(noRel).warnings.map((w) => w.code)).toEqual(['XLV_W_INVALID_VALUE']);
    const noPart = patch(writeWorkbook(wb), (name, text) => (name === 'xl/charts/chart1.xml' ? null : text));
    expect(readWorkbook(noPart).warnings.map((w) => w.code)).toEqual(['XLV_W_PART_MISSING']);
    const badAnchor = patch(writeWorkbook(wb), (name, text) => (name === 'xl/drawings/drawing1.xml' ? text.replace('<xdr:col>4</xdr:col>', '<xdr:col>x</xdr:col>') : text));
    expect(readWorkbook(badAnchor).warnings.map((w) => w.message)).toEqual(['A chart of the sheet has an invalid position; it is skipped.']);
  });

  it('inspect reports chartCount and samples; editing keeps charts and row insertion refuses them', () => {
    const wb = createWorkbook();
    const s = dataSheet(wb.addSheet('Data'));
    for (let i = 0; i < 12; i++) s.addChart({ type: i % 2 ? 'line' : 'column', anchor: `E${String(i * 20 + 2)}:L${String(i * 20 + 18)}`, ...(i === 0 ? { title: 'Sales' } : {}), series: [SALES] });
    const bytes = writeWorkbook(wb);
    const sheet = inspect(bytes).sheets[0]!;
    expect(sheet.chartCount).toBe(12);
    expect(sheet.charts).toHaveLength(10);
    expect(sheet.charts.slice(0, 2)).toEqual([
      { cell: 'E2', types: ['column'], title: { text: 'Sales' }, seriesCount: 1 },
      { cell: 'E22', types: ['line'], title: null, seriesCount: 1 },
    ]);
    const edited = editCells(bytes, [{ sheet: 'Data', cell: 'A2', value: 'January' }]).bytes;
    expect(parts(edited)['xl/charts/chart1.xml']).toBe(parts(bytes)['xl/charts/chart1.xml']);
    expect(readWorkbook(edited).getSheet('Data').charts).toHaveLength(12);
    expectXlvError(() => insertRows(bytes, { sheet: 'Data', at: 2, count: 1 }), 'XLV_E_NOT_SUPPORTED');
  });
});
