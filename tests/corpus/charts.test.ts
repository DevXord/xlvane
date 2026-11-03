// Charts of the Apache POI corpus (tests/corpus/poi, Apache-2.0) as read by xlvane (D-K01). The
// expected values were taken from the chart parts of each file (docs/facts/charts.md S-POI,
// CH-17, CH-18, CH-22, CH-29); charts on chart sheets are not read (CH-05, maintainer's decision).
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';

const read = (name: string) => readWorkbook(readFileSync(new URL(`./poi/${name}`, import.meta.url)));
const brief = (name: string) =>
  read(name).sheets.flatMap((s) =>
    s.charts.map((c) => ({ sheet: s.name, at: c.anchor.type === 'absolute' ? null : c.anchor.from.cell, types: c.types, title: c.title, legend: c.legend, series: c.series.map((x) => [x.type, x.name, x.categories, x.values]) })),
  );

describe('charts of the POI corpus', () => {
  it('WithChart.xlsx and WithTwoCharts.xlsx (Excel 12): line with literal series names, area', () => {
    const line = { sheet: 'Sheet2', at: 'A1', types: ['line'], title: null, legend: 'right', series: [['line', { text: '1st Column' }, null, 'Sheet1!$A$1:$A$6'], ['line', { text: '2nd Column' }, null, 'Sheet1!$B$1:$B$6']] };
    expect(brief('WithChart.xlsx')).toEqual([line]);
    expect(brief('WithTwoCharts.xlsx')).toEqual([line, { sheet: 'Sheet3', at: 'A1', types: ['area'], title: null, legend: 'right', series: [['area', { text: 'Squares' }, null, 'Sheet1!$E$1:$E$6'], ['area', { text: 'Base Numbers' }, null, 'Sheet1!$D$1:$D$6']] }]);
    expect(read('WithChart.xlsx').warnings).toEqual([]);
  });

  it('WithThreeCharts.xlsx: a title in two runs, a pie without axes, an area with a title', () => {
    const charts = brief('WithThreeCharts.xlsx');
    expect(charts.map((c) => [c.sheet, c.types, c.title])).toEqual([
      ['Sheet2', ['line'], null],
      ['Sheet2', ['pie'], { text: 'Pie Chart Title Thingy' }],
      ['Sheet3', ['area'], { text: 'Sheet 3 Chart with Title' }],
    ]);
    expect(charts[1]?.series).toEqual([['pie', null, null, 'Sheet1!$B$1:$B$6']]);
    expect(read('WithThreeCharts.xlsx').getSheet('Sheet2').charts[1]?.axes).toEqual({ x: null, y: null });
  });

  it('chart_sheet.xlsx and WithChartSheet.xlsx: charts on chart sheets are not read', () => {
    for (const name of ['chart_sheet.xlsx', 'WithChartSheet.xlsx']) {
      const wb = read(name);
      expect(wb.sheets.some((s) => s.kind === 'chartsheet')).toBe(true);
      expect(wb.sheets.flatMap((s) => s.charts)).toEqual([]);
    }
  });

  it('SimpleScatterChart.xlsx (Excel 15): the embedded scatter chart, reference name, legend at the bottom', () => {
    expect(brief('SimpleScatterChart.xlsx')).toEqual([{ sheet: 'Sheet1', at: 'A4', types: ['scatter'], title: null, legend: 'bottom', series: [['scatter', { ref: 'Sheet1!$B$1' }, 'Sheet1!$A$2:$A$3', 'Sheet1!$B$2:$B$3']] }]);
  });

  it('chartTitle_*.xlsx: no title, a literal title, a reference title (CH-22)', () => {
    const series = [['scatter', null, 'Sheet1!$A$1:$A$4', 'Sheet1!$B$1:$B$4']];
    expect(brief('chartTitle_noTitle.xlsx')).toEqual([{ sheet: 'Sheet1', at: 'C5', types: ['scatter'], title: null, legend: null, series }]);
    expect(brief('chartTitle_withTitle.xlsx')[0]).toMatchObject({ title: { text: 'Original title set in Excel 2013' }, series });
    expect(brief('chartTitle_withTitleFormula.xlsx')[0]).toMatchObject({ title: { ref: 'Sheet1!$E$1' }, series });
  });

  it('123233_charts.xlsx (not written by Excel, out of schema order, CH-29): two columns, a pie and a line', () => {
    const charts = brief('123233_charts.xlsx');
    expect(charts.map((c) => [c.at, c.types, c.title, c.series.length])).toEqual([
      ['A2', ['column'], null, 5],
      ['I2', ['pie'], { text: 'Revenue' }, 1],
      ['A24', ['column'], null, 5],
      ['I24', ['line'], null, 5],
    ]);
    expect(charts[1]?.series).toEqual([['pie', { ref: 'data_Page1_1_2!B1' }, 'data_Page1_1_2!A2:A6', 'data_Page1_1_2!B2:B6']]);
    const line = read('123233_charts.xlsx').sheets[0]!.charts[3]!;
    expect(line.axes.x?.title).toEqual({ text: 'Year' });
    expect(line.axes.y?.title).toEqual({ text: 'Revenue' });
  });

  it('other files: combination, 3-D pie, axis scale', () => {
    expect(brief('57362.xlsx')[0]?.types).toEqual(['column', 'line']);
    expect(brief('45544.xlsx').map((c) => c.types)).toEqual([['other:pie3DChart'], ['other:pie3DChart']]);
    const cities = read('DataTableCities.xlsx').getSheet('Formula').charts[0]!;
    expect(cities.axes).toEqual({ x: { title: { text: 'Longitude' }, min: -180, max: 180 }, y: { title: { text: 'Latitude' }, min: -90, max: 90 } });
  });
});
