// insertRows and deleteRows cannot rewrite the references of chart parts yet, so they refuse a
// workbook whose chart parts may refer to the sheet. Chart parts were found only by the name
// "/charts/*.xml" (1.3.0 penetration test): a chart part with another name, folder or extension
// let the rows move while the chart's references stayed where they were. Chart parts are now
// found by their content type (docs/facts/charts.md CH-01) and by the chart relationships of the
// package.
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { deleteRows, insertRows } from '../../packages/xlvane/src/rows';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError } from '../helpers/expect-error';

function original(): Uint8Array {
  const wb = createWorkbook();
  wb.addSheet('Data').set('B2', 1).set('B3', 2);
  wb.addSheet('Report').addChart({ type: 'column', anchor: 'D2:J12', series: [{ values: 'Data!B2:B3' }] });
  return writeWorkbook(wb);
}

/**
 * The workbook with its chart part stored as `zipName`, referred to from the drawing as `target`;
 * `types` edits [Content_Types].xml.
 */
function renamed(zipName: string, target: string, types: (xml: string) => string = (x) => x.replace('/xl/charts/chart1.xml', `/${zipName}`)): Uint8Array {
  const out: Record<string, Uint8Array> = {};
  for (const [name, data] of Object.entries(unzipSync(original()))) {
    if (name === 'xl/charts/chart1.xml') out[zipName] = data;
    else if (name === '[Content_Types].xml') out[name] = strToU8(types(strFromU8(data)));
    else if (name === 'xl/drawings/_rels/drawing1.xml.rels') out[name] = strToU8(strFromU8(data).replace('../charts/chart1.xml', target));
    else out[name] = data;
  }
  return zipSync(out);
}

const variants: readonly (readonly [string, Uint8Array])[] = [
  ['another folder (xl/graphs/chart1.xml)', renamed('xl/graphs/chart1.xml', '../graphs/chart1.xml')],
  ['a subfolder (xl/charts/sub/chart1.xml)', renamed('xl/charts/sub/chart1.xml', '../charts/sub/chart1.xml')],
  ['another extension with an Override (xl/charts/chart1.bin)', renamed('xl/charts/chart1.bin', '../charts/chart1.bin')],
  [
    'another extension with a Default content type (xl/charts/chart1.cht)',
    renamed('xl/charts/chart1.cht', '../charts/chart1.cht', (x) =>
      x.replace(/<Override PartName="\/xl\/charts\/chart1\.xml"[^>]*\/>/, '').replace('<Default Extension="xml"', '<Default Extension="cht" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/><Default Extension="xml"'),
    ),
  ],
  ['a percent-encoded target (../charts/chart%31.xml)', renamed('xl/charts/chart1.xml', '../charts/chart%31.xml', (x) => x)],
];

describe('row operations and chart parts with other names', () => {
  for (const [label, bytes] of variants) {
    it(`refuses insertRows and deleteRows for a chart part in ${label}`, () => {
      // The chart is read from the renamed part.
      expect(readWorkbook(bytes).getSheet('Report').charts[0]?.series[0]?.values).toBe('Data!$B$2:$B$3');
      const error = expectXlvError(() => insertRows(bytes, { sheet: 'Data', at: 1, count: 5 }), 'XLV_E_NOT_SUPPORTED');
      expect(error.message).toContain('chart');
      expectXlvError(() => deleteRows(bytes, { sheet: 'Data', at: 1, count: 1 }), 'XLV_E_NOT_SUPPORTED');
    });
  }

  it('refuses them for the target of a chart relationship whose part has another content type', () => {
    // xlvane skips such a chart when reading (CH-01), but another application may still show it.
    const bytes = renamed('xl/graphs/chart1.xml', '../graphs/chart1.xml', (x) => x.replace(/<Override PartName="\/xl\/charts\/chart1\.xml"[^>]*\/>/, ''));
    expect(readWorkbook(bytes).getSheet('Report').charts).toEqual([]);
    expectXlvError(() => insertRows(bytes, { sheet: 'Data', at: 1, count: 5 }), 'XLV_E_NOT_SUPPORTED');
  });

  it('still allows row operations on a sheet no chart refers to', () => {
    const wb = createWorkbook();
    wb.addSheet('Data').set('B2', 1).set('B3', 2);
    wb.addSheet('Other').set('A1', 1);
    wb.addSheet('Report').addChart({ type: 'column', anchor: 'D2:J12', series: [{ values: 'Data!B2:B3' }] });
    const bytes = writeWorkbook(wb);
    const result = insertRows(bytes, { sheet: 'Other', at: 1, count: 1 });
    expect(readWorkbook(result.bytes).getSheet('Other').cell('A2')?.value).toBe(1);
  });
});
