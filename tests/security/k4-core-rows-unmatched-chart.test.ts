// insertRows and deleteRows after the K3 review (B7): a chart part stored as the ZIP item
// "xl/charts/chart%31.xml" (with the Override "/xl/charts/chart%31.xml" and the drawing target
// "../charts/chart%31.xml") is shown by LibreOffice 24.2. For xlvane the item is not a part: its
// logical name holds a percent-encoded unreserved character, which a part name cannot contain
// (docs/facts/opc.md F-OPC-01 b, F-OPC-12, F-OPC-13), while the drawing's target resolves to
// "/xl/charts/chart1.xml" (D-OPC-02), which does not exist. readWorkbook reports
// XLV_W_PART_MISSING and no chart, and insertRows moved the rows while the chart kept its
// references. A chart (or drawing) relationship whose target is not a part of the package now
// makes insertRows and deleteRows refuse, since the chart it names may refer to the sheet.
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { deleteRows, insertRows } from '../../packages/xlvane/src/rows';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError } from '../helpers/expect-error';

function withChart(rename: (name: string) => string, edit: (name: string, xml: string) => string): Uint8Array {
  const wb = createWorkbook();
  wb.addSheet('Data').set('B2', 1).set('B3', 2);
  wb.addSheet('Report').addChart({ type: 'column', anchor: 'D2:J12', series: [{ values: 'Data!B2:B3' }] });
  const out: Record<string, Uint8Array> = {};
  for (const [name, data] of Object.entries(unzipSync(writeWorkbook(wb)))) {
    out[rename(name)] = name.endsWith('.xml') || name.endsWith('.rels') ? strToU8(edit(name, strFromU8(data))) : data;
  }
  return zipSync(out);
}

const percentEncoded = withChart(
  (name) => (name === 'xl/charts/chart1.xml' ? 'xl/charts/chart%31.xml' : name),
  (name, xml) => (name === '[Content_Types].xml' ? xml.replace('/xl/charts/chart1.xml', '/xl/charts/chart%31.xml') : name === 'xl/drawings/_rels/drawing1.xml.rels' ? xml.replace('../charts/chart1.xml', '../charts/chart%31.xml') : xml),
);

describe('B7: a chart relationship whose target is not a part', () => {
  it('the reproduction: a percent-encoded chart part name', () => {
    expect(readWorkbook(percentEncoded).warnings.map((w) => w.code)).toContain('XLV_W_PART_MISSING');
    const error = expectXlvError(() => insertRows(percentEncoded, { sheet: 'Data', at: 1, count: 2 }), 'XLV_E_NOT_SUPPORTED');
    expect(error.message).toContain('/xl/charts/chart1.xml');
    expectXlvError(() => deleteRows(percentEncoded, { sheet: 'Data', at: 1 }), 'XLV_E_NOT_SUPPORTED');
  });

  it('a chart relationship to a part that is missing', () => {
    // Stored under another name without the chart content type, so only the relationship leads to it.
    const missing = withChart((name) => (name === 'xl/charts/chart1.xml' ? 'xl/graphs/other.xml' : name), (_name, xml) => xml);
    expectXlvError(() => insertRows(missing, { sheet: 'Data', at: 1 }), 'XLV_E_NOT_SUPPORTED');
  });

  it('a drawing relationship of another sheet to a part that is missing', () => {
    const renames: Record<string, string> = { 'xl/drawings/drawing1.xml': 'xl/drawings/drawing%31.xml', 'xl/drawings/_rels/drawing1.xml.rels': 'xl/drawings/_rels/drawing%31.xml.rels', 'xl/charts/chart1.xml': 'xl/graphs/c.xml' };
    const missing = withChart((name) => renames[name] ?? name, (name, xml) => (name === 'xl/drawings/_rels/drawing1.xml.rels' ? xml.replace('../charts/chart1.xml', '../graphs/c.xml') : xml));
    expectXlvError(() => insertRows(missing, { sheet: 'Data', at: 1 }), 'XLV_E_NOT_SUPPORTED');
  });

  it('a workbook without charts still moves its rows', () => {
    const wb = createWorkbook();
    wb.addSheet('Data').set('B2', 1).set('B3', 2);
    wb.addSheet('Other').set('A1', { formula: 'SUM(Data!B2:B3)' });
    const out = readWorkbook(insertRows(writeWorkbook(wb), { sheet: 'Data', at: 1, count: 2 }).bytes);
    expect(out.getSheet('Other').cell('A1')?.formula?.text).toBe('SUM(Data!B4:B5)');
  });
});
