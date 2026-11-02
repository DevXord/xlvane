// insertRows and deleteRows find the sheet's drawing and legacyDrawing elements with the XML
// scanner, not with a regular expression (1.1.1 security review, D-H01): a comment holding many
// "<drawing " took quadratic time, and an element inside a comment was taken as the real one.
import { describe, expect, it } from 'vitest';
import { deleteRows, insertRows } from '../../packages/xlvane/src/rows';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { buildXlsx } from '../helpers/xlsx-builder';

const NSR = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const rels = (items: [string, string, string][]) =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items
    .map(([id, t, tgt]) => `<Relationship Id="${id}" Type="${NSR}/${t}" Target="${tgt}"/>`)
    .join('')}</Relationships>`;
const one = '<row r="1"><c r="A1"><v>1</v></c></row>';

describe('drawing elements of the changed sheet', () => {
  it('a comment with 60 000 "<drawing " and "<legacyDrawing " is skipped in linear time', () => {
    let comment = '<!--';
    for (let i = 0; i < 60000; i++) comment += `<drawing ${String(i)} <legacyDrawing `;
    comment += '-->';
    const bytes = buildXlsx({ sheets: [{ name: 'S', data: one, after: comment }] });
    for (const run of [() => insertRows(bytes, { sheet: 'S', at: 1, count: 1 }), () => deleteRows(bytes, { sheet: 'S', at: 2, count: 1 })]) {
      const start = performance.now();
      run();
      expect(performance.now() - start).toBeLessThan(3000);
    }
  });

  it('a drawing element inside a comment is not the sheet\'s drawing', () => {
    // The commented-out drawing holds a shape, which insertRows refuses to move; it must be ignored.
    const drawing = `<xdr:wsDr xmlns:xdr="${XDR}"><xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="1" cy="1"/><xdr:sp/><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>`;
    const bytes = buildXlsx({
      sheets: [{ name: 'S', data: one, after: '<!--<drawing r:id="rId1"/>-->' }],
      extraParts: { 'xl/worksheets/_rels/sheet1.xml.rels': rels([['rId1', 'drawing', '../drawings/drawing1.xml']]), 'xl/drawings/drawing1.xml': drawing },
    });
    const out = insertRows(bytes, { sheet: 'S', at: 1, count: 1 });
    expect(readWorkbook(out.bytes).sheets[0]?.cell('A2')?.value).toBe(1);
  });

  it('the real drawing element is still found, with any prefix bound to the relationships namespace', () => {
    const drawing = `<xdr:wsDr xmlns:xdr="${XDR}"><xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="1" cy="1"/><xdr:sp/><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>`;
    const bytes = buildXlsx({
      sheets: [{ name: 'S', data: one, after: `<drawing xmlns:q="${NSR}" q:id="rId1"/>` }],
      extraParts: { 'xl/worksheets/_rels/sheet1.xml.rels': rels([['rId1', 'drawing', '../drawings/drawing1.xml']]), 'xl/drawings/drawing1.xml': drawing },
    });
    expect(() => insertRows(bytes, { sheet: 'S', at: 1, count: 1 })).toThrow(/charts, shapes or other objects/);
  });

  it('a drawing element inside mc:AlternateContent is still found, as before 1.1.1 (independent review)', () => {
    const drawing = `<xdr:wsDr xmlns:xdr="${XDR}"><xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="1" cy="1"/><xdr:sp/><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>`;
    const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
    const bytes = buildXlsx({
      sheets: [{ name: 'S', data: one, after: `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice Requires="x14"><drawing r:id="rId1"/></mc:Choice><mc:Fallback><drawing r:id="rId1"/></mc:Fallback></mc:AlternateContent>` }],
      extraParts: { 'xl/worksheets/_rels/sheet1.xml.rels': rels([['rId1', 'drawing', '../drawings/drawing1.xml']]), 'xl/drawings/drawing1.xml': drawing },
    });
    expect(() => insertRows(bytes, { sheet: 'S', at: 1, count: 1 })).toThrow(/charts, shapes or other objects/);
  });
});
