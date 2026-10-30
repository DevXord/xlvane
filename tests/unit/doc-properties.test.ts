// Document properties: docProps/core.xml (ECMA-376 Part 2 Clause 8; docs/facts/opc.md
// F-OPC-16 … F-OPC-20; docs/decisions.md D-A02).
import { unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { parseW3cdtf } from '../../packages/xlvane/src/doc-properties';
import { editCells } from '../../packages/xlvane/src/edit';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { openWorkbookStream } from '../../packages/xlvane/src/stream-read';
import { writeWorkbookStream } from '../../packages/xlvane/src/stream-write';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError } from '../helpers/expect-error';
import { buildXlsx } from '../helpers/xlsx-builder';

const dec = new TextDecoder();
const CORE_REL = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
const CP = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';

function parts(bytes: Uint8Array): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [name, data] of Object.entries(unzipSync(bytes))) out[name] = dec.decode(data);
  return out;
}

function withCore(xml: string, rel = CORE_REL): Uint8Array {
  return buildXlsx({ rootRels: [[rel, 'docProps/core.xml']], extraParts: { 'docProps/core.xml': xml } });
}

const created = new Date(Date.UTC(2026, 9, 30, 12, 34, 56, 789));

describe('writing document properties', () => {
  it('writes no core properties part when none is given', () => {
    const wb = createWorkbook();
    wb.addSheet('S');
    const files = parts(writeWorkbook(wb));
    expect(files['docProps/core.xml']).toBeUndefined();
    expect(files['_rels/.rels']).not.toContain('core-properties');
  });

  it('writes the part, its relationship and content type', () => {
    const wb = createWorkbook({ properties: { title: 'Q3 & <sales>', creator: 'Finance', created, lastPrinted: created, keywords: 'a b' } });
    wb.addSheet('S');
    const files = parts(writeWorkbook(wb));
    const core = files['docProps/core.xml'] ?? '';
    expect(core).toContain(`<cp:coreProperties xmlns:cp="${CP}" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"`);
    expect(core).toContain('<dc:title>Q3 &amp; &lt;sales&gt;</dc:title>');
    expect(core).toContain('<dcterms:created xsi:type="dcterms:W3CDTF">2026-10-30T12:34:56Z</dcterms:created>');
    expect(core).toContain('<cp:lastPrinted>2026-10-30T12:34:56Z</cp:lastPrinted>');
    expect(files['_rels/.rels']).toContain(`<Relationship Id="rId2" Type="${CORE_REL}" Target="docProps/core.xml"/>`);
    expect(files['[Content_Types].xml']).toContain('<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>');
  });

  it('is deterministic', () => {
    const wb = createWorkbook({ properties: { title: 'T', created } });
    wb.addSheet('S');
    expect(writeWorkbook(wb)).toEqual(writeWorkbook(wb));
  });

  it('round-trips every property (dates to whole seconds)', () => {
    const properties = {
      title: 'Title',
      subject: 'Subject',
      creator: 'Ann',
      keywords: 'k1, k2',
      description: 'Line 1\nLine 2',
      lastModifiedBy: 'Ben',
      category: 'Report',
      contentStatus: 'Final',
      identifier: 'ID-1',
      language: 'en-US',
      revision: '3',
      version: '1.0',
      created,
      modified: new Date(Date.UTC(2026, 9, 31)),
      lastPrinted: new Date(Date.UTC(2000, 0, 1, 23, 59, 59)),
    };
    const wb = createWorkbook({ properties });
    wb.addSheet('S');
    const read = readWorkbook(writeWorkbook(wb));
    expect(read.properties).toEqual({ ...properties, created: new Date(Date.UTC(2026, 9, 30, 12, 34, 56)) });
    expect(read.warnings).toEqual([]);
  });

  it('streaming writer writes the same part', async () => {
    const bytes = new Uint8Array(await new Response(writeWorkbookStream({ properties: { title: 'Streamed', created }, sheets: [{ name: 'S', rows: [[1]] }] })).arrayBuffer());
    expect(readWorkbook(bytes).properties).toEqual({ title: 'Streamed', created: new Date(Date.UTC(2026, 9, 30, 12, 34, 56)) });
    const dynamic = new Uint8Array(
      await new Response(
        writeWorkbookStream({
          properties: { creator: 'Gen' },
          sheets: (function* () {
            yield { name: 'S', rows: [[1]] };
          })(),
        }),
      ).arrayBuffer(),
    );
    expect(readWorkbook(dynamic).properties).toEqual({ creator: 'Gen' });
    expect((await openWorkbookStream(dynamic)).properties).toEqual({ creator: 'Gen' });
  });

  it('rejects invalid properties', () => {
    expectXlvError(() => createWorkbook({ properties: 'x' as never }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => createWorkbook({ properties: { titel: 'x' } as never }), 'XLV_E_INVALID_OPTION', { option: 'properties.titel' });
    expectXlvError(() => createWorkbook({ properties: { company: 'x' } as never }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => createWorkbook({ properties: { title: 1 } as never }), 'XLV_E_INVALID_OPTION', { option: 'properties.title' });
    expectXlvError(() => createWorkbook({ properties: { title: 'a\u0001b' } }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => createWorkbook({ properties: { title: '\uD800' } }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => createWorkbook({ properties: { created: new Date(NaN) } }), 'XLV_E_INVALID_OPTION', { option: 'properties.created' });
    expectXlvError(() => createWorkbook({ properties: { created: '2026-01-01' as never } }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => createWorkbook({ properties: { modified: new Date(Date.UTC(10000, 0, 1)) } }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => createWorkbook({ properties: JSON.parse('{"__proto__":{"x":1}}') as never }), 'XLV_E_INVALID_OPTION');
  });

  it('copies the dates it was given', () => {
    const d = new Date(Date.UTC(2020, 0, 1));
    const wb = createWorkbook({ properties: { created: d } });
    d.setUTCFullYear(1999);
    expect(wb.properties.created?.getUTCFullYear()).toBe(2020);
  });
});

describe('reading document properties', () => {
  it('reads core.xml through the package relationship', () => {
    const xml = `<?xml version="1.0"?><cp:coreProperties xmlns:cp="${CP}" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
      <dc:creator>Alan</dc:creator><dc:title>OPC</dc:title><cp:unknown>x</cp:unknown>
      <dcterms:created xsi:type="dcterms:W3CDTF">2005-06-12</dcterms:created>
      <dcterms:modified xsi:type="dcterms:W3CDTF">2005-11-23T10:20:30+02:00</dcterms:modified>
      <cp:lastPrinted>2017-04-17T14:20:10.5</cp:lastPrinted>
      <cp:keywords xml:lang="en-US"> color <value xml:lang="en-CA">colour</value><value xml:lang="fr-FR">couleur</value></cp:keywords>
    </cp:coreProperties>`;
    const wb = readWorkbook(withCore(xml));
    expect(wb.properties).toEqual({
      creator: 'Alan',
      title: 'OPC',
      created: new Date(Date.UTC(2005, 5, 12)),
      modified: new Date(Date.UTC(2005, 10, 23, 8, 20, 30)),
      lastPrinted: new Date(Date.UTC(2017, 3, 17, 14, 20, 10, 500)),
      keywords: 'color colour couleur',
    });
    expect(wb.warnings).toEqual([]);
  });

  it('accepts the relationship type of ECMA-376 Part 4 §13.2.10', () => {
    const xml = `<coreProperties xmlns="${CP}" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>T</dc:title></coreProperties>`;
    expect(readWorkbook(withCore(xml, 'http://schemas.openxmlformats.org/officedocument/2006/relationships/metadata/core-properties')).properties).toEqual({ title: 'T' });
  });

  it('ignores invalid dates, a foreign root, malformed XML and a missing part with warnings', () => {
    const bad = readWorkbook(withCore(`<coreProperties xmlns="${CP}" xmlns:dcterms="http://purl.org/dc/terms/"><dcterms:created>yesterday</dcterms:created><dcterms:modified>2026-02-30</dcterms:modified></coreProperties>`));
    expect(bad.properties).toEqual({});
    expect(bad.warnings.map((w) => [w.code, w.count, w.location.part])).toEqual([['XLV_W_INVALID_VALUE', 2, '/docProps/core.xml']]);
    expect(readWorkbook(withCore('<other/>')).warnings.map((w) => w.code)).toEqual(['XLV_W_INVALID_VALUE']);
    expect(readWorkbook(withCore('<coreProperties')).warnings.map((w) => w.code)).toEqual(['XLV_W_INVALID_VALUE']);
    const missing = buildXlsx({ rootRels: [[CORE_REL, 'docProps/core.xml']] });
    expect(readWorkbook(missing).warnings.map((w) => w.code)).toEqual(['XLV_W_PART_MISSING']);
    expect(readWorkbook(buildXlsx()).properties).toEqual({});
  });

  it('rejects a DTD in the part like any other part', () => {
    expectXlvError(() => readWorkbook(withCore(`<!DOCTYPE x [<!ENTITY a "b">]><coreProperties xmlns="${CP}"/>`)), 'XLV_E_XML_DTD_FORBIDDEN');
  });

  it('parses W3CDTF values', () => {
    expect(parseW3cdtf('2026')).toEqual(new Date(Date.UTC(2026, 0, 1)));
    expect(parseW3cdtf('2026-10')).toEqual(new Date(Date.UTC(2026, 9, 1)));
    expect(parseW3cdtf('2026-10-30T01:02Z')).toEqual(new Date(Date.UTC(2026, 9, 30, 1, 2)));
    expect(parseW3cdtf(' 2026-10-30T01:02:03.25-01:30 ')).toEqual(new Date(Date.UTC(2026, 9, 30, 2, 32, 3, 250)));
    for (const bad of ['', '26-01-01', '2026-13-01', '2026-02-29', '2026-01-01T24:00:00Z', '2026-01-01T01:02:03+25:00', '2026-01-01 01:02']) expect(parseW3cdtf(bad)).toBeNull();
  });

  it('editCells keeps the part byte for byte', () => {
    const wb = createWorkbook({ properties: { title: 'Keep', created } });
    wb.addSheet('S').set('A1', 1);
    const bytes = writeWorkbook(wb);
    const edited = editCells(bytes, [{ sheet: 'S', cell: 'B1', value: 2 }]).bytes;
    expect(unzipSync(edited)['docProps/core.xml']).toEqual(unzipSync(bytes)['docProps/core.xml']);
    expect(readWorkbook(edited).properties.title).toBe('Keep');
  });
});
