// Extended document properties: docProps/app.xml (ECMA-376 Part 1 §15.2.12.3, §22.2, §18.17.6.5;
// Part 4 §13.2.12; [MS-OI29500] §2.1.1717, §2.1.1720, §2.1.1722; docs/facts/sml-write.md W2-52;
// docs/decisions.md D-A10).
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { editCells } from '../../packages/xlvane/src/edit';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { openWorkbookStream } from '../../packages/xlvane/src/stream-read';
import { writeWorkbookStream } from '../../packages/xlvane/src/stream-write';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError } from '../helpers/expect-error';
import { buildXlsx } from '../helpers/xlsx-builder';

const APP_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties';
const APP_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';
const VT = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';

function parts(bytes: Uint8Array): Record<string, string> {
  return Object.fromEntries(Object.entries(unzipSync(bytes)).map(([k, v]) => [k, strFromU8(v)]));
}

describe('writing docProps/app.xml', () => {
  it('is not written without extended properties or visible names', () => {
    const wb = createWorkbook({ properties: { title: 'T' } });
    wb.addSheet('S').autoFilter('A1:B2');
    const files = parts(writeWorkbook(wb));
    expect(files['docProps/app.xml']).toBeUndefined();
    expect(files['_rels/.rels']).not.toContain('extended-properties');
  });

  it('writes company and manager with sheet titles, its relationship and content type', () => {
    const wb = createWorkbook({ properties: { company: 'Acme & Co', manager: 'Ann', title: 'T' } });
    wb.addSheet('Data');
    wb.addSheet('Notes');
    const files = parts(writeWorkbook(wb));
    expect(files['docProps/app.xml']).toBe(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        `<Properties xmlns="${APP_NS}" xmlns:vt="${VT}">` +
        '<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant></vt:vector></HeadingPairs>' +
        '<TitlesOfParts><vt:vector size="2" baseType="lpstr"><vt:lpstr>Data</vt:lpstr><vt:lpstr>Notes</vt:lpstr></vt:vector></TitlesOfParts>' +
        '<Manager>Ann</Manager><Company>Acme &amp; Co</Company></Properties>',
    );
    expect(files['_rels/.rels']).toContain(`<Relationship Id="rId3" Type="${APP_REL}" Target="docProps/app.xml"/>`);
    expect(files['[Content_Types].xml']).toContain('<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>');
    expect(files['docProps/core.xml']).toContain('<dc:title>T</dc:title>');
  });

  it('lists visible defined names after the sheets (Part 1 §18.17.6.5)', () => {
    const wb = createWorkbook();
    wb.addSheet('Data');
    wb.addSheet("Bob's");
    wb.defineName('Total', 'Data!$B$10').defineName('Local', "'Bob''s'!$A$1", { sheet: "Bob's" }).defineName('Secret', 'Data!$A$1', { hidden: true });
    const app = parts(writeWorkbook(wb))['docProps/app.xml'] ?? '';
    expect(app).toContain('<vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant><vt:variant><vt:lpstr>Named Ranges</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant>');
    expect(app).toContain("<vt:vector size=\"4\" baseType=\"lpstr\"><vt:lpstr>Data</vt:lpstr><vt:lpstr>Bob's</vt:lpstr><vt:lpstr>Total</vt:lpstr><vt:lpstr>'Bob''s'!Local</vt:lpstr></vt:vector>");
    expect(app).not.toContain('Secret');
    expect(app).not.toContain('<Company>');
  });

  it('escapes text as ST_Xstring and checks AppVersion', () => {
    const wb = createWorkbook({ properties: { company: '_x0041_', application: 'Report tool', appVersion: '1.0' } });
    wb.addSheet('S');
    const bytes = writeWorkbook(wb);
    expect(parts(bytes)['docProps/app.xml']).toContain('<Company>_x005F_x0041_</Company>');
    expect(readWorkbook(bytes).properties).toEqual({ company: '_x0041_', application: 'Report tool', appVersion: '1.0' });
    for (const appVersion of ['1', '123.4', '1.12345', 'v1.0']) {
      expectXlvError(() => createWorkbook({ properties: { appVersion } }), 'XLV_E_INVALID_OPTION', { option: 'properties.appVersion' });
    }
    expectXlvError(() => createWorkbook({ properties: { company: 1 as never } }), 'XLV_E_INVALID_OPTION', { option: 'properties.company' });
  });

  it('the streaming writer writes the same part, also for sheets given as an iterable', async () => {
    const options = { properties: { manager: 'M' }, definedNames: [{ name: 'All', refersTo: 'S!$A:$A' }] };
    function* sheets() {
      yield { name: 'S', rows: [[1]] };
    }
    for (const stream of [writeWorkbookStream({ ...options, sheets: [{ name: 'S', rows: [[1]] }] }), writeWorkbookStream({ ...options, sheets: sheets() })]) {
      const bytes = new Uint8Array(await new Response(stream).arrayBuffer());
      const files = parts(bytes);
      expect(files['docProps/app.xml']).toContain('<vt:lpstr>S</vt:lpstr><vt:lpstr>All</vt:lpstr>');
      expect(files['[Content_Types].xml']).toContain('/docProps/app.xml');
      expect((await openWorkbookStream(bytes)).properties).toEqual({ manager: 'M' });
    }
  });
});

describe('reading docProps/app.xml', () => {
  const withApp = (xml: string, rel = APP_REL): Uint8Array => buildXlsx({ rootRels: [[rel, 'docProps/app.xml']], extraParts: { 'docProps/app.xml': xml } });

  it('reads Company, Manager, Application and AppVersion, Transitional and Strict', () => {
    const xml = `<Properties xmlns="${APP_NS}" xmlns:vt="${VT}"><Application>Microsoft Excel</Application><Company>A_x000D_B</Company><Manager/><AppVersion>16.0300</AppVersion><TotalTime>3</TotalTime></Properties>`;
    expect(readWorkbook(withApp(xml)).properties).toEqual({ application: 'Microsoft Excel', company: 'A\rB', manager: '', appVersion: '16.0300' });
    const strict = '<Properties xmlns="http://purl.oclc.org/ooxml/officeDocument/extendedProperties"><Company>C</Company></Properties>';
    expect(readWorkbook(withApp(strict, 'http://purl.oclc.org/ooxml/officeDocument/relationships/extendedProperties')).properties).toEqual({ company: 'C' });
  });

  it('warns about a bad part and reads on', () => {
    const malformed = readWorkbook(withApp('<Properties'));
    expect(malformed.properties).toEqual({});
    expect(malformed.warnings.map((w) => w.code)).toContain('XLV_W_INVALID_VALUE');
    const foreign = readWorkbook(withApp('<Other xmlns="urn:x"><Company>C</Company></Other>'));
    expect(foreign.properties).toEqual({});
    const missing = readWorkbook(buildXlsx({ rootRels: [[APP_REL, 'docProps/app.xml']] }));
    expect(missing.warnings.map((w) => w.code)).toContain('XLV_W_PART_MISSING');
    expectXlvError(() => readWorkbook(withApp('<!DOCTYPE x><Properties/>')), 'XLV_E_XML_DTD_FORBIDDEN');
  });

  it('editCells keeps the part byte for byte', () => {
    const wb = createWorkbook({ properties: { company: 'Acme' } });
    wb.addSheet('S').set('A1', 1);
    const bytes = writeWorkbook(wb);
    const edited = editCells(bytes, [{ sheet: 'S', cell: 'A2', value: 2 }]);
    expect(parts(edited.bytes)['docProps/app.xml']).toBe(parts(bytes)['docProps/app.xml']);
  });
});
