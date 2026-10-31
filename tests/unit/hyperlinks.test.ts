// Hyperlinks (ECMA-376 Part 1 §18.3.1.47, §18.3.1.48, §15.3; Part 4 §13.3; [MS-OI29500]
// §2.1.626; docs/facts/sml-write.md W2-42 … W2-44; docs/decisions.md D-A08).
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { editCells } from '../../packages/xlvane/src/edit';
import { inspect } from '../../packages/xlvane/src/inspect';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { openWorkbookStream } from '../../packages/xlvane/src/stream-read';
import { writeWorkbookStream } from '../../packages/xlvane/src/stream-write';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError, expectXlvErrorAsync } from '../helpers/expect-error';
import { buildXlsx, NS } from '../helpers/xlsx-builder';

const REL = `${NS.transitional.rel}hyperlink`;
const rels = (items: string): string =>
  `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`;

async function bytesOf(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

function parts(bytes: Uint8Array): Record<string, string> {
  return Object.fromEntries(Object.entries(unzipSync(bytes)).map(([k, v]) => [k, strFromU8(v)]));
}

describe('reading hyperlinks', () => {
  const file = buildXlsx({
    sheets: [
      {
        name: 'Links',
        data: '<row r="1"><c r="A1" t="inlineStr"><is><t>web</t></is></c><c r="B1" t="inlineStr"><is><t>inside</t></is></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c></row>',
        after:
          '<hyperlinks>' +
          '<hyperlink ref="A1" r:id="rId1" tooltip="Open_x000A_it"/>' +
          '<hyperlink ref="B1" location="\'Other sheet\'!A1" display="Go"/>' +
          '<hyperlink ref="A2:B2" r:id="rId2" location="top"/>' +
          '<hyperlink ref="C9" r:id="rId9"/>' +
          '<hyperlink ref="nope" location="x"/>' +
          '</hyperlinks>',
      },
    ],
    extraParts: {
      'xl/worksheets/_rels/sheet1.xml.rels': rels(
        `<Relationship Id="rId1" Type="${REL}" Target="https://example.com/?a=1&amp;b=2" TargetMode="External"/>` +
          `<Relationship Id="rId2" Type="${REL}" Target="file:///C:/secret.txt" TargetMode="External"/>`,
      ),
    },
  });

  it('lists links with their target, location, tooltip and display text', () => {
    const wb = readWorkbook(file);
    const sheet = wb.getSheet('Links');
    expect(sheet.hyperlinks).toEqual([
      { ref: 'A1', target: 'https://example.com/?a=1&b=2', location: null, tooltip: 'Open\nit', display: null },
      { ref: 'B1', target: null, location: "'Other sheet'!A1", tooltip: null, display: 'Go' },
      { ref: 'A2:B2', target: 'file:///C:/secret.txt', location: 'top', tooltip: null, display: null },
      { ref: 'C9', target: null, location: null, tooltip: null, display: null },
    ]);
    expect(sheet.cell('A1')?.hyperlink?.target).toBe('https://example.com/?a=1&b=2');
    expect(sheet.cell('B1')?.hyperlink?.location).toBe("'Other sheet'!A1");
    // A range link is attached to every cell of the range that exists.
    expect(sheet.cell('A2')?.hyperlink?.ref).toBe('A2:B2');
    expect(sheet.cell('B2')?.hyperlink?.ref).toBe('A2:B2');
    expect(sheet.cell('A1')?.value).toBe('web');
  });

  it('warns about a missing relationship, an invalid range and unusual targets', () => {
    const warnings = readWorkbook(file).warnings;
    // Warnings with the same code and part are merged; the first message is kept (here: "nope").
    const invalid = warnings.find((w) => w.code === 'XLV_W_INVALID_VALUE');
    expect(invalid?.message).toContain('"nope"');
    expect(invalid?.count).toBe(2);
    expect(warnings.find((w) => w.code === 'XLV_W_UNUSUAL_HYPERLINK')?.message).toContain('file:///C:/secret.txt');
  });

  it('reads Strict files and sheets without links', () => {
    const strict = buildXlsx({
      strict: true,
      sheets: [{ name: 'S', data: '<row r="1"><c r="A1"><v>1</v></c></row>', after: '<hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks>' }],
      extraParts: {
        'xl/worksheets/_rels/sheet1.xml.rels': rels(`<Relationship Id="rId1" Type="${NS.strict.rel}hyperlink" Target="https://example.org" TargetMode="External"/>`),
      },
    });
    expect(readWorkbook(strict).getSheet('S').hyperlinks[0]?.target).toBe('https://example.org');
    expect(readWorkbook(buildXlsx()).sheets[0]?.hyperlinks).toEqual([]);
  });

  it('inspect reports the links', () => {
    const s = inspect(readWorkbook(file)).sheets[0];
    expect(s?.hyperlinkCount).toBe(4);
    expect(s?.hyperlinks[0]).toEqual({ ref: 'A1', target: 'https://example.com/?a=1&b=2', location: null });
  });

  it('editCells keeps the links and the relationships part unchanged', () => {
    const edited = editCells(file, [{ sheet: 'Links', cell: 'D1', value: 'x' }]);
    expect(parts(edited.bytes)['xl/worksheets/_rels/sheet1.xml.rels']).toBe(parts(file)['xl/worksheets/_rels/sheet1.xml.rels']);
    expect(readWorkbook(edited.bytes).getSheet('Links').hyperlinks).toEqual(readWorkbook(file).getSheet('Links').hyperlinks);
  });
});

describe('writing hyperlinks', () => {
  function book() {
    const wb = createWorkbook();
    wb.addSheet('Links')
      .set('A1', 'Example')
      .link('A1', { url: 'https://example.com/?q=a&b', tooltip: 'Open example.com' })
      .set('A2', 'Back')
      .link('A2', { location: "'Links'!A1", display: 'Back' })
      .link('B3:C4', { url: 'mailto:team@example.com', location: 'anchor' });
    wb.addSheet('Plain').set('A1', 1);
    return wb;
  }

  it('writes external links through relationships and internal ones in location only', () => {
    const p = parts(writeWorkbook(book()));
    const sheet = p['xl/worksheets/sheet1.xml']!;
    expect(sheet).toContain(
      '</sheetData><hyperlinks><hyperlink ref="A1" r:id="rId1" tooltip="Open example.com"/><hyperlink ref="A2" location="\'Links\'!A1" display="Back"/><hyperlink ref="B3:C4" r:id="rId2" location="anchor"/></hyperlinks></worksheet>',
    );
    expect(p['xl/worksheets/_rels/sheet1.xml.rels']).toBe(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="rId1" Type="${REL}" Target="https://example.com/?q=a&amp;b" TargetMode="External"/>` +
        `<Relationship Id="rId2" Type="${REL}" Target="mailto:team@example.com" TargetMode="External"/></Relationships>`,
    );
    // A sheet without relationships gets no relationships part; content types are unchanged.
    expect(p['xl/worksheets/_rels/sheet2.xml.rels']).toBeUndefined();
    expect(p['[Content_Types].xml']).not.toContain('hyperlink');
  });

  it('reads back what it wrote, in batch and streaming', async () => {
    const expected = [
      { ref: 'A1', target: 'https://example.com/?q=a&b', location: null, tooltip: 'Open example.com', display: null },
      { ref: 'A2', target: null, location: "'Links'!A1", tooltip: null, display: 'Back' },
      { ref: 'B3:C4', target: 'mailto:team@example.com', location: 'anchor', tooltip: null, display: null },
    ];
    expect(readWorkbook(writeWorkbook(book())).getSheet('Links').hyperlinks).toEqual(expected);
    const streamed = await bytesOf(
      writeWorkbookStream({
        sheets: [
          {
            name: 'Links',
            hyperlinks: [
              { ref: 'A1', url: 'https://example.com/?q=a&b', tooltip: 'Open example.com' },
              { ref: 'A2', location: "'Links'!A1", display: 'Back' },
              { ref: 'B3:C4', url: 'mailto:team@example.com', location: 'anchor' },
            ],
            rows: [['Example'], ['Back']],
          },
          { name: 'Plain', rows: [[1]] },
        ],
      }),
    );
    expect(readWorkbook(streamed).getSheet('Links').hyperlinks).toEqual(expected);
    expect(parts(streamed)['xl/worksheets/_rels/sheet2.xml.rels']).toBeUndefined();
    // Sheets given as an iterable also get their relationships part.
    async function* sheets() {
      yield { name: 'One', hyperlinks: [{ ref: 'A1', url: 'https://a.example' }], rows: [['a']] };
    }
    const dynamic = await bytesOf(writeWorkbookStream({ sheets: sheets() }));
    expect(readWorkbook(dynamic).getSheet('One').hyperlinks[0]?.target).toBe('https://a.example');
    // The streaming reader does not attach links to its cells (they follow the rows).
    const stream = await openWorkbookStream(streamed);
    for await (const row of stream.rows('Links')) for (const c of row.cells) expect(c.hyperlink).toBeUndefined();
  });

  it('replaces and removes links, and keeps ST_Xstring escapes', () => {
    const wb = createWorkbook();
    const s = wb.addSheet('S').link('A1', { url: 'https://old.example' }).link('a1', { location: 'S!B2', tooltip: 'tab\there _x0041_' }).link('B1', { url: 'https://x.example' }).link('B1', null);
    expect(readWorkbook(writeWorkbook(wb)).getSheet('S').hyperlinks).toEqual([{ ref: 'A1', target: null, location: 'S!B2', tooltip: 'tab\there _x0041_', display: null }]);
    expect(s.linkMap.size).toBe(1);
  });

  it('rejects invalid links with XLV_E_INVALID_HYPERLINK', async () => {
    const sheet = createWorkbook().addSheet('S');
    expectXlvError(() => sheet.link('A1', {}), 'XLV_E_INVALID_HYPERLINK', { cell: 'A1' });
    expectXlvError(() => sheet.link('A1', { url: '' }), 'XLV_E_INVALID_HYPERLINK', { property: 'url' });
    expectXlvError(() => sheet.link('A1', { url: 'https://a\nb' }), 'XLV_E_INVALID_HYPERLINK', { property: 'url' });
    expectXlvError(() => sheet.link('A1', { location: 'x'.repeat(2085) }), 'XLV_E_INVALID_HYPERLINK', { property: 'location' });
    expectXlvError(() => sheet.link('A1', { location: 'S!A1', display: 'x'.repeat(2085) }), 'XLV_E_INVALID_HYPERLINK', { property: 'display' });
    expectXlvError(() => sheet.link('A1', { url: 'https://a', tooltip: 'x'.repeat(256) }), 'XLV_E_INVALID_HYPERLINK', { property: 'tooltip' });
    expectXlvError(() => sheet.link('A1', { url: 1 } as never), 'XLV_E_INVALID_HYPERLINK', { property: 'url' });
    expectXlvError(() => sheet.link('A1', { href: 'x' } as never), 'XLV_E_INVALID_HYPERLINK', { property: 'href' });
    expectXlvError(() => sheet.link('A1', 'https://a' as never), 'XLV_E_INVALID_HYPERLINK');
    expectXlvError(() => sheet.link('A1', { url: '\uD800' }), 'XLV_E_INVALID_HYPERLINK');
    expectXlvError(() => sheet.link('ZZZZ1', { url: 'https://a' }), 'XLV_E_INVALID_CELL_REF');
    // The documented maxima are accepted.
    sheet.link('A1', { location: 'x'.repeat(2084), display: 'y'.repeat(2084), tooltip: 'z'.repeat(255) });
    await expectXlvErrorAsync(
      () => bytesOf(writeWorkbookStream({ sheets: [{ name: 'S', hyperlinks: [{ ref: 'A1' }], rows: [] }] })),
      'XLV_E_INVALID_HYPERLINK',
    );
  });

  it('does not let a property named __proto__ through', () => {
    const sheet = createWorkbook().addSheet('S');
    expectXlvError(() => sheet.link('A1', JSON.parse('{"__proto__": {"url": "x"}, "location": "S!A1"}') as never), 'XLV_E_INVALID_HYPERLINK', { property: '__proto__' });
  });
});

describe('review fixes (D-D05)', () => {
  it('refuses URLs with characters that a relationship target cannot carry', () => {
    const s = createWorkbook().addSheet('S');
    for (const url of ['https://e.com/￿', 'https://e.com/￾', 'https://e.com/a b\tc', 'https://e.com/\n']) expectXlvError(() => s.link('A1', { url }), 'XLV_E_INVALID_HYPERLINK');
  });
});
