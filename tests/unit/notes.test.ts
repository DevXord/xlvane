// Notes (comments): ECMA-376 Part 1 §18.7, §12.3.3; Part 4 §8.2, §15.4.1.1, §19.4;
// [MS-OI29500] §2.1.689 … §2.1.692; [MS-XLSX] §2.3.7.3; docs/facts/sml-write.md W2-45 … W2-48;
// docs/decisions.md D-A12.
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

/** The notes without their box settings, which tests/unit/note-box.test.ts checks. */
const withoutBox = (notes: readonly { box?: unknown }[]): unknown[] => notes.map(({ box: _box, ...rest }) => rest);

const rels = (items: string): string =>
  `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`;

async function bytesOf(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

function parts(bytes: Uint8Array): Record<string, string> {
  return Object.fromEntries(Object.entries(unzipSync(bytes)).map(([k, v]) => [k, strFromU8(v)]));
}

const COMMENTS = (ns: string, body: string): string => `<?xml version="1.0" encoding="UTF-8"?><comments xmlns="${ns}">${body}</comments>`;

function fileWithComments(comments: string, opts: { strict?: boolean; extraRels?: string } = {}): Uint8Array {
  const ns = opts.strict ? NS.strict : NS.transitional;
  return buildXlsx({
    strict: opts.strict === true,
    sheets: [{ name: 'Notes', data: '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>' }],
    extraParts: {
      'xl/worksheets/_rels/sheet1.xml.rels': rels(`<Relationship Id="rId1" Type="${ns.rel}comments" Target="../comments1.xml"/>${opts.extraRels ?? ''}`),
      'xl/comments1.xml': COMMENTS(ns.main, comments),
    },
  });
}

describe('reading notes', () => {
  const body =
    '<authors><author>Ann</author><author>Bob_x000A_Smith</author></authors><commentList>' +
    '<comment ref="A1" authorId="0"><text><t>Plain note</t></text></comment>' +
    '<comment ref="b1" authorId="1"><text><r><rPr><b/><sz val="9"/><rFont val="Tahoma"/></rPr><t>Bob:</t></r><r><t xml:space="preserve"> why?</t></r><rPh sb="0" eb="1"><t>x</t></rPh></text></comment>' +
    '<comment ref="C5" authorId="7"><text><t>orphan</t></text></comment>' +
    '<comment ref="A1:B2" authorId="0"><text><t>range</t></text></comment>' +
    '</commentList>';

  it('lists the notes with text, author and rich text runs, and attaches them to cells', () => {
    const sheet = readWorkbook(fileWithComments(body)).getSheet('Notes');
    expect(sheet.notes).toEqual([
      { ref: 'A1', text: 'Plain note', author: 'Ann' },
      { ref: 'B1', text: 'Bob: why?', author: 'Bob\nSmith', runs: [{ text: 'Bob:', font: { bold: true, size: 9, name: 'Tahoma' } }, { text: ' why?' }] },
      { ref: 'C5', text: 'orphan', author: null },
    ]);
    expect(sheet.cell('A1')?.note?.text).toBe('Plain note');
    expect(sheet.cell('B1')?.note?.author).toBe('Bob\nSmith');
    // C5 has no cell: the note is only in the list.
    expect(sheet.cell('C5')).toBeUndefined();
  });

  it('warns about an unknown author and a range reference', () => {
    const warnings = readWorkbook(fileWithComments(body)).warnings.filter((w) => w.code === 'XLV_W_INVALID_VALUE');
    expect(warnings.length).toBe(1);
    expect(warnings[0]?.count).toBe(2);
  });

  it('reads Strict files, and ignores malformed parts with a warning', () => {
    const strict = readWorkbook(fileWithComments('<authors><author>A</author></authors><commentList><comment ref="A1" authorId="0"><text><t>s</t></text></comment></commentList>', { strict: true }));
    expect(strict.getSheet('Notes').notes).toEqual([{ ref: 'A1', text: 's', author: 'A' }]);
    const broken = readWorkbook(fileWithComments('<authors><author>A</author></authors><commentList><comment'));
    expect(broken.getSheet('Notes').notes).toEqual([]);
    expect(broken.warnings.some((w) => w.code === 'XLV_W_INVALID_VALUE' && w.message.includes('not well-formed'))).toBe(true);
  });

  it('reports threaded comments and a missing comments part', () => {
    const threaded = fileWithComments(
      '<authors><author>tc={11111111-2222-3333-4444-555555555555}</author></authors><commentList><comment ref="A1" authorId="0"><text><t>[Threaded comment]</t></text></comment></commentList>',
      { extraRels: '<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/>' },
    );
    const wb = readWorkbook(threaded);
    expect(wb.warnings.map((w) => w.code)).toContain('XLV_W_THREADED_COMMENTS');
    expect(wb.getSheet('Notes').notes[0]?.author).toMatch(/^tc=/);
    const missing = buildXlsx({
      sheets: [{ name: 'S', data: '' }],
      extraParts: { 'xl/worksheets/_rels/sheet1.xml.rels': rels(`<Relationship Id="rId1" Type="${NS.transitional.rel}comments" Target="../comments9.xml"/>`) },
    });
    expect(readWorkbook(missing).warnings.map((w) => w.code)).toContain('XLV_W_PART_MISSING');
  });

  it('rejects a DTD in the comments part', () => {
    const file = buildXlsx({
      sheets: [{ name: 'S', data: '' }],
      extraParts: {
        'xl/worksheets/_rels/sheet1.xml.rels': rels(`<Relationship Id="rId1" Type="${NS.transitional.rel}comments" Target="../comments1.xml"/>`),
        'xl/comments1.xml': '<?xml version="1.0"?><!DOCTYPE x [<!ENTITY a "b">]><comments/>',
      },
    });
    expectXlvError(() => readWorkbook(file), 'XLV_E_XML_DTD_FORBIDDEN');
  });

  it('the streaming reader sets notes on cells and lists them', async () => {
    const book = await openWorkbookStream(fileWithComments(body));
    const cells = [];
    for await (const row of book.rows('Notes')) cells.push(...row.cells);
    expect(cells.find((c) => c.ref === 'A1')?.note?.text).toBe('Plain note');
    expect((await book.notes('Notes')).map((n) => n.ref)).toEqual(['A1', 'B1', 'C5']);
    await expectXlvErrorAsync(() => book.notes('Nope'), 'XLV_E_SHEET_NOT_FOUND');
  });
});

describe('writing notes', () => {
  const build = (): Uint8Array => {
    const wb = createWorkbook();
    const sheet = wb.addSheet('Budget');
    sheet.set('B2', 1250).note('B2', 'Estimate & check');
    sheet.note('A1', { runs: [{ text: 'Ann: ', font: { bold: true } }, { text: 'see <B2>' }], author: 'Ann' });
    sheet.note('C3', { text: 'by Ann too', author: 'Ann' });
    sheet.set('A5', 'link').link('A5', { url: 'https://example.com' });
    wb.addSheet('Plain').set('A1', 1);
    return writeWorkbook(wb);
  };

  it('writes the Comments part, the VML part, legacyDrawing, relationships and content types', () => {
    const p = parts(build());
    const comments = p['xl/comments1.xml'] ?? '';
    expect(comments).toContain('<authors><author>Ann</author><author></author></authors>');
    expect(comments).toContain('<comment ref="A1" authorId="0"><text><r><rPr><b/></rPr><t xml:space="preserve">Ann: </t></r><r><t>see &lt;B2&gt;</t></r></text></comment>');
    expect(comments).toContain('<comment ref="B2" authorId="1"><text><t>Estimate &amp; check</t></text></comment>');
    // Row, then column order.
    expect(comments.indexOf('ref="A1"')).toBeLessThan(comments.indexOf('ref="B2"'));
    expect(comments.indexOf('ref="B2"')).toBeLessThan(comments.indexOf('ref="C3"'));
    const vml = p['xl/drawings/vmlDrawing1.vml'] ?? '';
    expect(vml).toMatch(/^<xml xmlns:v="urn:schemas-microsoft-com:vml"/);
    expect(vml).toContain('<o:idmap v:ext="edit" data="1"/>');
    expect(vml).toContain('id="_x0000_s1025"');
    expect(vml).toContain('id="_x0000_s1027"');
    expect(vml).toContain('<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>2, 15, 0, 10, 4, 15, 4, 4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>1</x:Row><x:Column>1</x:Column></x:ClientData>');
    expect(vml).not.toContain('x:Visible');
    expect(vml).toContain('visibility:hidden');
    const sheetRels = p['xl/worksheets/_rels/sheet1.xml.rels'] ?? '';
    expect(sheetRels).toContain(`Type="${NS.transitional.rel}comments" Target="../comments1.xml"`);
    expect(sheetRels).toContain(`Id="rId2" Type="${NS.transitional.rel}vmlDrawing" Target="../drawings/vmlDrawing1.vml"`);
    expect(p['xl/worksheets/sheet1.xml']).toMatch(/<\/hyperlinks><legacyDrawing r:id="rId2"\/><\/worksheet>$/);
    const types = p['[Content_Types].xml'] ?? '';
    expect(types).toContain('<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>');
    expect(types).toContain('<Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>');
    expect(p['xl/comments2.xml']).toBeUndefined();
    expect(p['xl/worksheets/sheet2.xml']).not.toContain('legacyDrawing');
  });

  it('reads back what it wrote', () => {
    const sheet = readWorkbook(build()).getSheet('Budget');
    expect(withoutBox(sheet.notes)).toEqual([
      { ref: 'A1', text: 'Ann: see <B2>', author: 'Ann', runs: [{ text: 'Ann: ', font: { bold: true } }, { text: 'see <B2>' }] },
      { ref: 'B2', text: 'Estimate & check', author: '' },
      { ref: 'C3', text: 'by Ann too', author: 'Ann' },
    ]);
    expect(sheet.cell('B2')?.note?.text).toBe('Estimate & check');
    expect(sheet.cell('A5')?.hyperlink?.target).toBe('https://example.com');
  });

  it('replaces and removes notes, and keeps output deterministic', () => {
    const make = (): Uint8Array => {
      const b = createWorkbook();
      b.addSheet('S').note('A1', 'one').note('A1', 'two').note('B1', 'gone').note('B1', null);
      return writeWorkbook(b);
    };
    expect(withoutBox(readWorkbook(make()).getSheet('S').notes)).toEqual([{ ref: 'A1', text: 'two', author: '' }]);
    expect(make()).toEqual(make());
  });

  it('validates notes', () => {
    const sheet = createWorkbook().addSheet('S');
    expectXlvError(() => sheet.note('A0', 'x'), 'XLV_E_INVALID_CELL_REF');
    expectXlvError(() => sheet.note('A1', 42 as never), 'XLV_E_INVALID_VALUE');
    expectXlvError(() => sheet.note('A1', { text: 'a', runs: [{ text: 'b' }] }), 'XLV_E_INVALID_VALUE');
    expectXlvError(() => sheet.note('A1', {} as never), 'XLV_E_INVALID_VALUE');
    expectXlvError(() => sheet.note('A1', { text: 'a', extra: 1 } as never), 'XLV_E_INVALID_VALUE');
    expectXlvError(() => sheet.note('A1', 'x'.repeat(32_768)), 'XLV_E_INVALID_VALUE');
    expectXlvError(() => sheet.note('A1', { text: 'a', author: 'x'.repeat(65_536) }), 'XLV_E_INVALID_VALUE');
    expectXlvError(() => sheet.note('A1', { runs: [] }), 'XLV_E_INVALID_VALUE');
    expectXlvError(() => sheet.note('A1', { runs: [{ text: 'a', font: { size: 0 } }] }), 'XLV_E_INVALID_STYLE');
    expectXlvError(() => sheet.note('A1', JSON.parse('{"__proto__":{"x":1},"text":"a"}') as never), 'XLV_E_INVALID_VALUE');
    expect(() => sheet.note('A1', { text: 'x'.repeat(32_767), author: 'y'.repeat(65_535) })).not.toThrow();
  });

  it('uses several shape-id blocks for more than 1024 notes, and separate blocks per sheet', () => {
    const wb = createWorkbook();
    const big = wb.addSheet('Big');
    for (let r = 1; r <= 1100; r++) big.note(`A${String(r)}`, String(r));
    wb.addSheet('Small').note('A1', 'x');
    const p = parts(writeWorkbook(wb));
    expect(p['xl/drawings/vmlDrawing1.vml']).toContain('data="1,2"');
    expect(p['xl/drawings/vmlDrawing1.vml']).toContain('id="_x0000_s2124"');
    expect(p['xl/drawings/vmlDrawing2.vml']).toContain('data="3"');
    expect(p['xl/drawings/vmlDrawing2.vml']).toContain('id="_x0000_s3073"');
    expect(readWorkbook(writeWorkbook(wb)).getSheet('Big').notes.length).toBe(1100);
  });
});

describe('streaming notes and links per row', () => {
  it('writes notes and hyperlinks given with the rows', async () => {
    const bytes = await bytesOf(
      writeWorkbookStream({
        sheets: [
          {
            name: 'Data',
            rows: [
              { values: ['id', 'value'], notes: ['Identifier'], hyperlinks: [undefined, { url: 'https://example.com/value' }] },
              [1, 2],
              { values: [3, 4], notes: [undefined, { runs: [{ text: 'big', font: { italic: true } }], author: 'Ann' }] },
            ],
          },
          { name: 'Empty', rows: [[1]] },
        ],
      }),
    );
    const wb = readWorkbook(bytes);
    const sheet = wb.getSheet('Data');
    expect(withoutBox(sheet.notes)).toEqual([
      { ref: 'A1', text: 'Identifier', author: '' },
      { ref: 'B3', text: 'big', author: 'Ann', runs: [{ text: 'big', font: { italic: true } }] },
    ]);
    expect(sheet.cell('B1')?.hyperlink?.target).toBe('https://example.com/value');
    const names = Object.keys(unzipSync(bytes));
    // Package parts follow the sheets, because the notes are known only then.
    expect(names.indexOf('[Content_Types].xml')).toBeGreaterThan(names.indexOf('xl/comments1.xml'));
    expect(parts(bytes)['[Content_Types].xml']).toContain('/xl/comments1.xml');
  });

  it('streams notes from an iterable of sheets and validates per-row input', async () => {
    async function* sheets() {
      yield { name: 'A', rows: [{ values: [1], notes: ['n'] }] };
    }
    const bytes = await bytesOf(writeWorkbookStream({ sheets: sheets() }));
    expect(withoutBox(readWorkbook(bytes).getSheet('A').notes)).toEqual([{ ref: 'A1', text: 'n', author: '' }]);
    await expectXlvErrorAsync(() => bytesOf(writeWorkbookStream({ sheets: [{ name: 'S', rows: [{ values: [1], notes: 'x' as never }] }] })), 'XLV_E_INVALID_OPTION');
    await expectXlvErrorAsync(() => bytesOf(writeWorkbookStream({ sheets: [{ name: 'S', rows: [{ values: [1], notes: [42 as never] }] }] })), 'XLV_E_INVALID_VALUE');
  });
});

describe('notes elsewhere', () => {
  it('editCells keeps the notes of a file', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', 1).note('A1', 'keep');
    const edited = editCells(writeWorkbook(wb), [{ sheet: 'S', cell: 'B1', value: 2 }]);
    const sheet = readWorkbook(edited.bytes).getSheet('S');
    expect(withoutBox(sheet.notes)).toEqual([{ ref: 'A1', text: 'keep', author: '' }]);
    expect(sheet.cell('B1')?.value).toBe(2);
  });

  it('inspect counts notes and lists up to 10 cells', () => {
    const wb = createWorkbook();
    const sheet = wb.addSheet('S');
    for (let r = 1; r <= 12; r++) sheet.note(`B${String(r)}`, 'x');
    const report = inspect(readWorkbook(writeWorkbook(wb)));
    expect(report.sheets[0]?.noteCount).toBe(12);
    expect(report.sheets[0]?.notes).toEqual(['B1', 'B2', 'B3', 'B4', 'B5', 'B6', 'B7', 'B8', 'B9', 'B10']);
  });
});
