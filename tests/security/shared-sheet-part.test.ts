// Several sheets whose relationships lead to the same worksheet part made the part be read and
// parsed once per sheet, outside the limit on decompressed bytes: a 1 MB file could make the
// reader parse gigabytes (1.1.1 security review, D-H01). The first sheet keeps the part; the
// others are returned without cells, with XLV_W_SHEET_PART_SHARED, in both readers.
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { openWorkbookStream } from '../../packages/xlvane/src/stream-read';
import { buildXlsx } from '../helpers/xlsx-builder';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** `n` sheets; relationship rId1 leads to sheet1.xml and every other rIdN does too. */
function shared(n: number, data: string): Uint8Array {
  const rels = Array.from({ length: n }, (_, i) => `<Relationship Id="rId${String(i + 1)}" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/>`).join('');
  return buildXlsx({
    sheets: Array.from({ length: n }, (_, i) => ({ name: `S${String(i + 1)}`, data: i === 0 ? data : '' })),
    extraParts: { 'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>` },
  });
}

describe('a worksheet part shared by several sheets', () => {
  const row = '<row r="1"><c r="A1"><v>1</v></c></row>';

  it('readWorkbook reads it for the first sheet only', () => {
    const wb = readWorkbook(shared(3, row));
    expect(wb.sheets.map((s) => s.name)).toEqual(['S1', 'S2', 'S3']);
    expect(wb.sheets[0]?.cell('A1')?.value).toBe(1);
    expect(wb.sheets[1]?.cell('A1')).toBeUndefined();
    expect([...(wb.sheets[2]?.cells() ?? [])]).toHaveLength(0);
    const shared_ = wb.warnings.filter((w) => w.code === 'XLV_W_SHEET_PART_SHARED');
    expect(shared_.map((w) => w.location.sheet)).toEqual(['S2', 'S3']);
    expect(shared_[0]?.message).toContain('"S1"');
  });

  it('openWorkbookStream reads it for the first sheet only', async () => {
    const book = await openWorkbookStream(shared(3, row));
    const counts: number[] = [];
    for (const s of book.sheets) {
      let n = 0;
      for await (const _ of book.rows(s.name)) n++;
      counts.push(n);
    }
    expect(counts).toEqual([1, 0, 0]);
    expect(book.warnings.filter((w) => w.code === 'XLV_W_SHEET_PART_SHARED')).toHaveLength(2);
  });

  it('200 sheets sharing a part with 4 MiB of comments are read once, quickly', () => {
    let filler = '';
    for (let i = 0; i < 4096; i++) filler += `<!--${String(i).padStart(1000, 'x')}-->`;
    const bytes = shared(200, filler + row);
    const start = performance.now();
    const wb = readWorkbook(bytes);
    expect(performance.now() - start).toBeLessThan(2000);
    expect(wb.sheets).toHaveLength(200);
  }, 30_000);
});

const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

describe('shared parts found by the independent review of the patch', () => {
  it('a worksheet part named with other letter cases is still one part', () => {
    const n = 4;
    const targets = ['worksheets/sheet1.xml', 'Worksheets/sheet1.xml', 'wOrksheets/SHEET1.xml', 'WORKSHEETS/sheet1.XML'];
    const rels = targets.slice(0, n).map((t, i) => `<Relationship Id="rId${String(i + 1)}" Type="${REL}/worksheet" Target="${t}"/>`).join('');
    const bytes = buildXlsx({
      sheets: Array.from({ length: n }, (_, i) => ({ name: `S${String(i + 1)}`, data: i === 0 ? '<row r="1"><c r="A1"><v>1</v></c></row>' : '' })),
      extraParts: { 'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="${RELS_NS}">${rels}</Relationships>` },
    });
    const wb = readWorkbook(bytes);
    expect(wb.sheets.map((s) => [...s.cells()].length)).toEqual([1, 0, 0, 0]);
    expect(wb.warnings.filter((w) => w.code === 'XLV_W_SHEET_PART_SHARED')).toHaveLength(3);
  });

  /** `n` sheets with their own parts, each relating to the same comments and VML part; sheet 1 also twice to one table. */
  function sharedRelated(n: number, comments: string): Uint8Array {
    const extraParts: Record<string, string> = {
      'xl/comments1.xml': comments,
      'xl/drawings/vmlDrawing1.vml': '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shape><x:ClientData ObjectType="Note"><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData></v:shape></xml>',
      'xl/tables/table1.xml': '<?xml version="1.0"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" displayName="T" ref="A1:A2"><tableColumns count="1"><tableColumn id="1" name="a"/></tableColumns></table>',
    };
    for (let i = 1; i <= n; i++) {
      const own = `<Relationship Id="rId1" Type="${REL}/comments" Target="../comments1.xml"/><Relationship Id="rId2" Type="${REL}/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>`;
      const tables = i === 1 ? `<Relationship Id="rId3" Type="${REL}/table" Target="../tables/table1.xml"/><Relationship Id="rId4" Type="${REL}/table" Target="../Tables/TABLE1.xml"/>` : '';
      extraParts[`xl/worksheets/_rels/sheet${String(i)}.xml.rels`] = `<?xml version="1.0"?><Relationships xmlns="${RELS_NS}">${own}${tables}</Relationships>`;
    }
    return buildXlsx({ sheets: Array.from({ length: n }, (_, i) => ({ name: `S${String(i + 1)}`, data: '<row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c></row><row r="2"><c r="A2"><v>1</v></c></row>' })), extraParts });
  }
  const note = '<?xml version="1.0"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>x</author></authors><commentList><comment ref="A1" authorId="0"><text><t>hi</t></text></comment></commentList></comments>';

  it('comments, VML and table parts shared by sheets are read for the first sheet only, and a table once', async () => {
    const wb = readWorkbook(sharedRelated(3, note));
    expect(wb.sheets.map((s) => s.notes.length)).toEqual([1, 0, 0]);
    expect(wb.sheets[0]?.tables).toHaveLength(1);
    // One warning per sheet and part: the comments and the VML part of S2 and S3.
    expect(wb.warnings.filter((w) => w.code === 'XLV_W_SHEET_PART_SHARED').map((w) => w.location.sheet)).toEqual(['S2', 'S2', 'S3', 'S3']);
    const book = await openWorkbookStream(sharedRelated(3, note));
    const notes: number[] = [];
    for (const s of book.sheets) {
      let n = 0;
      for await (const r of book.rows(s.name)) n += r.cells.filter((c) => c.note !== undefined).length;
      notes.push(n);
    }
    expect(notes).toEqual([1, 0, 0]);
  });

  it('200 sheets sharing a 4 MiB comments part are read quickly', () => {
    let list = '';
    for (let i = 0; i < 4096; i++) list += `<comment ref="B${String(i + 1)}" authorId="0"><text><t>${String(i).padStart(1000, 'x')}</t></text></comment>`;
    const big = `<?xml version="1.0"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>x</author></authors><commentList>${list}</commentList></comments>`;
    const bytes = sharedRelated(200, big);
    const start = performance.now();
    readWorkbook(bytes);
    expect(performance.now() - start).toBeLessThan(3000);
  }, 60_000);
});
