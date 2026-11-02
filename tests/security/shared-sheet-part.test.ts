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
