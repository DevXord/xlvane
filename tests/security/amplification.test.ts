// Small files that try to make the reader do disproportionate work (found by review).
import { describe, expect, it } from 'vitest';
import { inspect } from '../../packages/xlvane/src/inspect';
import { DEFAULT_READ_LIMITS } from '../../packages/xlvane/src/limits';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { openWorkbookStream } from '../../packages/xlvane/src/stream-read';
import { parseXml } from '../../packages/xlvane/src/xml/parser';
import { parseXmlTree } from '../../packages/xlvane/src/xml/tree';
import { expectXlvError, expectXlvErrorAsync } from '../helpers/expect-error';
import { buildXlsx, c, row } from '../helpers/xlsx-builder';

function timed<T>(fn: () => T): { result: T; ms: number } {
  const start = performance.now();
  const result = fn();
  return { result, ms: performance.now() - start };
}

describe('amplification attacks', () => {
  it('classifies each number format once, however many style indexes cells use', () => {
    const code = 'G'.repeat(200_000);
    const cells = Array.from({ length: 5000 }, (_, i) => row(i + 1, c(`A${i + 1}`, 1, `s="${i + 1}"`))).join('');
    const xlsx = buildXlsx({
      styles: `<numFmts count="1"><numFmt numFmtId="0" formatCode="${code}"/></numFmts><cellXfs count="1"><xf numFmtId="0"/></cellXfs>`,
      sheets: [{ name: 'S', data: cells }],
    });
    const { result, ms } = timed(() => readWorkbook(xlsx));
    expect(result.getSheet('S').cellCount).toBe(5000);
    expect(ms).toBeLessThan(2000);
  });

  it('rejects namespace names longer than maxXmlNameLength', () => {
    const uri = `urn:${'x'.repeat(2000)}`;
    expectXlvError(
      () => parseXml(`<a xmlns:p="${uri}" p:b="1"/>`, { start: () => {}, end: () => {}, text: () => {} }, { part: '/t.xml', limits: DEFAULT_READ_LIMITS }),
      'XLV_E_LIMIT_EXCEEDED',
      { limit: 'maxXmlNameLength' },
    );
  });

  it('counts attributes against maxXmlTreeElements', () => {
    const attrs = Array.from({ length: 10 }, (_, i) => `a${i}="1"`).join(' ');
    const xml = `<r>${`<e ${attrs}/>`.repeat(10)}</r>`;
    expectXlvError(
      () => parseXmlTree(xml, { part: '/t.xml', limits: { ...DEFAULT_READ_LIMITS, maxXmlTreeElements: 50 } }),
      'XLV_E_LIMIT_EXCEEDED',
      { limit: 'maxXmlTreeElements' },
    );
    expect(parseXmlTree(xml, { part: '/t.xml', limits: { ...DEFAULT_READ_LIMITS, maxXmlTreeElements: 111 } }).children).toHaveLength(10);
  });

  it('checks merged ranges in time independent of merges × cells', () => {
    const merges = Array.from({ length: 20_000 }, () => '<mergeCell ref="Z1:Z1048576"/>').join('');
    const cells = Array.from({ length: 20_000 }, (_, i) => row(i + 1, c(`A${i + 1}`, 1))).join('');
    const xlsx = buildXlsx({ sheets: [{ name: 'M', data: cells, after: `<mergeCells>${merges}</mergeCells>` }] });
    const { result, ms } = timed(() => inspect(xlsx));
    expect(result.sheets[0]?.cellCount).toBe(20_000);
    expect(ms).toBeLessThan(3000);
  });

  it('does not slow down on warnings for a sheet with a very long name', () => {
    const name = 'N'.repeat(1_000_000);
    const cells = Array.from({ length: 20_000 }, (_, i) => row(i + 1, c(`A${i + 1}`, 'x', 't="q"'))).join('');
    const xlsx = buildXlsx({ sheets: [{ name, data: cells }] });
    const { result, ms } = timed(() => readWorkbook(xlsx));
    expect(result.warnings[0]).toMatchObject({ code: 'XLV_W_INVALID_VALUE', count: 20_000 });
    expect(ms).toBeLessThan(3000);
  });

  it('limits shared-string text repeated across cells (maxSharedStringChars)', async () => {
    // 2 000 cells point at one 100 000-character shared string: 200 M characters once copied.
    const cells = Array.from({ length: 2000 }, (_, i) => row(i + 1, c(`A${i + 1}`, 0, 't="s"'))).join('');
    const xlsx = buildXlsx({ sharedStrings: ['x'.repeat(100_000)], sheets: [{ name: 'S', data: cells }] });
    const limits = { maxSharedStringChars: 50_000_000 };
    expectXlvError(() => readWorkbook(xlsx, { limits }), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxSharedStringChars' });
    await expectXlvErrorAsync(async () => {
      const book = await openWorkbookStream(xlsx, { limits });
      for await (const _row of book.rows('S')) {
        // consume
      }
    }, 'XLV_E_LIMIT_EXCEEDED');
    // Within the limit the same file reads, and a second pass over a streamed sheet starts afresh.
    expect(readWorkbook(xlsx).getSheet('S').cellCount).toBe(2000);
    const book = await openWorkbookStream(xlsx, { limits: { maxSharedStringChars: 300_000_000 } });
    for (let pass = 0; pass < 2; pass++) {
      let n = 0;
      for await (const r of book.rows('S')) n += r.cells.length;
      expect(n).toBe(2000);
    }
  });
});

