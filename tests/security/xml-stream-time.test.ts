// The streaming XML parser stays linear on a long token that arrives in many pieces (1.1.1
// security review, D-H01): appending each piece to the held text and searching it again made a
// 32 MiB comment take 7 s, and 256 MiB (the largest part) minutes. Comments and processing
// instructions longer than maxXmlTextLength are refused, as text is, in whole-text and streaming
// parsing alike.
import { describe, expect, it } from 'vitest';
import { DEFAULT_READ_LIMITS } from '../../packages/xlvane/src/limits';
import { createXmlByteStreamParser, createXmlStreamParser, parseXml, type XmlHandler } from '../../packages/xlvane/src/xml/parser';
import { expectXlvError } from '../helpers/expect-error';

const handler: XmlHandler = { start() {}, end() {}, text() {} };
const options = { limits: DEFAULT_READ_LIMITS, part: '/x.xml' };
const piece = 'a'.repeat(65536);

/** Writes `open`, then `mib` MiB in 64 KiB pieces, then `close`, and returns the time taken. */
function streamed(open: string, mib: number, close: string, limits = DEFAULT_READ_LIMITS): number {
  const p = createXmlStreamParser(handler, { ...options, limits });
  const start = performance.now();
  p.write(open);
  for (let i = 0; i < mib * 16; i++) p.write(piece);
  p.write(close);
  p.end();
  return performance.now() - start;
}

describe('streaming parser on long tokens', () => {
  const big = { ...DEFAULT_READ_LIMITS, maxXmlTextLength: 64 * 1024 * 1024 };
  for (const [name, open, close] of [
    ['comment', '<r><!--', '--></r>'],
    ['processing instruction', '<r><?pi ', '?></r>'],
    ['CDATA section', '<r><![CDATA[', ']]></r>'],
    ['attribute value', '<r a="', '"/>'],
    ['XML declaration', '<?xml version="1.0"', '?><r/>'],
    ['text', '<r>', '</r>'],
  ] as const) {
    it(`a 48 MiB ${name} in 64 KiB pieces takes linear time`, () => {
      expect(streamed(open, 48, close, big)).toBeLessThan(5000);
    }, 60_000);
  }

  it('a start tag whose quoted values hold ">" in every piece stays linear', () => {
    const p = createXmlStreamParser(handler, { ...options, limits: big });
    const start = performance.now();
    p.write('<r a="');
    for (let i = 0; i < 48 * 16; i++) p.write(`>${piece.slice(1)}`);
    p.write('"/>');
    p.end();
    expect(performance.now() - start).toBeLessThan(5000);
  }, 60_000);

  it('a 48 MiB XML declaration given as bytes takes linear time', () => {
    const p = createXmlByteStreamParser(handler, { ...options, limits: big });
    const bytes = new TextEncoder().encode(piece);
    const start = performance.now();
    p.write(new TextEncoder().encode('<?xml version="1.0"'));
    for (let i = 0; i < 48 * 16; i++) p.write(bytes);
    p.write(new TextEncoder().encode('?><r/>'));
    p.end();
    expect(performance.now() - start).toBeLessThan(5000);
  }, 60_000);

  it('short tokens are delivered with the piece that completes them; after a long one, by end()', () => {
    const seen: string[] = [];
    const p = createXmlStreamParser({ start: (_ns, local) => seen.push(local), end() {}, text() {} }, options);
    p.write('<r><a/><b');
    expect(seen).toEqual(['r', 'a']);
    p.write('/>');
    expect(seen).toEqual(['r', 'a', 'b']);
    p.write('<c x="');
    for (let i = 0; i < 64; i++) p.write(piece);
    p.write('"/><d/>');
    // Held until as much input as the unfinished token has arrived, or until end().
    p.write('<f/>'.repeat(1 << 20));
    expect(seen.slice(0, 5)).toEqual(['r', 'a', 'b', 'c', 'd']);
    p.write('<e/>');
    expect(seen[seen.length - 1]).toBe('e');
    p.write('</r>');
    p.end();
  });
});

describe('comments and processing instructions longer than maxXmlTextLength', () => {
  const small = { ...DEFAULT_READ_LIMITS, maxXmlTextLength: 1000 };
  for (const [name, doc] of [
    ['comment', `<r><!--${'a'.repeat(1001)}--></r>`],
    ['processing instruction', `<r><?pi ${'a'.repeat(1001)}?></r>`],
    ['comment before the root', `<!--${'a'.repeat(1001)}--><r/>`],
    ['XML declaration', `<?xml version="1.0"${' '.repeat(1001)}?><r/>`],
  ] as const) {
    it(`a ${name} is refused in whole-text and streaming parsing`, () => {
      const error = expectXlvError(() => parseXml(doc, handler, { ...options, limits: small }), 'XLV_E_LIMIT_EXCEEDED');
      expect(error.details?.limit).toBe('maxXmlTextLength');
      const streamError = expectXlvError(() => {
        const p = createXmlStreamParser(handler, { ...options, limits: small });
        for (let i = 0; i < doc.length; i += 7) p.write(doc.slice(i, i + 7));
        p.end();
      }, 'XLV_E_LIMIT_EXCEEDED');
      expect(streamError.message).toBe(error.message);
      expect(streamError.location).toEqual(error.location);
    });
  }

  it('a comment and a processing instruction of exactly the limit are accepted', () => {
    const small = { ...DEFAULT_READ_LIMITS, maxXmlTextLength: 1000 };
    expect(() => parseXml(`<r><!--${'a'.repeat(1000)}--><?pi ${'a'.repeat(995)}?></r>`, handler, { ...options, limits: small })).not.toThrow();
  });

  it('an unterminated 48 MiB comment fails at the limit while it arrives, holding little', () => {
    const p = createXmlStreamParser(handler, options);
    p.write('<r><!--');
    let error: unknown = null;
    let written = 0;
    try {
      for (let i = 0; i < 48 * 16; i++) {
        p.write(piece);
        written += piece.length;
        expect(p.buffered).toBeLessThan(4 * DEFAULT_READ_LIMITS.maxXmlTextLength);
      }
    } catch (e) {
      error = e;
    }
    expect((error as { code?: string }).code).toBe('XLV_E_LIMIT_EXCEEDED');
    expect(written).toBeLessThan(4 * DEFAULT_READ_LIMITS.maxXmlTextLength);
  });
});

describe('unfinished tokens are not held without limit (independent review of the patch)', () => {
  const limits = { ...DEFAULT_READ_LIMITS, maxXmlTextLength: 1 << 20 };
  for (const [name, open] of [
    ['CDATA section', '<r><![CDATA['],
    ['attribute value', '<r a="'],
    ['reference without ";"', '<r>&'],
  ] as const) {
    it(`a ${name} that never ends fails near the limit, holding little`, () => {
      const p = createXmlStreamParser(handler, { ...options, limits });
      p.write(open);
      let error: unknown = null;
      let written = 0;
      try {
        for (let i = 0; i < 1024; i++) {
          p.write(piece);
          written += piece.length;
          expect(p.buffered).toBeLessThan(4 * limits.maxXmlTextLength);
        }
      } catch (e) {
        error = e;
      }
      expect((error as { code?: string }).code).toMatch(/^XLV_E_(LIMIT_EXCEEDED|XML_MALFORMED)$/);
      expect(written).toBeLessThan(4 * limits.maxXmlTextLength);
    });
  }

  it('the same attribute value as one text gives the same error', () => {
    const doc = `<r a="${'a'.repeat((1 << 20) + 1)}"/>`;
    const whole = expectXlvError(() => parseXml(doc, handler, { ...options, limits }), 'XLV_E_LIMIT_EXCEEDED');
    const streamed = expectXlvError(() => {
      const p = createXmlStreamParser(handler, { ...options, limits });
      for (let i = 0; i < doc.length; i += 65536) p.write(doc.slice(i, i + 65536));
      p.end();
    }, 'XLV_E_LIMIT_EXCEEDED');
    expect(streamed.message).toBe(whole.message);
    expect(streamed.location).toEqual(whole.location);
  });

  it('an XML declaration given as bytes that never ends fails near the limit', () => {
    const p = createXmlByteStreamParser(handler, { ...options, limits });
    const bytes = new TextEncoder().encode(piece);
    p.write(new TextEncoder().encode('<?xml version="1.0"'));
    let error: unknown = null;
    let written = 0;
    try {
      for (let i = 0; i < 1024; i++) {
        p.write(bytes);
        written += bytes.length;
      }
    } catch (e) {
      error = e;
    }
    expect((error as { code?: string }).code).toBe('XLV_E_LIMIT_EXCEEDED');
    expect(written).toBeLessThan(4 * limits.maxXmlTextLength);
  });
});

describe('a comment whose "--" ends a piece (found by fuzzing, regression of 1.1.1)', () => {
  it('is complete when the next piece starts with ">"', () => {
    const doc = '<!--!;--><a></a>';
    const limits = { ...DEFAULT_READ_LIMITS, maxXmlTextLength: 2 };
    const p = createXmlStreamParser(handler, { ...options, limits });
    expect(() => {
      for (const ch of doc) p.write(ch);
      p.end();
    }).not.toThrow();
  });

  it('is not taken for a comment longer than the limit when much text without "--" follows', () => {
    const limits = { ...DEFAULT_READ_LIMITS, maxXmlTextLength: 1000 };
    const p = createXmlStreamParser(handler, { ...options, limits });
    expect(() => {
      p.write('<r><!--a--');
      p.write(`><t>${'x'.repeat(999)}</t>`);
      for (let i = 0; i < 20; i++) p.write(`<t>${'x'.repeat(999)}</t>`);
      p.write('</r>');
      p.end();
    }).not.toThrow();
  });
});
