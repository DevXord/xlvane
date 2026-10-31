// The chunked XML parsers give the same events, spans and errors as whole-text parsing.
import { describe, expect, it } from 'vitest';
import { XlvError } from '../../packages/xlvane/src/errors';
import { DEFAULT_READ_LIMITS } from '../../packages/xlvane/src/limits';
import {
  createXmlByteStreamParser,
  createXmlStreamParser,
  decodeXml,
  parseXml,
  type XmlHandler,
  type XmlLimits,
} from '../../packages/xlvane/src/xml/parser';

const PART = '/xl/t.xml';

type Event = readonly unknown[];

function recorder(): { handler: XmlHandler; events: Event[] } {
  const events: Event[] = [];
  return {
    events,
    handler: {
      start: (ns, local, attrs, span) => events.push(['start', ns, local, [...attrs], span.start, span.end, span.qname, span.empty]),
      end: (ns, local, span) => events.push(['end', ns, local, span.start, span.end]),
      text: (value) => events.push(['text', value]),
    },
  };
}

type Outcome = { events: Event[]; error: { code: string; message: string; location: unknown } | null };

function capture(run: (handler: XmlHandler) => void): Outcome {
  const { handler, events } = recorder();
  try {
    run(handler);
    return { events, error: null };
  } catch (error) {
    if (!(error instanceof XlvError)) throw error;
    return { events, error: { code: error.code, message: error.message, location: error.location } };
  }
}

function whole(xml: string, limits: XmlLimits = DEFAULT_READ_LIMITS): Outcome {
  return capture((h) => parseXml(xml, h, { part: PART, limits }));
}

function chunked(xml: string, sizes: number[], limits: XmlLimits = DEFAULT_READ_LIMITS): Outcome {
  return capture((h) => {
    const p = createXmlStreamParser(h, { part: PART, limits });
    let at = 0;
    let k = 0;
    while (at < xml.length) {
      const n = Math.max(1, sizes[k++ % sizes.length] ?? 1);
      p.write(xml.slice(at, at + n));
      at += n;
    }
    p.end();
  });
}

const DOCS = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<a xmlns="urn:a" xmlns:x="urn:x"><x:b k="1" x:k=\'2\'>t &amp; &#x41;&#66;<![CDATA[<raw>]]>u</x:b><!-- c -- d? no --><?pi data?><c/>\r\r\n</a>\n',
  '<a>😀 text\r\nmore</a>',
  '<a b="x > y" c=\'<\'/>',
  '<a>]]></a>',
  '<a>&unknown;</a>',
  '<!DOCTYPE a><a/>',
  '<a><b></a>',
  '<a>text',
  '<a/><b/>',
  '<?xml version="1.0"?><?xml version="1.0"?><a/>',
  '<a>&#0;</a>',
  '<a>bad \u0001 char</a>',
  '<a>lone \uD800 high</a>',
  '<a>\uDE00</a>',
  '<a x="1" x="2"/>',
  '<a><!-- unterminated </a>',
  '<a><![CDATA[ open </a>',
  '<a>&amp</a>',
  '<a>&a&b;</a>',
  ' \n\n <a>\n</a>\r\n x',
  '',
  '<?xml',
  '<?xml version="1.0"?>',
  '<a></a',
];

describe('createXmlStreamParser', () => {
  for (const [i, doc] of DOCS.entries()) {
    it(`matches whole-text parsing for document ${String(i)} in any chunking`, () => {
      const expected = whole(doc);
      for (const sizes of [[1], [2], [3], [7], [1, 5, 2], [1000]]) {
        const got = chunked(doc, sizes);
        expect(got.error).toEqual(expected.error);
        if (expected.error === null) expect(got.events).toEqual(expected.events);
      }
    });
  }

  it('gives spans as offsets into the whole normalised text', () => {
    const doc = '<a>\r\n<b x="1"/>\r\n</a>';
    const got = chunked(doc, [1]);
    expect(got.events).toEqual(whole(doc).events);
    expect(got.events[2]).toEqual(['start', '', 'b', [['x', '1']], 4, 14, 'b', true]);
  });

  it('delivers the character data between two tags in one call', () => {
    const got = chunked('<a>one &amp; two ]] three</a>', [1]);
    expect(got.events.filter((e) => e[0] === 'text')).toEqual([['text', 'one & two ]] three']]);
  });

  it('enforces the text limit with the location of the start of the text', () => {
    const limits = { ...DEFAULT_READ_LIMITS, maxXmlTextLength: 10 };
    const doc = `<a>\n  ${'x'.repeat(30)}</a>`;
    expect(chunked(doc, [3], limits).error).toEqual(whole(doc, limits).error);
    expect(whole(doc, limits).error?.code).toBe('XLV_E_LIMIT_EXCEEDED');
  });

  it('drops consumed input, so memory follows the largest token, not the document', () => {
    let events = 0;
    const p = createXmlStreamParser({ start: () => void events++, end() {}, text() {} }, { part: PART, limits: DEFAULT_READ_LIMITS });
    p.write('<sheetData>');
    const row = '<row r="1"><c r="A1"><v>1</v></c></row>';
    let held = 0;
    for (let i = 0; i < 20_000; i++) {
      // Rows arrive split in the middle of a tag.
      p.write(row.slice(0, 13));
      p.write(row.slice(13));
      held = Math.max(held, p.buffered);
    }
    p.write('</sheetData>');
    p.end();
    expect(events).toBe(1 + 20_000 * 3);
    // 20 000 rows are about 800 000 characters; the parser never holds more than its compaction window.
    expect(held).toBeLessThan(2 * 65_536 + row.length);
  });

  it('processes a text run longer than the text limit in parts, so it is not held whole', () => {
    const limits = { ...DEFAULT_READ_LIMITS, maxXmlTextLength: 1000 };
    let held = 0;
    const p = createXmlStreamParser({ start() {}, end() {}, text() {} }, { part: PART, limits });
    p.write('<a>');
    let error: unknown = null;
    try {
      for (let i = 0; i < 100; i++) {
        p.write('&amp;&#65;'.repeat(20));
        held = Math.max(held, p.buffered);
      }
    } catch (e) {
      error = e;
    }
    expect(error).toBeInstanceOf(XlvError);
    expect((error as XlvError).code).toBe('XLV_E_LIMIT_EXCEEDED');
    expect(held).toBeLessThan(2000 + 200);
  });

  it('rejects write after end', () => {
    const p = createXmlStreamParser({ start() {}, end() {}, text() {} }, { part: PART, limits: DEFAULT_READ_LIMITS });
    p.write('<a/>');
    p.end();
    expect(() => p.write('x')).toThrow();
  });

  it('stays linear in the size of the input', () => {
    const row = '<row r="1" spans="1:3"><c r="A1" t="s"><v>12</v></c><c r="B1"><v>3.5</v></c><c r="C1" t="inlineStr"><is><t>text &amp; more</t></is></c></row>\n';
    const doc = `<sheetData>${row.repeat(40_000)}</sheetData>`; // about 5 MB
    const noop: XmlHandler = { start() {}, end() {}, text() {} };
    const t0 = performance.now();
    parseXml(doc, noop, { part: PART, limits: DEFAULT_READ_LIMITS });
    const wholeTime = performance.now() - t0;
    const t1 = performance.now();
    const p = createXmlStreamParser(noop, { part: PART, limits: DEFAULT_READ_LIMITS });
    for (let i = 0; i < doc.length; i += 16_384) p.write(doc.slice(i, i + 16_384));
    p.end();
    const chunkTime = performance.now() - t1;
    expect(chunkTime).toBeLessThan(Math.max(3 * wholeTime, 200));
  });
});

describe('createXmlByteStreamParser', () => {
  function bytesOutcome(bytes: Uint8Array, sizes: number[]): Outcome {
    return capture((h) => {
      const p = createXmlByteStreamParser(h, { part: PART, limits: DEFAULT_READ_LIMITS });
      let at = 0;
      let k = 0;
      while (at < bytes.length) {
        const n = Math.max(1, sizes[k++ % sizes.length] ?? 1);
        p.write(bytes.subarray(at, at + n));
        at += n;
      }
      p.end();
    });
  }

  function decodeThenParse(bytes: Uint8Array): Outcome {
    return capture((h) => parseXml(decodeXml(bytes, PART), h, { part: PART, limits: DEFAULT_READ_LIMITS }));
  }

  const utf16 = (text: string, le: boolean): Uint8Array => {
    const out = new Uint8Array(2 + text.length * 2);
    out[0] = le ? 0xff : 0xfe;
    out[1] = le ? 0xfe : 0xff;
    for (let i = 0; i < text.length; i++) {
      const c = text.charCodeAt(i);
      out[2 + i * 2] = le ? c & 0xff : c >> 8;
      out[3 + i * 2] = le ? c >> 8 : c & 0xff;
    }
    return out;
  };
  const enc = new TextEncoder();
  const cases: [string, Uint8Array][] = [
    ['UTF-8', enc.encode('<?xml version="1.0" encoding="UTF-8"?><a>zażółć 😀</a>')],
    ['UTF-8 with BOM', new Uint8Array([0xef, 0xbb, 0xbf, ...enc.encode('<a>x</a>')])],
    ['UTF-16LE', utf16('<?xml version="1.0" encoding="UTF-16"?><a>zażółć 😀</a>', true)],
    ['UTF-16BE', utf16('<a>x</a>', false)],
    ['wrong declaration', enc.encode('<?xml version="1.0" encoding="ISO-8859-1"?><a/>')],
    ['invalid UTF-8', new Uint8Array([0x3c, 0x61, 0x3e, 0xc3, 0x28, 0x3c, 0x2f, 0x61, 0x3e])],
    ['truncated UTF-8', new Uint8Array([0x3c, 0x61, 0x2f, 0x3e, 0xe2, 0x82])],
    ['one byte', new Uint8Array([0x3c])],
    ['empty', new Uint8Array(0)],
  ];
  for (const [name, bytes] of cases) {
    it(`matches decodeXml + parseXml for ${name}`, () => {
      const expected = decodeThenParse(bytes);
      for (const sizes of [[1], [2], [3], [5, 1], [10_000]]) {
        const got = bytesOutcome(bytes, sizes);
        expect(got.error?.code).toEqual(expected.error?.code);
        expect(got.error?.message).toEqual(expected.error?.message);
        if (expected.error === null) expect(got.events).toEqual(expected.events);
      }
    });
  }
});
