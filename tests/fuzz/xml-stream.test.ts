// Property tests: parsing a document in random pieces gives the same events, spans and errors as
// parsing it whole, and hostile inputs stay contained in chunked mode.
import fc from 'fast-check';
import { describe, expect, it } from 'vitest';
import { XlvError } from '../../packages/xlvane/src/errors';
import { DEFAULT_READ_LIMITS } from '../../packages/xlvane/src/limits';
import { createXmlStreamParser, parseXml, type XmlHandler, type XmlLimits } from '../../packages/xlvane/src/xml/parser';
import { NUM_RUNS } from './runs';

const PART = '/xl/fuzz.xml';
type Outcome = { events: unknown[]; error: { code: string; message: string; location: unknown } | null };

function capture(run: (handler: XmlHandler) => void): Outcome {
  const events: unknown[] = [];
  const handler: XmlHandler = {
    start: (ns, local, attrs, span) => events.push(['s', ns, local, [...attrs], span.start, span.end, span.qname, span.empty]),
    end: (ns, local, span) => events.push(['e', ns, local, span.start, span.end]),
    text: (value) => events.push(['t', value]),
  };
  try {
    run(handler);
    return { events, error: null };
  } catch (error) {
    if (!(error instanceof XlvError)) throw error;
    return { events, error: { code: error.code, message: error.message, location: error.location } };
  }
}

const whole = (xml: string, limits: XmlLimits): Outcome => capture((h) => parseXml(xml, h, { part: PART, limits }));

function chunked(xml: string, cuts: readonly number[], limits: XmlLimits): Outcome {
  return capture((h) => {
    const p = createXmlStreamParser(h, { part: PART, limits });
    let last = 0;
    for (const cut of [...new Set(cuts.map((c) => Math.min(c, xml.length)))].sort((a, b) => a - b)) {
      p.write(xml.slice(last, cut));
      last = cut;
    }
    p.write(xml.slice(last));
    p.end();
  });
}

// Only characters that match the Char production: with invalid characters, whole-text parsing
// reports them before any other error (a documented difference, D-402).
const UNITS = [
  '<', '>', '/', '&', ';', '#', 'x', '=', '"', "'", '!', '-', '?', '[', ']', ':', ' ', '\t', '\n', '\r', '\r\n',
  'a', 'b', 'Z', '_', '0', 'é', '中', '\u{1F600}',
  '<!DOCTYPE', '<![CDATA[', ']]>', '<!--', '-->', '--', '<?xml ', '<?pi ', '?>',
  '&amp;', '&lt;', '&#x41;', '&#65;', '&#0;', '&nbsp;', '&#',
  'xmlns:p="u"', 'xmlns="u"', ' p:v="1"', ' v="1"', 'p:', '<a>', '</a>', '<a/>', '<p:b ', '</p:b>', '/>', 'text ', '\r',
];
const body = fc.string({ unit: fc.constantFrom(...UNITS), maxLength: 80 });
const doc = fc
  .tuple(fc.constantFrom('', '<a>', '<?xml version="1.0"?>\r\n<a x="1">', '<p:r xmlns:p="u">', '<a><![CDATA['), body, fc.constantFrom('', '</a>', '</p:r>', ']]></a>'))
  .map(([h, b, t]) => h + b + t);
const limits = fc.oneof(
  fc.constant<XmlLimits>(DEFAULT_READ_LIMITS),
  fc.record({ maxXmlDepth: fc.integer({ min: 1, max: 4 }), maxXmlAttributes: fc.integer({ min: 0, max: 3 }), maxXmlNameLength: fc.integer({ min: 1, max: 6 }), maxXmlTextLength: fc.integer({ min: 0, max: 12 }) }),
);
const cuts = (xml: string) => fc.oneof(fc.array(fc.nat({ max: xml.length }), { maxLength: 12 }), fc.constant(Array.from({ length: xml.length }, (_, i) => i)));

/**
 * Exact equality with the default limits. With tiny text limits, text runs longer than the limit
 * are processed in parts, and when such a run holds several problems the one reported first may
 * differ (D-402); then both must fail.
 */
function compare(a: Outcome, b: Outcome, exact: boolean): void {
  if (exact || a.error === null || b.error === null) {
    expect(b.error).toEqual(a.error);
    if (a.error === null) expect(b.events).toEqual(a.events);
  }
}

describe('chunked XML parsing', () => {
  it('equals whole-text parsing for random documents and random pieces', () => {
    fc.assert(
      fc.property(doc.chain((xml) => fc.tuple(fc.constant(xml), cuts(xml), limits)), ([xml, at, lim]) => {
        compare(whole(xml, lim), chunked(xml, at, lim), lim === DEFAULT_READ_LIMITS);
      }),
      { numRuns: NUM_RUNS },
    );
  });

  it('equals whole-text parsing exactly for documents with at most one problem', () => {
    // Well-formed documents with one optional defect inserted at a random place.
    const defect = fc.constantFrom('', '&nbsp;', ']]>', '<!DOCTYPE x>', '<', '&#0;', '<!--', '<?xml ?>', '</b>');
    const ok = fc.array(fc.constantFrom('<b x="1" y=\'2\'>', 'text', '&amp;', '<![CDATA[c<d]]>', '<!-- k -->', '<?p q?>', '\r\n', '<e/>', '</b>'), { maxLength: 20 });
    fc.assert(
      fc.property(ok, defect, fc.nat(), (parts, bad, where) => {
        // Keep the elements balanced.
        let depth = 0;
        const balanced: string[] = [];
        for (const p of parts) {
          if (p === '</b>') {
            if (depth === 0) continue;
            depth--;
          } else if (p.startsWith('<b')) depth++;
          balanced.push(p);
        }
        while (depth-- > 0) balanced.push('</b>');
        balanced.splice(where % (balanced.length + 1), 0, bad);
        const xml = `<root>${balanced.join('')}</root>`;
        for (const size of [1, 2, 5]) {
          const at = Array.from({ length: Math.ceil(xml.length / size) }, (_, i) => i * size);
          const a = whole(xml, DEFAULT_READ_LIMITS);
          const b = chunked(xml, at, DEFAULT_READ_LIMITS);
          expect(b.error).toEqual(a.error);
          if (a.error === null) expect(b.events).toEqual(a.events);
        }
      }),
      { numRuns: NUM_RUNS },
    );
  });

  it('contains hostile inputs in chunked mode', () => {
    const inputs: [string, string][] = [
      ['<!DOCTYPE lolz [<!ENTITY lol "lol"><!ENTITY lol2 "&lol;&lol;">]><lolz>&lol2;</lolz>', 'XLV_E_XML_DTD_FORBIDDEN'],
      ['<?xml version="1.0"?><!DOCTYPE a SYSTEM "file:///etc/passwd"><a/>', 'XLV_E_XML_DTD_FORBIDDEN'],
      [`${'<a>'.repeat(200)}${'</a>'.repeat(200)}`, 'XLV_E_LIMIT_EXCEEDED'],
      [`<a ${Array.from({ length: 300 }, (_, i) => `a${String(i)}="1"`).join(' ')}/>`, 'XLV_E_LIMIT_EXCEEDED'],
      [`<a>${'x'.repeat(5000)}</a>`, 'XLV_E_LIMIT_EXCEEDED'],
      ['<a>&xxe;</a>', 'XLV_E_XML_ENTITY_FORBIDDEN'],
    ];
    const lim = { ...DEFAULT_READ_LIMITS, maxXmlTextLength: 1000 };
    for (const [xml, code] of inputs) {
      const a = whole(xml, lim);
      expect(a.error?.code).toBe(code);
      for (const size of [1, 3, 64]) {
        const at = Array.from({ length: Math.ceil(xml.length / size) }, (_, i) => i * size);
        expect(chunked(xml, at, lim).error).toEqual(a.error);
      }
    }
  });
});
