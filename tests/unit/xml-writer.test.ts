import fc from 'fast-check';
import { describe, expect, it } from 'vitest';
import { decodeXstring } from '../../packages/xlvane/src/sml/xstring';
import { encodeXstring, escapeAttr, escapeText, firstLoneSurrogate } from '../../packages/xlvane/src/xml/writer';
import { parseXmlTree } from '../../packages/xlvane/src/xml/tree';
import { DEFAULT_READ_LIMITS } from '../../packages/xlvane/src/limits';

const ctx = () => ({ part: '/t.xml', limits: DEFAULT_READ_LIMITS });

function roundTripText(s: string): string {
  const root = parseXmlTree(`<a>${escapeText(s)}</a>`, ctx());
  return root.text;
}

function roundTripAttr(s: string): string | undefined {
  return parseXmlTree(`<a v="${escapeAttr(s)}"/>`, ctx()).attrs.get('v');
}

describe('escapeText / escapeAttr', () => {
  it('escapes markup characters', () => {
    expect(escapeText('a<b>&c')).toBe('a&lt;b&gt;&amp;c');
    expect(escapeAttr('"<&>\'')).toBe('&quot;&lt;&amp;&gt;\'');
  });

  it('keeps CR in text and tab, LF, CR in attributes through XML normalisation', () => {
    expect(roundTripText('a\r\nb\rc\n')).toBe('a\r\nb\rc\n');
    expect(roundTripAttr('a\tb\nc\r\nd ')).toBe('a\tb\nc\r\nd ');
  });

  it('round-trips any XML-representable string', () => {
    const isXmlChar = (cp: number): boolean =>
      cp === 9 || cp === 10 || cp === 13 || (cp >= 0x20 && cp <= 0xd7ff) || (cp >= 0xe000 && cp <= 0xfffd) || cp >= 0x10000;
    const xmlChar = fc.string({ unit: 'binary' }).map((s) => [...s].filter((ch) => isXmlChar(ch.codePointAt(0) ?? 0)).join(''));
    fc.assert(
      fc.property(xmlChar, (s) => {
        expect(roundTripText(s)).toBe(s);
        expect(roundTripAttr(s)).toBe(s);
      }),
    );
  });
});

describe('encodeXstring', () => {
  it('escapes characters that XML 1.0 cannot represent', () => {
    expect(encodeXstring('a\u0008b')).toBe('a_x0008_b');
    expect(encodeXstring('\u0000￾￿')).toBe('_x0000__xFFFE__xFFFF_');
    expect(encodeXstring('tab\tlf\ncr\r')).toBe('tab\tlf\ncr\r');
  });

  it('escapes the underscore of text that looks like an escape', () => {
    expect(encodeXstring('_x0041_')).toBe('_x005F_x0041_');
    expect(encodeXstring('_x00zz_ _x41_ _X0041_')).toBe('_x00zz_ _x41_ _X0041_');
  });

  it('is the inverse of decodeXstring for well-formed strings', () => {
    expect(decodeXstring(encodeXstring('a\uD800b'))).toBe('a_xD800_b'); // documented: lone surrogates do not round-trip
    const tricky = fc.constantFrom('_', 'x', '_x', '0', 'A', 'f', '_x0041_', '_x005F_', '\u0001', '￿', '😀', 'ż');
    // Lone surrogates are rejected by the writer before encoding (firstLoneSurrogate), so they are excluded here.
    const str = fc
      .oneof(fc.string({ unit: 'binary' }), fc.array(tricky).map((a) => a.join('')))
      .filter((s) => firstLoneSurrogate(s) === -1);
    fc.assert(
      fc.property(str, (s) => {
        expect(decodeXstring(encodeXstring(s))).toBe(s);
      }),
    );
  });
});

describe('firstLoneSurrogate', () => {
  it('finds unpaired surrogates only', () => {
    expect(firstLoneSurrogate('ok 😀')).toBe(-1);
    expect(firstLoneSurrogate('a\uD800b')).toBe(1);
    expect(firstLoneSurrogate('\uDC00')).toBe(0);
    expect(firstLoneSurrogate('x\uD83D')).toBe(1);
  });
});
