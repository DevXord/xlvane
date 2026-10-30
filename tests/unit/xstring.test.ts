import { describe, expect, it } from 'vitest';
import { decodeXstring } from '../../packages/xlvane/src/sml/xstring';

/** True if `s` contains an unpaired UTF-16 surrogate. */
function hasLoneSurrogate(s: string): boolean {
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c >= 0xd800 && c <= 0xdbff) {
      const next = s.charCodeAt(i + 1);
      if (next >= 0xdc00 && next <= 0xdfff) {
        i++;
        continue;
      }
      return true;
    }
    if (c >= 0xdc00 && c <= 0xdfff) return true;
  }
  return false;
}

describe('decodeXstring', () => {
  it('returns text without escapes unchanged', () => {
    expect(decodeXstring('')).toBe('');
    expect(decodeXstring('plain text, x_ and _ alone')).toBe('plain text, x_ and _ alone');
  });

  it('decodes the example from Part 1 §22.9.2.19: _x0008_ is U+0008', () => {
    expect(decodeXstring('a_x0008_b')).toBe('a\u0008b');
  });

  it('accepts lower-case and mixed-case hex digits (spec examples use _x000a_ and _x000d_)', () => {
    expect(decodeXstring('line 1_x000a_line 2')).toBe('line 1\nline 2');
    expect(decodeXstring('a_x000d__x000a_b')).toBe('a\r\nb');
    expect(decodeXstring('_x00aB_')).toBe('«');
  });

  it('decodes characters that XML could represent too (any _xHHHH_ is an escape)', () => {
    expect(decodeXstring('_x0041__x0042_')).toBe('AB');
  });

  it('reads _x005F_ as a literal underscore, so _x005F_x0008_ is the text "_x0008_" (§22.4.2.4)', () => {
    expect(decodeXstring('_x005F_')).toBe('_');
    expect(decodeXstring('_x005F_x0008_')).toBe('_x0008_');
    expect(decodeXstring('_x005f_x005F_')).toBe('_x005F_');
  });

  it('decodes escapes at the start, the end and next to each other', () => {
    expect(decodeXstring('_x0009_tab')).toBe('\ttab');
    expect(decodeXstring('tab_x0009_')).toBe('tab\t');
    expect(decodeXstring('_x0009_')).toBe('\t');
    expect(decodeXstring('_x0001__x0002__x0003_')).toBe('\u0001\u0002\u0003');
    expect(decodeXstring('__x0041_')).toBe('_A');
    expect(decodeXstring('_x_x0041_')).toBe('_xA');
  });

  it('leaves incomplete escapes at the end of the string as text', () => {
    for (const s of ['_', '_x', '_x0', '_x00', '_x004', '_x0041', 'abc_x0041']) {
      expect(decodeXstring(s)).toBe(s);
    }
  });

  it('leaves sequences that are not exactly _x + four hex digits + _ as text', () => {
    for (const s of ['_xZZZZ_', '_x00G1_', '_x 041_', '_x41_', '_x00041_', '_X0041_', '_x-041_', '_x+041_', 'x0041_']) {
      expect(decodeXstring(s)).toBe(s);
    }
    // An invalid escape does not hide a valid one that starts at its last underscore.
    expect(decodeXstring('_xZZZZ_x0041_')).toBe('_xZZZZA');
  });

  it('decodes characters that XML 1.0 forbids: NUL, U+FFFE and U+FFFF', () => {
    expect(decodeXstring('_x0000_')).toBe('\u0000');
    expect(decodeXstring('_xFFFE__xffff_')).toBe('￾￿');
  });

  it('leaves escapes of lone surrogates as text, so the result is always well-formed UTF-16', () => {
    for (const s of ['_xD800_', '_xDBFF_', '_xDC00_', '_xDFFF_', '_xDC00__xD800_', 'a_xD800_b', '_xD800_x0041']) {
      const decoded = decodeXstring(s);
      expect(decoded).toBe(s);
      expect(hasLoneSurrogate(decoded)).toBe(false);
    }
    expect(decodeXstring('_xD800__x0041_')).toBe('_xD800_A');
    expect(decodeXstring('_xD800__xD800_')).toBe('_xD800__xD800_');
  });

  it('decodes an escaped surrogate pair into the character it encodes', () => {
    expect(decodeXstring('_xD83D__xDE00_')).toBe('\u{1F600}');
    expect(decodeXstring('a_xd83d__xde00_b')).toBe('a\u{1F600}b');
  });

  it('passes prototype-related names through as ordinary text', () => {
    expect(decodeXstring('__proto__')).toBe('__proto__');
    expect(decodeXstring('constructor_x005F_')).toBe('constructor_');
  });

  it('runs in linear time on inputs full of near-escapes', () => {
    const n = 400_000;
    const inputs = ['_x'.repeat(n), '_x000'.repeat(n), '_xD800_'.repeat(n), '_x0041_'.repeat(n), '_'.repeat(n * 2)];
    const started = Date.now();
    for (const s of inputs) decodeXstring(s);
    expect(decodeXstring('_x0041_'.repeat(n))).toBe('A'.repeat(n));
    expect(Date.now() - started).toBeLessThan(3000);
  });
});
