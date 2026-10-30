// ST_Xstring escape decoding (ECMA-376 Part 1 §22.9.2.19 ST_Xstring; facts F-XS-01 … F-XS-03 and
// F-XML-CHAR, decisions D-XS-1 … D-XS-5 in the facts file).
//
// §22.9.2.19: characters that XML 1.0 cannot represent are written as `_xHHHH_`, where H is a
// hexadecimal digit of the character's value. §22.4.2.4 (bstr) adds that a literal escape
// sequence is stored by escaping its initial underscore as `_x005F_`; decoding every `_xHHHH_`
// in one left-to-right pass handles that case without a special rule.
//
// Decisions (not spec rules, see the facts file): hex digits are accepted in either case (the
// spec's own examples use `_x000a_` and `_x0008_`); an escape of a lone surrogate (D800–DFFF)
// is not a character (XML 1.0 §2.2, Char) and is left as text, so the result is always
// well-formed UTF-16; an escaped high surrogate directly followed by an escaped low surrogate
// is decoded as the pair.

const UNDERSCORE = 0x5f;
const LOWER_X = 0x78;
/** Length of one escape sequence, `_xHHHH_`. */
const ESCAPE_LENGTH = 7;

function hexValue(c: number): number {
  if (c >= 0x30 && c <= 0x39) return c - 0x30;
  if (c >= 0x41 && c <= 0x46) return c - 0x37;
  if (c >= 0x61 && c <= 0x66) return c - 0x57;
  return -1;
}

/** The code unit escaped at `i` (`_xHHHH_`), or -1 if there is no escape at `i`. */
function escapeAt(s: string, i: number): number {
  if (s.charCodeAt(i) !== UNDERSCORE || s.charCodeAt(i + 1) !== LOWER_X || s.charCodeAt(i + 6) !== UNDERSCORE) {
    return -1;
  }
  let value = 0;
  for (let k = 2; k < 6; k++) {
    const d = hexValue(s.charCodeAt(i + k));
    if (d < 0) return -1;
    value = value * 16 + d;
  }
  return value;
}

/**
 * Decodes the `_xHHHH_` escapes of an ST_Xstring value (ECMA-376 Part 1 §22.9.2.19), for
 * example `"a_x000D__x000A_b"` → `"a\r\nb"` and `"_x005F_x0041_"` → `"_x0041_"`. Sequences that
 * are not an escape, including escapes of lone surrogates, are kept as text. Runs in linear
 * time.
 */
export function decodeXstring(value: string): string {
  let at = value.indexOf('_x');
  if (at === -1) return value;
  let out = '';
  let last = 0;
  while (at !== -1) {
    const unit = escapeAt(value, at);
    let length = 0;
    let decoded = '';
    if (unit >= 0xd800 && unit <= 0xdbff) {
      const low = escapeAt(value, at + ESCAPE_LENGTH);
      if (low >= 0xdc00 && low <= 0xdfff) {
        decoded = String.fromCharCode(unit, low);
        length = 2 * ESCAPE_LENGTH;
      }
    } else if (unit !== -1 && (unit < 0xdc00 || unit > 0xdfff)) {
      decoded = String.fromCharCode(unit);
      length = ESCAPE_LENGTH;
    }
    if (length === 0) {
      // Not an escape: the underscore is text; look for the next candidate after it.
      at = value.indexOf('_x', at + 1);
      continue;
    }
    out += value.slice(last, at) + decoded;
    last = at + length;
    at = value.indexOf('_x', last);
  }
  return out + value.slice(last);
}
