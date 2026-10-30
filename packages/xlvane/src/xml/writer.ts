// Helpers for writing XML text. The output is parsed back by our own reader, so every rule here
// is the inverse of a reader rule:
// - `&`, `<`, `>` are written as the predefined entities (XML 1.0 §4.6, X-07); `"` too in
//   attribute values, which are always quoted with `"`.
// - CR is written as `&#13;` because a literal CR would be turned into LF by end-of-line
//   handling (X-04); in attribute values tab, LF and CR are written as character references
//   because attribute-value normalisation turns literal ones into spaces (X-05).
// - Characters outside the XML Char production (F-XML-CHAR) cannot be written at all; in
//   ST_Xstring values they are written as `_xHHHH_` (§22.9.2.19, F-XS-01), and an underscore
//   that would start such an escape is written as `_x005F_` (§22.4.2.4, F-XS-02).

/** Escapes character data for element content. */
export function escapeText(s: string): string {
  return /[&<>\r]/.test(s)
    ? s.replace(/[&<>\r]/g, (c) => (c === '&' ? '&amp;' : c === '<' ? '&lt;' : c === '>' ? '&gt;' : '&#13;'))
    : s;
}

const ATTR_ESCAPES: ReadonlyMap<string, string> = new Map([
  ['&', '&amp;'],
  ['<', '&lt;'],
  ['>', '&gt;'],
  ['"', '&quot;'],
  ['\t', '&#9;'],
  ['\n', '&#10;'],
  ['\r', '&#13;'],
]);

/** Escapes an attribute value that will be enclosed in double quotes. */
export function escapeAttr(s: string): string {
  return /[&<>"\t\n\r]/.test(s) ? s.replace(/[&<>"\t\n\r]/g, (c) => ATTR_ESCAPES.get(c) ?? c) : s;
}

function hex4(code: number): string {
  return code.toString(16).toUpperCase().padStart(4, '0');
}

/**
 * Encodes a string as an ST_Xstring value: characters that XML 1.0 cannot represent become
 * `_xHHHH_`, and an underscore that starts `_x` + four hex digits + `_` becomes `_x005F_`, so
 * `decodeXstring(encodeXstring(s)) === s` for every well-formed string. Lone surrogates must be
 * rejected by the caller first ({@link firstLoneSurrogate}); they are encoded here as
 * escapes, which the reader keeps as text (D-XS-3).
 */
export function encodeXstring(s: string): string {
  if (!/[\x00-\x08\x0B\x0C\x0E-\x1F\uD800-\uDFFF￾￿]|_x[0-9A-Fa-f]{4}_/.test(s)) return s;
  let out = '';
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c >= 0xd800 && c <= 0xdbff) {
      const next = s.charCodeAt(i + 1);
      if (next >= 0xdc00 && next <= 0xdfff) {
        out += s[i]! + s[i + 1]!;
        i++;
        continue;
      }
      out += `_x${hex4(c)}_`;
    } else if (
      (c < 0x20 && c !== 0x09 && c !== 0x0a && c !== 0x0d) ||
      (c >= 0xdc00 && c <= 0xdfff) ||
      c === 0xfffe ||
      c === 0xffff
    ) {
      out += `_x${hex4(c)}_`;
    } else if (c === 0x5f && /^_x[0-9A-Fa-f]{4}_/.test(s.slice(i, i + 7))) {
      out += '_x005F_';
    } else {
      out += s[i]!;
    }
  }
  return out;
}

/** Index of the first unpaired surrogate code unit in `s`, or -1 if `s` is well-formed UTF-16. */
export function firstLoneSurrogate(s: string): number {
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c < 0xd800 || c > 0xdfff) continue;
    if (c <= 0xdbff) {
      const next = s.charCodeAt(i + 1);
      if (next >= 0xdc00 && next <= 0xdfff) {
        i++;
        continue;
      }
    }
    return i;
  }
  return -1;
}
