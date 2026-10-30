// Built-in number formats and date detection for number format codes (ECMA-376 Part 1 §18.8.30,
// §18.8.31). Format codes come from untrusted files: `isDateFormat` is a single forward scan with
// no regular expressions and no backtracking, so it runs in time linear in the code's length.

/** A `Map` whose contents cannot change after construction. */
class FrozenMap<K, V> extends Map<K, V> {
  constructor(entries: readonly (readonly [K, V])[]) {
    super();
    for (const [key, value] of entries) super.set(key, value);
    Object.freeze(this);
  }

  override set(_key: K, _value: V): this {
    throw new TypeError('This map is read-only.');
  }

  override delete(_key: K): boolean {
    throw new TypeError('This map is read-only.');
  }

  override clear(): void {
    throw new TypeError('This map is read-only.');
  }
}

/**
 * Number formats implied by their `numFmtId` without a `numFmt` element: the "All Languages"
 * table of ECMA-376 Part 1 §18.8.30, verbatim (ids 37 and 38 have a space before ";" in the
 * spec). Ids not listed, such as 5–8, have no implied format and need an explicit `numFmt`.
 * The locale-dependent ids (27–36, 50–81) have different codes per language and are not
 * included; see {@link isDateFormatId}.
 */
export const BUILTIN_NUMBER_FORMATS: ReadonlyMap<number, string> = new FrozenMap<number, string>([
  [0, 'General'],
  [1, '0'],
  [2, '0.00'],
  [3, '#,##0'],
  [4, '#,##0.00'],
  [9, '0%'],
  [10, '0.00%'],
  [11, '0.00E+00'],
  [12, '# ?/?'],
  [13, '# ??/??'],
  [14, 'mm-dd-yy'],
  [15, 'd-mmm-yy'],
  [16, 'd-mmm'],
  [17, 'mmm-yy'],
  [18, 'h:mm AM/PM'],
  [19, 'h:mm:ss AM/PM'],
  [20, 'h:mm'],
  [21, 'h:mm:ss'],
  [22, 'm/d/yy h:mm'],
  [37, '#,##0 ;(#,##0)'],
  [38, '#,##0 ;[Red](#,##0)'],
  [39, '#,##0.00;(#,##0.00)'],
  [40, '#,##0.00;[Red](#,##0.00)'],
  [45, 'mm:ss'],
  [46, '[h]:mm:ss'],
  [47, 'mmss.0'],
  [48, '##0.0E+0'],
  [49, '@'],
]);

/** Built-in ids whose implied formats are dates or times (ECMA-376 Part 1 §18.8.30). */
const DATE_FORMAT_IDS: ReadonlySet<number> = new Set([
  // All Languages: 14–22 and 45–47 use the date and time codes of §18.8.31.
  14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47,
  // zh-tw, zh-cn, ja-jp and ko-kr: every code listed for 27–36 and 50–58 is a date or time code.
  27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 50, 51, 52, 53, 54, 55, 56, 57, 58,
  // th-th: 71–81. Their codes use Thai letters that §18.8.31 does not define; they are classed as
  // date/time because they mirror ids 14–22 and 45–47 (79 is "[x]:yy:zz" like 46, 80 ends in ".0"
  // like 47) and 81 is "d/m/bb". 59–62 and 67–70 are numeric ("t0", "t#,##0", …).
  71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81,
]);

/**
 * Returns `true` if `id` is a built-in number format id whose implied format displays a date or
 * time, including the locale-dependent date ids 27–36 and 50–58 (Chinese, Japanese, Korean) and
 * 71–81 (Thai) listed in ECMA-376 Part 1 §18.8.30. Use it for an `xf@numFmtId` that has no
 * `numFmt` element; use {@link isDateFormat} for explicit format codes.
 */
export function isDateFormatId(id: number): boolean {
  return DATE_FORMAT_IDS.has(id);
}

const CODE_QUOTE = 0x22; // "
const CODE_DOLLAR = 0x24; // $
const CODE_ASTERISK = 0x2a; // *
const CODE_PLUS = 0x2b; // +
const CODE_HYPHEN = 0x2d; // -
const CODE_LEFT_BRACKET = 0x5b; // [
const CODE_BACKSLASH = 0x5c; // \
const CODE_UNDERSCORE = 0x5f; // _
const CODE_UPPER_G = 0x47; // G
const CODE_LOWER_D = 0x64; // d
const CODE_LOWER_E = 0x65; // e
const CODE_LOWER_G = 0x67; // g
const CODE_LOWER_H = 0x68; // h
const CODE_LOWER_M = 0x6d; // m
const CODE_LOWER_R = 0x72; // r
const CODE_LOWER_S = 0x73; // s
const CODE_LOWER_Y = 0x79; // y

/** Language-info values meaning "system long date format" and "system time format". */
const LANGUAGE_SYSTEM_LONG_DATE = 0xf800;
const LANGUAGE_SYSTEM_TIME = 0xf400;
const MAX_LANGUAGE_INFO_DIGITS = 8;

function hexValue(code: number): number {
  if (code >= 0x30 && code <= 0x39) return code - 0x30;
  if (code >= 0x41 && code <= 0x46) return code - 0x41 + 10;
  if (code >= 0x61 && code <= 0x66) return code - 0x61 + 10;
  return -1;
}

/** True if the bracket content `code[from..to)` makes the format display a date or time. */
function isDateTimeBracket(code: string, from: number, to: number): boolean {
  if (from >= to) return false;
  const first = code.charCodeAt(from);
  // Elapsed time: [h], [mm], [ss] and longer runs of one letter (§18.8.31 "Dates and times").
  if (first === CODE_LOWER_H || first === CODE_LOWER_M || first === CODE_LOWER_S) {
    for (let i = from + 1; i < to; i++) if (code.charCodeAt(i) !== first) return false;
    return true;
  }
  // [$<currency string>-<language info>]: language info 0xf800 or 0xf400 selects the system long
  // date or system time format (§18.8.31). Only its low 16 bits, the language ID, are compared.
  if (first === CODE_DOLLAR) {
    let hyphen = -1;
    for (let i = to - 1; i > from; i--) {
      if (code.charCodeAt(i) === CODE_HYPHEN) {
        hyphen = i;
        break;
      }
    }
    const digits = to - hyphen - 1;
    if (hyphen === -1 || digits < 1 || digits > MAX_LANGUAGE_INFO_DIGITS) return false;
    let value = 0;
    for (let i = hyphen + 1; i < to; i++) {
      const digit = hexValue(code.charCodeAt(i));
      if (digit < 0) return false;
      value = value * 16 + digit;
    }
    const languageId = value % 0x10000;
    return languageId === LANGUAGE_SYSTEM_LONG_DATE || languageId === LANGUAGE_SYSTEM_TIME;
  }
  // Colours ([Red], [Color7]), conditions ([>=100]) and anything else.
  return false;
}

/** True if `code` has the keyword "General" (any case) at `at`. */
function isGeneralAt(code: string, at: number): boolean {
  return at + 7 <= code.length && code.slice(at, at + 7).toLowerCase() === 'general';
}

/**
 * Returns `true` if the number format code displays a date or time in any of its sections,
 * following the format-code grammar of ECMA-376 Part 1 §18.8.31. Date and time codes are `y`,
 * `m`, `d`, `h`, `s`, the elapsed-time brackets `[h]`, `[mm]`, `[ss]`, the era codes `e`, `g`,
 * `r`, and the language-info values for the system date and time formats (`[$-F800]`,
 * `[$-F400]`); they are lowercase in the spec and matched case-sensitively. Text in double
 * quotes, a character after `\`, `_` or `*`, colour, condition and locale brackets, the keyword
 * `General` and the exponent codes `E+`, `E-`, `e+`, `e-` are not dates. An unterminated quote
 * or bracket runs to the end of the code. Runs in linear time.
 */
export function isDateFormat(code: string): boolean {
  const length = code.length;
  let i = 0;
  while (i < length) {
    const c = code.charCodeAt(i);
    switch (c) {
      case CODE_QUOTE: {
        // Quoted text is never part of the format code lexicon (§18.8.31 "Text and spacing").
        const end = code.indexOf('"', i + 1);
        if (end === -1) return false;
        i = end + 1;
        break;
      }
      case CODE_BACKSLASH: // the next character is literal
      case CODE_UNDERSCORE: // skip the width of the next character
      case CODE_ASTERISK: // repeat the next character to fill the cell
        i += 2;
        break;
      case CODE_LEFT_BRACKET: {
        const end = code.indexOf(']', i + 1);
        if (end === -1) return false;
        if (isDateTimeBracket(code, i + 1, end)) return true;
        i = end + 1;
        break;
      }
      case CODE_UPPER_G:
        i += isGeneralAt(code, i) ? 7 : 1;
        break;
      case CODE_LOWER_G:
        if (isGeneralAt(code, i)) {
          i += 7;
          break;
        }
        return true; // g, gg, ggg: era (§18.8.31 "International Considerations")
      case CODE_LOWER_E: {
        // e+ and e- are exponent codes; any other e is the era year.
        const next = code.charCodeAt(i + 1);
        if (next === CODE_PLUS || next === CODE_HYPHEN) {
          i += 2;
          break;
        }
        return true;
      }
      case CODE_LOWER_Y: // years
      case CODE_LOWER_M: // months, or minutes after h or before s
      case CODE_LOWER_D: // days
      case CODE_LOWER_H: // hours
      case CODE_LOWER_S: // seconds
      case CODE_LOWER_R: // r, rr: era year codes for ja-jp and zh-tw
        return true;
      default:
        i++;
    }
  }
  return false;
}
