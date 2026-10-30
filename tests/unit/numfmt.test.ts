import { describe, expect, it } from 'vitest';
import { BUILTIN_NUMBER_FORMATS, isDateFormat, isDateFormatId } from '../../packages/xlvane/src/sml/numfmt';

// Part 1 §18.8.30, "All Languages" table, transcribed from the spec (PDF p.1786–1787).
const ALL_LANGUAGES: readonly (readonly [number, string])[] = [
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
];

describe('BUILTIN_NUMBER_FORMATS', () => {
  it('contains exactly the "All Languages" implied formats of Part 1 §18.8.30', () => {
    expect([...BUILTIN_NUMBER_FORMATS.entries()]).toEqual(ALL_LANGUAGES);
  });

  it('does not invent formats for ids the spec leaves to the formatCode attribute (5–8 and others)', () => {
    for (const id of [5, 6, 7, 8, 23, 24, 25, 26, 41, 42, 43, 44, 50, 164]) {
      expect(BUILTIN_NUMBER_FORMATS.has(id)).toBe(false);
    }
  });

  it('keeps the space before ";" in ids 37 and 38 as printed in the spec', () => {
    expect(BUILTIN_NUMBER_FORMATS.get(37)).toBe('#,##0 ;(#,##0)');
    expect(BUILTIN_NUMBER_FORMATS.get(38)).toBe('#,##0 ;[Red](#,##0)');
  });

  it('cannot be changed through the exported reference', () => {
    expect(() => (BUILTIN_NUMBER_FORMATS as Map<number, string>).set(14, 'x')).toThrow(TypeError);
    expect(() => (BUILTIN_NUMBER_FORMATS as Map<number, string>).delete(14)).toThrow(TypeError);
    expect(() => (BUILTIN_NUMBER_FORMATS as Map<number, string>).clear()).toThrow(TypeError);
    expect(BUILTIN_NUMBER_FORMATS.get(14)).toBe('mm-dd-yy');
    expect(BUILTIN_NUMBER_FORMATS.size).toBe(ALL_LANGUAGES.length);
  });
});

describe('isDateFormatId', () => {
  const DATE_IDS = new Set([
    // All Languages: 14–22 and 45–47.
    14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47,
    // zh-tw, zh-cn, ja-jp, ko-kr: every code listed for 27–36 and 50–58 is a date or time code.
    27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 50, 51, 52, 53, 54, 55, 56, 57, 58,
    // th-th: 71–81 are the date and time codes (59–62 and 67–70 are numeric).
    71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81,
  ]);

  it('is true for the built-in date and time ids', () => {
    for (const id of DATE_IDS) expect(isDateFormatId(id), String(id)).toBe(true);
  });

  it('is false for every other id from 0 to 200', () => {
    for (let id = 0; id <= 200; id++) {
      if (!DATE_IDS.has(id)) expect(isDateFormatId(id), String(id)).toBe(false);
    }
  });

  it('is false for ids that are not whole numbers', () => {
    for (const id of [-14, 14.5, Number.NaN, Number.POSITIVE_INFINITY]) expect(isDateFormatId(id)).toBe(false);
  });

  it('agrees with isDateFormat for every "All Languages" format', () => {
    for (const [id, code] of BUILTIN_NUMBER_FORMATS) {
      expect(isDateFormat(code), `${id}: ${code}`).toBe(isDateFormatId(id));
    }
  });
});

describe('isDateFormat', () => {
  it('recognises the date tokens y, m, d and the time tokens h, s (Part 1 §18.8.31)', () => {
    for (const code of ['yy', 'yyyy', 'm', 'mm', 'mmm', 'mmmm', 'mmmmm', 'd', 'dd', 'ddd', 'dddd', 'h', 'hh', 's', 'ss']) {
      expect(isDateFormat(code), code).toBe(true);
    }
  });

  it('recognises common date and time codes', () => {
    for (const code of [
      'yyyy-mm-dd',
      'dd/mm/yyyy hh:mm',
      'h:mm:ss.00',
      'h:mm:ss A/P',
      'h AM/PM',
      '[$-409]mmmm d, yyyy',
      '[$-409]h:mm:ss AM/PM',
      '[Red]yyyy-mm-dd',
      '[>=100]d',
      'yyyy\\-mm\\-dd',
      'd" de "mmmm" de "yyyy',
      '\\"d',
      'yyyy"年"m"月"d"日"',
      'h"時"mm"分"',
    ]) {
      expect(isDateFormat(code), code).toBe(true);
    }
  });

  it('recognises elapsed-time brackets [h], [mm], [ss]', () => {
    for (const code of ['[h]:mm', '[mm]:ss', '[ss].00', '[h]', '[hh]', '[m]', '[s]', '[hhh]:mm']) {
      expect(isDateFormat(code), code).toBe(true);
    }
  });

  it('recognises the international era codes e, ee, g, gg, ggg, r and rr', () => {
    for (const code of ['[$-404]e/m/d', '[$-411]ge.m.d', '[$-411]ggge"年"m"月"d"日"', 'e', 'ee', 'g', 'gg', 'ggg', 'r', 'rr']) {
      expect(isDateFormat(code), code).toBe(true);
    }
  });

  it('recognises the system long date and system time language info values 0xf800 and 0xf400', () => {
    for (const code of ['[$-F800]', '[$-f800]', '[$-F400]', '[$-F800]dddd, mmmm dd, yyyy', '[$-F400]h:mm:ss AM/PM']) {
      expect(isDateFormat(code), code).toBe(true);
    }
  });

  it('checks every section of a multi-section code', () => {
    expect(isDateFormat('0.00;yyyy-mm-dd')).toBe(true);
    expect(isDateFormat('[<1]h:mm;0.00')).toBe(true);
    expect(isDateFormat('0;0;d')).toBe(true);
    expect(isDateFormat('0;-0;0;@')).toBe(false);
    expect(isDateFormat('#,##0.00;[Red]-#,##0.00;"zero";"text: "@')).toBe(false);
  });

  it('is false for General and numeric codes', () => {
    for (const code of [
      '',
      'General',
      'general',
      'GENERAL',
      '[$-409]General',
      '[Red]General',
      '0',
      '0.00',
      '#,##0',
      '#,##0.00',
      '0%',
      '0.00%',
      '# ?/?',
      '# ??/??',
      '@',
      '#,##0 ;(#,##0)',
      '#,##0 ;[Red](#,##0)',
      '0 AM/PM',
      't0',
      't#,##0.00',
      'b1',
      'b2',
    ]) {
      expect(isDateFormat(code), code).toBe(false);
    }
  });

  it('treats E+, E-, e+ and e- as exponent codes, not era years', () => {
    for (const code of ['0.00E+00', '##0.0E+0', '0.00E-00', '0.00e+00', '0.0e-0', '#0.0E+0']) {
      expect(isDateFormat(code), code).toBe(false);
    }
  });

  it('ignores quoted text, backslash escapes, _x spacing and *x fill characters', () => {
    for (const code of ['"yyyy"0', '0" days"', '"hh:mm"', '\\d0', '0\\h', '\\y\\m\\d', '_m0', '_(0_)', '*d0', '0*s', '"Date: "@']) {
      expect(isDateFormat(code), code).toBe(false);
    }
  });

  it('ignores colour, condition and locale/currency brackets', () => {
    for (const code of [
      '[Red]0',
      '[Blue]#,##0',
      '[Color10]0',
      '[Red][<=100];[Blue][>100]',
      '[Blue][>100]0',
      '[$-409]#,##0',
      '[$USD-409]#,##0.00',
      '[$€-2] #,##0.00',
      '[$-404]0',
      '[hm]0',
      '[H]0',
      '[]0',
    ]) {
      expect(isDateFormat(code), code).toBe(false);
    }
  });

  it('treats date letters as case-sensitive, as the spec lists them in lowercase', () => {
    for (const code of ['YYYY', 'DD', 'MM', 'HH', 'SS', 'Y-M-D']) {
      expect(isDateFormat(code), code).toBe(false);
    }
  });

  it('treats an unterminated quote or bracket as running to the end of the code', () => {
    expect(isDateFormat('0"yyyy')).toBe(false);
    expect(isDateFormat('[h')).toBe(false);
    expect(isDateFormat('[Red')).toBe(false);
    expect(isDateFormat('h[')).toBe(true);
    expect(isDateFormat('d"')).toBe(true);
    expect(isDateFormat('0\\')).toBe(false);
    expect(isDateFormat('0_')).toBe(false);
    expect(isDateFormat('0*')).toBe(false);
  });

  it('classifies 100 000-character hostile codes in linear time', () => {
    const n = 100_000;
    const hostile: readonly (readonly [string, boolean])[] = [
      ['['.repeat(n), false],
      [']'.repeat(n), false],
      ['"'.repeat(n), false],
      ['\\'.repeat(n), false],
      ['_'.repeat(n), false],
      ['*'.repeat(n), false],
      [';'.repeat(n), false],
      ['0'.repeat(n), false],
      ['[$-'.repeat(n / 4), false],
      ['[]'.repeat(n / 2), false],
      ['[Red]'.repeat(n / 5), false],
      ['[<='.repeat(n / 4), false],
      ['E+'.repeat(n / 2), false],
      ['General'.repeat(n / 10), false],
      ['A/P'.repeat(n / 4), false],
      ['AM/PM'.repeat(n / 5), false],
      [`[${'h'.repeat(n)}]`, true],
      [`[$-${'f'.repeat(n)}]`, false],
      [`[${'h'.repeat(n)}`, false],
      [`${'"x"'.repeat(n / 3)}d`, true],
      [`${'[Red]'.repeat(n / 5)}s`, true],
      [`${'\\d'.repeat(n / 2)}y`, true],
      [`${'[h'.repeat(n / 2)}]`, false],
    ];
    for (const [code, expected] of hostile) {
      const start = performance.now();
      const result = isDateFormat(code);
      const elapsed = performance.now() - start;
      expect(result, code.slice(0, 20)).toBe(expected);
      expect(elapsed, code.slice(0, 20)).toBeLessThan(100);
    }
  });
});
