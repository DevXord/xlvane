import { describe, expect, it } from 'vitest';
import { DEFAULT_READ_LIMITS } from '../../packages/xlvane/src/limits';
import { SML_NS } from '../../packages/xlvane/src/sml/namespaces';
import type { PartContext } from '../../packages/xlvane/src/sml/shared-strings';
import { EMPTY_STYLES, parseStyles, type Styles } from '../../packages/xlvane/src/sml/styles';
import { WarningCollector } from '../../packages/xlvane/src/util/warnings';
import { expectXlvError } from '../helpers/expect-error';
import { expectPrototypesUntouched } from '../helpers/prototypes';

const PART = '/xl/styles.xml';

function context(): PartContext {
  return { part: PART, limits: DEFAULT_READ_LIMITS, warnings: new WarningCollector() };
}

/** Wraps `content` in a `styleSheet` root in the given namespace. */
function sheet(content: string, ns: string = SML_NS.strict): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="${ns}">${content}</styleSheet>`;
}

function parse(content: string): { styles: Styles; warnings: ReturnType<WarningCollector['list']> } {
  const ctx = context();
  const styles = parseStyles(sheet(content), ctx);
  return { styles, warnings: ctx.warnings.list() };
}

function numFmtIds(styles: Styles): number[] {
  const out: number[] = [];
  for (let i = 0; i < styles.cellFormatCount; i++) out.push(styles.numFmtId(i));
  return out;
}

/** cellXfs with one xf per attribute string. */
function cellXfs(...xfs: string[]): string {
  return `<cellXfs count="${xfs.length}">${xfs.map((a) => `<xf ${a}/>`).join('')}</cellXfs>`;
}

/** cellStyleXfs with one xf per attribute string. */
function styleXfs(...xfs: string[]): string {
  return `<cellStyleXfs count="${xfs.length}">${xfs.map((a) => `<xf ${a}/>`).join('')}</cellStyleXfs>`;
}

describe('EMPTY_STYLES', () => {
  it('has no cell formats and maps every index to General', () => {
    expect(EMPTY_STYLES.cellFormatCount).toBe(0);
    expect(EMPTY_STYLES.numFmtId(0)).toBe(0);
    expect(EMPTY_STYLES.numFmtId(7)).toBe(0);
    expect(EMPTY_STYLES.customFormatCode(0)).toBeUndefined();
    expect(EMPTY_STYLES.customFormatCode(164)).toBeUndefined();
  });
});

describe('parseStyles', () => {
  it('maps cell format indexes to numFmtIds and declared custom format codes', () => {
    const { styles, warnings } = parse(
      '<numFmts count="2"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/><numFmt numFmtId="165" formatCode="0.000"/></numFmts>' +
        styleXfs('numFmtId="0" fontId="0" fillId="0" borderId="0"') +
        cellXfs(
          'numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"',
          'numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"',
          'numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"',
        ),
    );
    expect(styles.cellFormatCount).toBe(3);
    expect(numFmtIds(styles)).toEqual([0, 14, 164]);
    expect(styles.customFormatCode(164)).toBe('yyyy-mm-dd');
    expect(styles.customFormatCode(165)).toBe('0.000');
    // Built-in formats are implied, not declared (§18.8.30): they have no custom code here.
    expect(styles.customFormatCode(14)).toBeUndefined();
    expect(styles.customFormatCode(0)).toBeUndefined();
    expect(warnings).toEqual([]);
  });

  it('reads the Annex L example layout (cellStyleXfs, cellXfs with apply attributes)', () => {
    const { styles } = parse(
      '<numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0.00"/></numFmts>' +
        styleXfs(
          'numFmtId="0" fontId="0" fillId="0" borderId="0"',
          'numFmtId="0" fontId="4" fillId="2" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"',
        ) +
        cellXfs(
          'numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"',
          'numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"',
          'numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyBorder="1"',
          'numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyBorder="1"',
          'numFmtId="0" fontId="4" fillId="2" borderId="2" xfId="1" applyBorder="1"',
        ),
    );
    expect(numFmtIds(styles)).toEqual([0, 0, 4, 164, 0]);
    expect(styles.customFormatCode(164)).toBe('"$"#,##0.00');
  });

  it('accepts the Transitional namespace and prefixed element names', () => {
    const transitional = parseStyles(sheet(cellXfs('numFmtId="14"'), SML_NS.transitional), context());
    expect(transitional.numFmtId(0)).toBe(14);
    const prefixed = parseStyles(
      `<x:styleSheet xmlns:x="${SML_NS.strict}"><x:numFmts><x:numFmt numFmtId="200" formatCode="0%"/></x:numFmts>` +
        '<x:cellXfs><x:xf numFmtId="200"/></x:cellXfs></x:styleSheet>',
      context(),
    );
    expect(prefixed.numFmtId(0)).toBe(200);
    expect(prefixed.customFormatCode(200)).toBe('0%');
  });

  it('counts xf elements, not the count attribute', () => {
    const { styles } = parse('<cellXfs count="999"><xf numFmtId="1"/><xf numFmtId="2"/></cellXfs>');
    expect(styles.cellFormatCount).toBe(2);
  });

  it('works without cellXfs or numFmts', () => {
    const { styles, warnings } = parse('<fonts count="0"/>');
    expect(styles.cellFormatCount).toBe(0);
    expect(styles.numFmtId(0)).toBe(0);
    expect(styles.customFormatCode(164)).toBeUndefined();
    expect(warnings).toEqual([]);
    expect(parseStyles(`<styleSheet xmlns="${SML_NS.strict}"/>`, context()).cellFormatCount).toBe(0);
  });

  it('returns General (0) for indexes outside the cell formats', () => {
    const { styles } = parse(cellXfs('numFmtId="14"', 'numFmtId="22"'));
    for (const index of [2, 3, 1000, -1, 0.5, 1.5, Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY, 2 ** 53, -0]) {
      expect(styles.numFmtId(index), String(index)).toBe(index === -0 ? 14 : 0);
    }
  });

  it('decodes ST_Xstring escapes in formatCode', () => {
    const { styles } = parse('<numFmts><numFmt numFmtId="164" formatCode="0_x0022_x_x0022__x005F_x0041_"/></numFmts>');
    expect(styles.customFormatCode(164)).toBe('0"x"_x0041_');
  });

  it('keeps an empty formatCode', () => {
    const { styles, warnings } = parse('<numFmts><numFmt numFmtId="164" formatCode=""/></numFmts>');
    expect(styles.customFormatCode(164)).toBe('');
    expect(warnings).toEqual([]);
  });

  it('ignores elements of other namespaces and unknown children', () => {
    const { styles } = parse(
      '<numFmts><o:numFmt xmlns:o="urn:other" numFmtId="164" formatCode="x"/><numFmt numFmtId="165" formatCode="y"/></numFmts>' +
        `<cellXfs><xf numFmtId="1"><alignment horizontal="left"/></xf><o:xf xmlns:o="urn:other" numFmtId="2"/><xf numFmtId="3"/></cellXfs>` +
        '<o:cellXfs xmlns:o="urn:other"><xf numFmtId="9"/></o:cellXfs>',
    );
    expect(styles.customFormatCode(164)).toBeUndefined();
    expect(styles.customFormatCode(165)).toBe('y');
    expect(numFmtIds(styles)).toEqual([1, 3]);
  });

  it('rejects a part whose root is not styleSheet in a SpreadsheetML namespace', () => {
    for (const xml of [
      '<styleSheet><cellXfs><xf numFmtId="1"/></cellXfs></styleSheet>',
      '<styleSheet xmlns="urn:wrong"/>',
      `<sst xmlns="${SML_NS.strict}"/>`,
      `<stylesheet xmlns="${SML_NS.strict}"/>`,
    ]) {
      const error = expectXlvError(() => parseStyles(xml, context()), 'XLV_E_PART_INVALID');
      expect(error.location.part).toBe(PART);
    }
  });

  it('lets XML errors from the parser through', () => {
    expectXlvError(() => parseStyles(sheet('<cellXfs>'), context()), 'XLV_E_XML_MALFORMED');
    expectXlvError(() => parseStyles(`<!DOCTYPE x><styleSheet xmlns="${SML_NS.strict}"/>`, context()), 'XLV_E_XML_DTD_FORBIDDEN');
  });

  it('applies the tree element limit', () => {
    const ctx: PartContext = { part: PART, limits: { ...DEFAULT_READ_LIMITS, maxXmlTreeElements: 3 }, warnings: new WarningCollector() };
    expectXlvError(() => parseStyles(sheet(cellXfs('numFmtId="1"', 'numFmtId="2"')), ctx), 'XLV_E_LIMIT_EXCEEDED', {
      limit: 'maxXmlTreeElements',
    });
  });
});

describe('parseStyles: only the cell xf record decides the number format (MS-OI29500, 18.8.10 note b)', () => {
  const styles = styleXfs('numFmtId="0"', 'numFmtId="10"', 'numFmtId="11" applyNumberFormat="0"', 'applyNumberFormat="1"');

  it('uses the cell format numFmtId whatever applyNumberFormat says', () => {
    const { styles: s, warnings } = parse(
      styles +
        cellXfs(
          'numFmtId="14" applyNumberFormat="0"',
          'numFmtId="14" xfId="1" applyNumberFormat="0"',
          'numFmtId="14" xfId="1" applyNumberFormat="false"',
          'numFmtId="14" xfId="1" applyNumberFormat="1"',
          'numFmtId="14" xfId="1" applyNumberFormat="true"',
          'numFmtId="14" xfId="1"',
          'numFmtId="14"',
        ),
    );
    expect(numFmtIds(s)).toEqual([14, 14, 14, 14, 14, 14, 14]);
    expect(warnings).toEqual([]);
  });

  it('does not take the number format from the cell style record (cellStyleXfs)', () => {
    const { styles: s, warnings } = parse(styles + cellXfs('xfId="1"', 'xfId="2"', 'xfId="3"', 'numFmtId="0" xfId="1"', ''));
    expect(numFmtIds(s)).toEqual([0, 0, 0, 0, 0]);
    expect(warnings).toEqual([]);
  });

  it('does not read xfId or applyNumberFormat, so their odd values do not warn', () => {
    const { styles: s, warnings } = parse(
      styles + cellXfs('numFmtId="14" xfId="99"', 'numFmtId="15" xfId="-1" applyNumberFormat="yes"', 'numFmtId="16" xfId="__proto__"'),
    );
    expect(numFmtIds(s)).toEqual([14, 15, 16]);
    expect(warnings).toEqual([]);
  });

  it('works without cellStyleXfs', () => {
    const { styles: s, warnings } = parse(cellXfs('numFmtId="14" xfId="0"', 'numFmtId="0" xfId="5" applyNumberFormat="1"'));
    expect(numFmtIds(s)).toEqual([14, 0]);
    expect(warnings).toEqual([]);
  });
});

describe('parseStyles with hostile input', () => {
  const badIds = ['-1', '-0.5', '4294967296', '99999999999999999999', '1.5', '1e3', '0x10', 'abc', '', '__proto__', 'constructor', 'NaN', 'Infinity', '+-1'];

  it('warns about invalid cellXfs numFmtId values and uses General for them', () => {
    for (const bad of badIds) {
      const { styles, warnings } = parse(styleXfs('numFmtId="10"') + cellXfs(`numFmtId="${bad}" xfId="0"`, `numFmtId="${bad}"`, 'numFmtId="3"'));
      expect(numFmtIds(styles), bad).toEqual([0, 0, 3]);
      expect(warnings, bad).toMatchObject([{ code: 'XLV_W_INVALID_VALUE', count: 2, location: { part: PART } }]);
    }
  });

  it('does not read cellStyleXfs, so invalid values there do not matter', () => {
    const { styles, warnings } = parse(styleXfs('numFmtId="-3"', 'numFmtId="__proto__"') + cellXfs('xfId="0"', 'numFmtId="2" xfId="1"'));
    expect(numFmtIds(styles)).toEqual([0, 2]);
    expect(warnings).toEqual([]);
  });

  it('ignores numFmt elements with an invalid or missing numFmtId and warns', () => {
    for (const bad of badIds) {
      const { styles, warnings } = parse(`<numFmts><numFmt numFmtId="${bad}" formatCode="x"/><numFmt numFmtId="164" formatCode="y"/></numFmts>`);
      expect(styles.customFormatCode(164), bad).toBe('y');
      expect(warnings, bad).toMatchObject([{ code: 'XLV_W_INVALID_VALUE', count: 1 }]);
    }
    const { styles, warnings } = parse('<numFmts><numFmt formatCode="x"/><numFmt numFmtId="164"/></numFmts>');
    expect(styles.customFormatCode(164)).toBeUndefined();
    expect(warnings).toMatchObject([{ code: 'XLV_W_INVALID_VALUE', count: 2 }]);
  });

  it('accepts the whole unsignedInt range, with surrounding white space', () => {
    const { styles, warnings } = parse(
      '<numFmts><numFmt numFmtId="4294967295" formatCode="max"/><numFmt numFmtId=" 0164 " formatCode="padded"/></numFmts>' +
        cellXfs('numFmtId="4294967295"', 'numFmtId="00000000000000000164"'),
    );
    expect(numFmtIds(styles)).toEqual([4294967295, 164]);
    expect(styles.customFormatCode(4294967295)).toBe('max');
    expect(styles.customFormatCode(164)).toBe('padded');
    expect(warnings).toEqual([]);
  });

  it('keeps the first of duplicate numFmt declarations and warns', () => {
    const { styles, warnings } = parse(
      '<numFmts><numFmt numFmtId="164" formatCode="first"/><numFmt numFmtId="164" formatCode="second"/></numFmts>',
    );
    expect(styles.customFormatCode(164)).toBe('first');
    expect(warnings).toMatchObject([{ code: 'XLV_W_INVALID_VALUE', count: 1 }]);
  });

  it('uses only the first numFmts and cellXfs element and warns about the others', () => {
    const { styles, warnings } = parse(
      '<numFmts><numFmt numFmtId="164" formatCode="a"/></numFmts><numFmts><numFmt numFmtId="165" formatCode="b"/></numFmts>' +
        styleXfs('numFmtId="10"') +
        styleXfs('numFmtId="11"') +
        cellXfs('numFmtId="7" xfId="0"') +
        cellXfs('numFmtId="1"', 'numFmtId="2"'),
    );
    expect(numFmtIds(styles)).toEqual([7]);
    expect(styles.customFormatCode(164)).toBe('a');
    expect(styles.customFormatCode(165)).toBeUndefined();
    expect(warnings).toMatchObject([{ code: 'XLV_W_INVALID_VALUE', count: 2 }]);
  });

  it('keeps prototype-related format codes as plain values', () => {
    const codes = ['__proto__', 'constructor', 'prototype', 'toString', 'hasOwnProperty'];
    const { styles } = parse(
      `<numFmts>${codes.map((c, i) => `<numFmt numFmtId="${164 + i}" formatCode="${c}"/>`).join('')}</numFmts>` +
        cellXfs(...codes.map((_, i) => `numFmtId="${164 + i}"`)),
    );
    codes.forEach((c, i) => expect(styles.customFormatCode(164 + i)).toBe(c));
    expect(numFmtIds(styles)).toEqual([164, 165, 166, 167, 168]);
    expect(styles.customFormatCode(Number('__proto__'))).toBeUndefined();
    expectPrototypesUntouched();
  });

  it('answers customFormatCode for odd ids without throwing', () => {
    const { styles } = parse('<numFmts><numFmt numFmtId="164" formatCode="x"/></numFmts>');
    for (const id of [-1, 1.5, Number.NaN, Number.POSITIVE_INFINITY, 2 ** 40]) expect(styles.customFormatCode(id)).toBeUndefined();
  });
});
