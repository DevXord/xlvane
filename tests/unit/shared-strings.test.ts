import { describe, expect, it } from 'vitest';
import { DEFAULT_READ_LIMITS } from '../../packages/xlvane/src/limits';
import { SML_NS } from '../../packages/xlvane/src/sml/namespaces';
import { parseSharedStrings, type PartContext } from '../../packages/xlvane/src/sml/shared-strings';
import { WarningCollector } from '../../packages/xlvane/src/util/warnings';
import { expectXlvError } from '../helpers/expect-error';
import { expectPrototypesUntouched } from '../helpers/prototypes';

const PART = '/xl/sharedStrings.xml';

function context(): PartContext {
  return { part: PART, limits: DEFAULT_READ_LIMITS, warnings: new WarningCollector() };
}

/** Wraps `items` in an `sst` root in the given namespace. */
function sst(items: string, ns: string = SML_NS.strict): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="${ns}">${items}</sst>`;
}

function parse(xml: string): string[] {
  return parseSharedStrings(xml, context());
}

describe('parseSharedStrings', () => {
  it('reads simple string items (first example of Part 1 §18.4)', () => {
    const xml = `<sst xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"
       count="8" uniqueCount="4">
       <si>
          <t>United States</t>
       </si>
       <si>
          <t>Seattle</t>
       </si>
       <si>
          <t>Denver</t>
       </si>
       <si>
          <t>New York</t>
       </si>
   </sst>`;
    expect(parse(xml)).toEqual(['United States', 'Seattle', 'Denver', 'New York']);
  });

  it('concatenates the text of rich text runs (second example of Part 1 §18.4)', () => {
    const xml = `<sst xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main" count="8" uniqueCount="4">
       <si>
          <r>
             <t xml:space="preserve">United </t>
          </r>
          <r>
             <rPr>
                 <sz val="11"/>
                 <color rgb="FFFF0000"/>
                 <rFont val="Calibri"/>
                 <family val="2"/>
                 <scheme val="minor"/>
             </rPr>
             <t>States</t>
          </r>
       </si>
       <si>
          <t>Seattle</t>
       </si>
   </sst>`;
    expect(parse(xml)).toEqual(['United States', 'Seattle']);
  });

  it('accepts the Transitional namespace and prefixed element names', () => {
    expect(parse(sst('<si><t>a</t></si>', SML_NS.transitional))).toEqual(['a']);
    const prefixed = `<x:sst xmlns:x="${SML_NS.transitional}"><x:si><x:r><x:t>b</x:t></x:r></x:si></x:sst>`;
    expect(parse(prefixed)).toEqual(['b']);
  });

  it('returns an empty list for an empty table', () => {
    expect(parse(sst(''))).toEqual([]);
    expect(parse(`<sst xmlns="${SML_NS.strict}"/>`)).toEqual([]);
  });

  it('keeps empty items so that indexes stay aligned', () => {
    expect(parse(sst('<si/><si><t/></si><si><t></t></si><si><r><t/></r></si><si><t>x</t></si>'))).toEqual([
      '',
      '',
      '',
      '',
      'x',
    ]);
  });

  it('preserves white space in t exactly as written, with or without xml:space', () => {
    const items =
      '<si><t xml:space="preserve">  two  spaces  </t></si>' +
      '<si><t>  no attribute\n\tline 2  </t></si>' +
      '<si><t xml:space="default"> default </t></si>' +
      '<si><t> </t></si>';
    expect(parse(sst(items))).toEqual(['  two  spaces  ', '  no attribute\n\tline 2  ', ' default ', ' ']);
  });

  it('does not include phonetic runs (rPh) in the value', () => {
    const items =
      '<si><t>東京</t><rPh sb="0" eb="2"><t>トウキョウ</t></rPh><phoneticPr fontId="1"/></si>' +
      '<si><r><t>東</t></r><r><t>京</t></r><rPh sb="0" eb="1"><t>トウ</t></rPh></si>';
    expect(parse(sst(items))).toEqual(['東京', '東京']);
  });

  it('concatenates a leading t with the runs that follow it (CT_Rst sequence t?, r*)', () => {
    expect(parse(sst('<si><t>a</t><r><t>b</t></r><r><t>c</t></r></si>'))).toEqual(['abc']);
  });

  it('decodes ST_Xstring escapes in each t separately', () => {
    const items =
      '<si><t>a_x000D__x000A_b</t></si>' +
      '<si><r><t>_x005F_x0041_</t></r><r><t>_x0042_</t></r></si>' +
      '<si><r><t>_x00</t></r><r><t>41_</t></r></si>';
    expect(parse(sst(items))).toEqual(['a\r\nb', '_x0041_B', '_x0041_']);
  });

  it('decodes ST_Xstring escapes after expanding XML references, CDATA and comments', () => {
    const items =
      '<si><t>&lt;&amp;&gt; &#x41;</t></si>' +
      '<si><t>a<![CDATA[<b>]]>c</t></si>' +
      '<si><t>_x00<!-- comment -->41_</t></si>' +
      '<si><t>&#95;x0041_</t></si>';
    expect(parse(sst(items))).toEqual(['<&> A', 'a<b>c', 'A', 'A']);
  });

  it('leaves escapes of lone surrogates as text', () => {
    expect(parse(sst('<si><t>_xD800_</t></si>'))).toEqual(['_xD800_']);
  });

  it('ignores text and elements outside t, and elements of other namespaces', () => {
    const items =
      '<si>\n  <t>a</t>\n  <unknown>x<t>y</t></unknown>\n</si>' +
      `<si><o:t xmlns:o="urn:other">x</o:t><t>b</t><o:r xmlns:o="urn:other"><t>x</t></o:r></si>` +
      '<si><r><rPr><b/>x</rPr><t>c</t></r></si>' +
      '<o:si xmlns:o="urn:other"><t>x</t></o:si>' +
      '<other><si><t>x</t></si></other>' +
      '<si><t>d</t><si><t>x</t></si></si>' +
      '<extLst><ext uri="urn:x"><t>x</t></ext></extLst>';
    expect(parse(sst(items))).toEqual(['a', 'b', 'c', 'd']);
  });

  it('ignores child elements inside t but keeps the character data around them', () => {
    expect(parse(sst('<si><t>a<b>x</b>c</t></si>'))).toEqual(['ac']);
  });

  it('does not report warnings for a valid part', () => {
    const ctx = context();
    parseSharedStrings(sst('<si><t>a</t></si><si><r><t>b</t></r></si>'), ctx);
    expect(ctx.warnings.list()).toEqual([]);
  });

  it('rejects a part whose root is not sst in a SpreadsheetML namespace', () => {
    for (const xml of [
      '<sst><si><t>a</t></si></sst>',
      '<sst xmlns="urn:wrong"><si><t>a</t></si></sst>',
      `<styleSheet xmlns="${SML_NS.strict}"/>`,
      `<SST xmlns="${SML_NS.strict}"/>`,
    ]) {
      const error = expectXlvError(() => parse(xml), 'XLV_E_PART_INVALID');
      expect(error.location.part).toBe(PART);
    }
  });

  it('lets XML errors from the parser through', () => {
    expectXlvError(() => parse(sst('<si><t>a</si>')), 'XLV_E_XML_MALFORMED');
    expectXlvError(() => parse(`<!DOCTYPE sst><sst xmlns="${SML_NS.strict}"/>`), 'XLV_E_XML_DTD_FORBIDDEN');
    expectXlvError(() => parse(sst('<si><t>&ext;</t></si>')), 'XLV_E_XML_ENTITY_FORBIDDEN');
  });

  it('applies the XML text length limit to t', () => {
    // 60 is enough for the namespace URI, which is an attribute value and counts too.
    const ctx: PartContext = { part: PART, limits: { ...DEFAULT_READ_LIMITS, maxXmlTextLength: 60 }, warnings: new WarningCollector() };
    expect(parseSharedStrings(sst(`<si><t>${'x'.repeat(60)}</t></si>`), ctx)).toEqual(['x'.repeat(60)]);
    expectXlvError(() => parseSharedStrings(sst(`<si><t>${'x'.repeat(61)}</t></si>`), ctx), 'XLV_E_LIMIT_EXCEEDED', {
      limit: 'maxXmlTextLength',
    });
  });
});

describe('parseSharedStrings with hostile input', () => {
  it('keeps prototype-related strings as plain array values', () => {
    const names = ['__proto__', 'constructor', 'prototype', 'toString', 'hasOwnProperty', '__defineGetter__'];
    const result = parse(sst(names.map((n) => `<si><t>${n}</t></si>`).join('')));
    expect(result).toEqual(names);
    expect(Array.isArray(result)).toBe(true);
    expect(Object.getPrototypeOf(result)).toBe(Array.prototype);
    expect(Object.keys(result)).toEqual(['0', '1', '2', '3', '4', '5']);
    expectPrototypesUntouched();
  });

  it('does not trust count and uniqueCount', () => {
    const xml = `<sst xmlns="${SML_NS.strict}" count="4294967295" uniqueCount="99999999999999999999"><si><t>a</t></si></sst>`;
    expect(parse(xml)).toEqual(['a']);
    expect(parse(`<sst xmlns="${SML_NS.strict}" count="-1" uniqueCount="__proto__"/>`)).toEqual([]);
  });

  it('parses a large table in linear time', () => {
    const n = 200_000;
    const parts: string[] = [];
    for (let i = 0; i < n; i++) parts.push(i % 2 === 0 ? `<si><t>s${i}</t></si>` : `<si><r><t>r</t></r><r><t>${i}</t></r></si>`);
    const started = Date.now();
    const result = parse(sst(parts.join('')));
    expect(Date.now() - started).toBeLessThan(5000);
    expect(result.length).toBe(n);
    expect(result[0]).toBe('s0');
    expect(result[n - 1]).toBe(`r${n - 1}`);
  });

  it('handles a long string made of many runs', () => {
    const runs = '<r><t>ab</t></r>'.repeat(100_000);
    expect(parse(sst(`<si>${runs}</si>`))).toEqual(['ab'.repeat(100_000)]);
  });
});
