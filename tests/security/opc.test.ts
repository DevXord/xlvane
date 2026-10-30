// Hostile OPC packages: prototype keys in names and attributes, targets that climb out of the
// package, and very large relationship lists.
import { describe, expect, it } from 'vitest';
import { DEFAULT_READ_LIMITS, type ReadLimits } from '../../packages/xlvane/src/limits';
import { openPackage, resolveTarget, type OpcPackage } from '../../packages/xlvane/src/opc/package';
import { openZip } from '../../packages/xlvane/src/zip/reader';
import { expectXlvError } from '../helpers/expect-error';
import { expectPrototypesUntouched } from '../helpers/prototypes';
import { buildZip } from '../helpers/raw-zip';

const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const PROTO_NAMES = ['__proto__', 'constructor', 'prototype', 'toString', 'hasOwnProperty', 'valueOf'];

const limits = (overrides: Partial<ReadLimits> = {}): ReadLimits => ({ ...DEFAULT_READ_LIMITS, ...overrides });

function types(...items: string[]): string {
  return `<Types xmlns="${CT_NS}">${items.join('')}</Types>`;
}

function rels(...items: string[]): string {
  return `<Relationships xmlns="${REL_NS}">${items.join('')}</Relationships>`;
}

function open(files: [string, string | Uint8Array][], readLimits: ReadLimits = limits()): OpcPackage {
  return openPackage(openZip(buildZip(files.map(([name, data]) => ({ name, data }))), readLimits), readLimits);
}

describe('prototype pollution', () => {
  it('treats relationship Ids, Types and Targets named like prototype keys as plain data', () => {
    const items = PROTO_NAMES.map((name) => `<Relationship Id="${name}" Type="${name}" Target="${name}"/>`);
    const pkg = open([
      ['[Content_Types].xml', types()],
      ['_rels/.rels', rels(...items)],
      ['xl/_rels/workbook.xml.rels', rels(...items)],
    ]);
    expect(pkg.relationships(null)).toEqual(PROTO_NAMES.map((name) => ({ id: name, type: name, target: `/${name}`, external: false })));
    expect(pkg.relationships('/xl/workbook.xml').map((r) => r.target)).toEqual(PROTO_NAMES.map((name) => `/xl/${name}`));
    expectPrototypesUntouched();
  });

  it('detects duplicate Ids even when the Id is "__proto__"', () => {
    const pkg = open([
      ['[Content_Types].xml', types()],
      ['_rels/.rels', rels('<Relationship Id="__proto__" Type="t" Target="a.xml"/>', '<Relationship Id="__proto__" Type="t" Target="b.xml"/>')],
    ]);
    expectXlvError(() => pkg.relationships(null), 'XLV_E_PART_INVALID');
    expectPrototypesUntouched();
  });

  it('treats content type Extensions and PartNames named like prototype keys as plain data', () => {
    const pkg = open([
      [
        '[Content_Types].xml',
        types(
          '<Default Extension="constructor" ContentType="text/constructor"/>',
          '<Default Extension="__proto__" ContentType="text/proto"/>',
          '<Override PartName="/__proto__" ContentType="text/override-proto"/>',
          '<Override PartName="/constructor" ContentType="text/override-constructor"/>',
        ),
      ],
      ['__proto__', 'p'],
      ['constructor', 'c'],
      ['a.constructor', 'x'],
      ['b.__proto__', 'y'],
    ]);
    expect(pkg.partNames).toEqual(['/__proto__', '/constructor', '/a.constructor', '/b.__proto__']);
    expect(pkg.contentType('/__proto__')).toBe('text/override-proto');
    expect(pkg.contentType('/constructor')).toBe('text/override-constructor');
    expect(pkg.contentType('/a.constructor')).toBe('text/constructor');
    expect(pkg.contentType('/b.__proto__')).toBe('text/proto');
    // Names that exist on Object.prototype but not in the stream must not resolve to anything.
    expect(pkg.contentType('/x.toString')).toBeUndefined();
    expect(pkg.contentType('/x.hasOwnProperty')).toBeUndefined();
    expect(pkg.contentType('/valueOf')).toBeUndefined();
    expect(pkg.hasPart('/toString')).toBe(false);
    expect(pkg.readBytes('/__proto__')).toEqual(new TextEncoder().encode('p'));
    expect(pkg.readBytes('/hasOwnProperty')).toBeUndefined();
    expectPrototypesUntouched();
  });

  it('detects duplicate Defaults for the extension "__proto__"', () => {
    expectXlvError(
      () =>
        open([
          ['[Content_Types].xml', types('<Default Extension="__proto__" ContentType="a/b"/>', '<Default Extension="__PROTO__" ContentType="c/d"/>')],
        ]),
      'XLV_E_PART_INVALID',
    );
    expectPrototypesUntouched();
  });
});

describe('targets outside the package', () => {
  it('collapses "../" above the root to the root, so lookups stay inside the archive', () => {
    const target = '../../../../../../etc/passwd';
    const pkg = open([
      ['[Content_Types].xml', types()],
      ['xl/_rels/workbook.xml.rels', rels(`<Relationship Id="r" Type="t" Target="${target}"/>`)],
    ]);
    const [rel] = pkg.relationships('/xl/workbook.xml');
    expect(rel?.target).toBe('/etc/passwd');
    expect(pkg.hasPart('/etc/passwd')).toBe(false);
    expect(pkg.readBytes('/etc/passwd')).toBeUndefined();
    expect(pkg.readXml('/etc/passwd')).toBeUndefined();
  });

  it('resolves to a part inside the archive when one has the collapsed name', () => {
    const pkg = open([
      ['[Content_Types].xml', types()],
      ['secret.xml', '<s/>'],
      ['xl/_rels/workbook.xml.rels', rels('<Relationship Id="r" Type="t" Target="../../../../secret.xml"/>')],
    ]);
    expect(pkg.relationships('/xl/workbook.xml')[0]?.target).toBe('/secret.xml');
    expect(pkg.readXml('/secret.xml')).toBe('<s/>');
  });

  it('never resolves or reads external targets', () => {
    const pkg = open([
      ['[Content_Types].xml', types()],
      ['etc/passwd', 'inside'],
      [
        '_rels/.rels',
        rels(
          '<Relationship Id="a" Type="t" Target="../../etc/passwd" TargetMode="External"/>',
          '<Relationship Id="b" Type="t" Target="file:///etc/passwd" TargetMode="External"/>',
        ),
      ],
    ]);
    expect(pkg.relationships(null)).toEqual([
      { id: 'a', type: 't', target: '../../etc/passwd', external: true },
      { id: 'b', type: 't', target: 'file:///etc/passwd', external: true },
    ]);
  });

  it('ignores ZIP items with absolute or traversing names', () => {
    const pkg = open([
      ['[Content_Types].xml', types()],
      ['/etc/passwd', 'abs'],
      ['../evil.xml', 'up'],
      ['a/../../evil.xml', 'up'],
      ['C:\\evil.xml', 'win'],
    ]);
    expect(pkg.partNames).toEqual([]);
  });

  it('runs in linear time on very long dot-segment chains', () => {
    const target = `${'a/../'.repeat(200_000)}${'../'.repeat(200_000)}x.xml`;
    const started = performance.now();
    expect(resolveTarget('/xl/workbook.xml', target)).toBe('/x.xml');
    expect(resolveTarget('/xl/workbook.xml', `${'%2E%2E/'.repeat(100_000)}x.xml`)).toBe('/x.xml');
    expect(performance.now() - started).toBeLessThan(2_000);
  });
});

describe('large relationship lists', () => {
  it('reads many relationships within the limits', () => {
    const count = 20_000;
    const items = Array.from({ length: count }, (_, i) => `<Relationship Id="rId${i}" Type="t" Target="worksheets/sheet${i}.xml"/>`);
    const pkg = open([
      ['[Content_Types].xml', types()],
      ['xl/_rels/workbook.xml.rels', rels(...items)],
    ]);
    const started = performance.now();
    const list = pkg.relationships('/xl/workbook.xml');
    expect(list).toHaveLength(count);
    expect(list[count - 1]).toEqual({ id: `rId${count - 1}`, type: 't', target: `/xl/worksheets/sheet${count - 1}.xml`, external: false });
    expect(performance.now() - started).toBeLessThan(2_000);
    expectPrototypesUntouched();
  });

  it('stops at maxXmlTreeElements', () => {
    const items = Array.from({ length: 100 }, (_, i) => `<Relationship Id="r${i}" Type="t" Target="a${i}.xml"/>`);
    const pkg = open(
      [
        ['[Content_Types].xml', types()],
        ['_rels/.rels', rels(...items)],
      ],
      limits({ maxXmlTreeElements: 50 }),
    );
    expectXlvError(() => pkg.relationships(null), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxXmlTreeElements' });
  });

  it('applies the XML limits to the Media Types stream', () => {
    const items = Array.from({ length: 100 }, (_, i) => `<Default Extension="e${i}" ContentType="a/b"/>`);
    expectXlvError(() => open([['[Content_Types].xml', types(...items)]], limits({ maxXmlTreeElements: 50 })), 'XLV_E_LIMIT_EXCEEDED');
  });

  it('refuses the decompressed size of the Media Types stream above maxEntryBytes', () => {
    const big = types(...Array.from({ length: 2_000 }, (_, i) => `<Default Extension="e${i}" ContentType="a/b"/>`));
    expectXlvError(() => open([['[Content_Types].xml', big]], limits({ maxEntryBytes: 1_000 })), 'XLV_E_LIMIT_EXCEEDED');
  });
});
