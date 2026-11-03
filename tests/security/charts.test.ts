// Chart parts are untrusted input (CLAUDE.md): the read limits apply to them, a part shown by many
// frames is parsed once, broken parts are warnings and never crashes, and names or values from
// the file never reach object prototypes (D-K01; docs/facts/charts.md CH-29, CH-34).
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError } from '../helpers/expect-error';
import { expectPrototypesUntouched } from '../helpers/prototypes';

function base(): Uint8Array {
  const wb = createWorkbook();
  wb.addSheet('Data').set('A1', 'x').set('B1', 1).addChart({ type: 'column', anchor: 'D2:J12', series: [{ name: 'Data!A1', values: 'Data!B1:B2' }] });
  return writeWorkbook(wb);
}

function patch(bytes: Uint8Array, edit: (name: string, text: string) => string | null): Uint8Array {
  const out: Record<string, Uint8Array> = {};
  for (const [name, data] of Object.entries(unzipSync(bytes))) {
    const t = edit(name, strFromU8(data));
    if (t !== null) out[name] = strToU8(t);
  }
  return zipSync(out);
}

const withChart = (edit: (xml: string) => string): Uint8Array => patch(base(), (name, text) => (name === 'xl/charts/chart1.xml' ? edit(text) : text));

describe('chart parts: limits', () => {
  it('maxXmlTreeElements counts the elements of a chart part', () => {
    const many = withChart((x) => x.replace('<c:layout/>', `<c:layout/>${'<c:junk/>'.repeat(5000)}`));
    expect(readWorkbook(many).getSheet('Data').charts).toHaveLength(1);
    const error = expectXlvError(() => readWorkbook(many, { limits: { maxXmlTreeElements: 2000 } }), 'XLV_E_LIMIT_EXCEEDED');
    expect(error.location.part).toBe('/xl/charts/chart1.xml');
  });

  it('maxXmlDepth applies to chart parts', () => {
    const deep = withChart((x) => x.replace('<c:layout/>', `<c:layout/>${'<c:x>'.repeat(200)}${'</c:x>'.repeat(200)}`));
    expectXlvError(() => readWorkbook(deep), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxXmlDepth' });
  });

  it('maxXmlTextLength applies to a long reference text', () => {
    const long = withChart((x) => x.replace('<c:f>Data!$B$1:$B$2</c:f>', `<c:f>${'A'.repeat(100_000)}</c:f>`));
    expect(readWorkbook(long).getSheet('Data').charts[0]?.series[0]?.values).toHaveLength(100_000);
    expectXlvError(() => readWorkbook(long, { limits: { maxXmlTextLength: 50_000 } }), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxXmlTextLength' });
  });

  it('refuses a DTD in a chart part', () => {
    const dtd = withChart((x) => x.replace('?>\n', '?>\n<!DOCTYPE c:chartSpace [<!ENTITY a "aaaa">]>'));
    expectXlvError(() => readWorkbook(dtd), 'XLV_E_XML_DTD_FORBIDDEN');
  });
});

describe('chart parts: sharing and damage', () => {
  it('parses a chart part shown by many frames once', () => {
    const frames = 3000;
    const shared = patch(base(), (name, text) => {
      if (name === 'xl/charts/chart1.xml') return text.replace('<c:layout/>', `<c:layout/>${'<c:junk a="1"/>'.repeat(20_000)}`);
      if (name !== 'xl/drawings/drawing1.xml') return text;
      const anchor = /<xdr:twoCellAnchor>.*<\/xdr:twoCellAnchor>/.exec(text)![0];
      return text.replace(anchor, anchor.repeat(frames));
    });
    const start = performance.now();
    const charts = readWorkbook(shared).getSheet('Data').charts;
    // Parsing the 40 000-element part for every frame would take minutes.
    expect(performance.now() - start).toBeLessThan(5000);
    expect(charts).toHaveLength(frames);
    expect(charts[1]!.series).toBe(charts[0]!.series);
    expect(charts[frames - 1]!.types).toBe(charts[0]!.types);
  });

  it('gives warnings, not errors, for broken chart parts and skips the chart', () => {
    const cases: [string, (x: string) => string][] = [
      ['not well-formed', (x) => x.replace('</c:chartSpace>', '')],
      ['another root', (x) => x.replace(/c:chartSpace/g, 'c:chartSpaceX')],
      ['no chart', (x) => x.replace(/<c:chart>.*<\/c:chart>/, '')],
      ['no plot area', (x) => x.replace(/<c:plotArea>.*<\/c:plotArea>/, '')],
    ];
    for (const [what, edit] of cases) {
      const wb = readWorkbook(withChart(edit));
      expect(wb.getSheet('Data').charts, what).toEqual([]);
      expect(wb.warnings.map((w) => [w.code, w.location.part]), what).toEqual([['XLV_W_INVALID_VALUE', '/xl/charts/chart1.xml']]);
    }
  });

  it('reads odd but well-formed parts without throwing', () => {
    const odd = withChart((x) =>
      x
        .replace('<c:barDir val="col"/>', '<c:barDir val="constructor"/>')
        .replace('<c:grouping val="clustered"/>', '<c:grouping val="__proto__"/>')
        .replace('<c:legendPos val="r"/>', '<c:legendPos val="__proto__"/>')
        .replace('<c:max', '<c:junk')
        .replace('<c:layout/>', '<c:layout/><c:__proto__><c:ser/></c:__proto__><c:constructor/><c:valAx><c:axId val="toString"/></c:valAx>')
        .replace('<c:axId val="500000002"/></c:barChart>', '<c:axId val="toString"/></c:barChart>')
        .replace('<c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>', '<c:orientation val="minMax"/><c:max val="1e999"/><c:min val="-0x10"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>'),
    );
    const c = readWorkbook(odd).getSheet('Data').charts[0]!;
    expect(c.types).toEqual(['column']);
    expect(c.stacked).toBe(false);
    expect(c.legend).toBe('right');
    expect(c.axes.y).toEqual({ title: null, min: null, max: null });
    expect(Object.getPrototypeOf(c)).toBe(Object.prototype);
    expectPrototypesUntouched();
  });

  it('keeps element and attribute names from the file out of object prototypes', () => {
    const hostile = withChart((x) =>
      x.replace('<c:plotArea>', '<c:plotArea __proto__="x" constructor="y">').replace('<c:ser>', '<c:ser prototype="1"><c:__proto__ val="polluted"/>').replace('<c:tx>', '<c:tx __proto__="1">'),
    );
    readWorkbook(hostile);
    expectPrototypesUntouched();
  });
});
