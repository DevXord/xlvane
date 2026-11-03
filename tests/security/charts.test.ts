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
  // A chart is optional content: a chart part over a limit of its own XML is skipped with
  // XLV_W_LIMIT_REACHED, and the rest of the workbook is read (D-K01).
  const skipped = (bytes: Uint8Array, limits: Record<string, number>, limit: string): void => {
    const wb = readWorkbook(bytes, { limits });
    expect(wb.getSheet('Data').charts).toEqual([]);
    expect(wb.getSheet('Data').cell('B1')?.value).toBe(1);
    expect(wb.warnings.map((w) => [w.code, w.location.part, w.location.sheet])).toEqual([['XLV_W_LIMIT_REACHED', '/xl/charts/chart1.xml', 'Data']]);
    expect(wb.warnings[0]!.message).toContain(limit);
  };

  it('maxXmlTreeElements counts the elements of a chart part; over it the chart is skipped with a warning', () => {
    const many = withChart((x) => x.replace('<c:layout/>', `<c:layout/>${'<c:junk/>'.repeat(5000)}`));
    expect(readWorkbook(many).getSheet('Data').charts).toHaveLength(1);
    skipped(many, { maxXmlTreeElements: 2000 }, 'maxXmlTreeElements');
  });

  it('does not count cached and literal values against maxXmlTreeElements', () => {
    const pts = (n: number): string => Array.from({ length: n }, (_, i) => `<c:pt idx="${String(i)}"><c:v>${String(i)}</c:v></c:pt>`).join('');
    const cached = withChart((x) =>
      x
        .replace(/<c:numCache>.*?<\/c:numCache>/, `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="5000"/>${pts(5000)}<c:extLst/></c:numCache>`)
        .replace(/<c:strCache>.*?<\/c:strCache>/, `<c:strCache><c:ptCount val="5000"/>${pts(5000)}</c:strCache>`)
        .replace('<c:layout/>', `<c:layout/><c:extLst>${'<c:ext/>'.repeat(5000)}</c:extLst>`)
        .replace('</c:barChart>', `<c:ser><c:idx val="1"/><c:order val="1"/><c:cat><c:strLit><c:ptCount val="5000"/>${pts(5000)}</c:strLit></c:cat><c:val><c:numLit><c:ptCount val="5000"/>${pts(5000)}</c:numLit></c:val></c:ser></c:barChart>`),
    );
    const wb = readWorkbook(cached, { limits: { maxXmlTreeElements: 2000 } });
    expect(wb.warnings).toEqual([]);
    expect(wb.getSheet('Data').charts[0]?.series).toEqual([
      { type: 'column', name: { ref: 'Data!$A$1' }, categories: null, values: 'Data!$B$1:$B$2' },
      { type: 'column', name: null, categories: null, values: null },
    ]);
  });

  it('skips a chart part deeper than maxXmlDepth with a warning', () => {
    const deep = withChart((x) => x.replace('<c:layout/>', `<c:layout/>${'<c:x>'.repeat(200)}${'</c:x>'.repeat(200)}`));
    skipped(deep, {}, 'maxXmlDepth');
  });

  it('skips a chart part with a text longer than maxXmlTextLength with a warning', () => {
    const long = withChart((x) => x.replace('<c:f>Data!$B$1:$B$2</c:f>', `<c:f>${'A'.repeat(100_000)}</c:f>`));
    expect(readWorkbook(long).getSheet('Data').charts[0]?.series[0]?.values).toHaveLength(100_000);
    skipped(long, { maxXmlTextLength: 50_000 }, 'maxXmlTextLength');
  });

  it('applies maxXmlDepth and maxXmlTextLength inside caches too', () => {
    const deep = withChart((x) => x.replace('<c:ptCount val="1"/>', `<c:ptCount val="1"/>${'<c:x>'.repeat(200)}${'</c:x>'.repeat(200)}`));
    skipped(deep, {}, 'maxXmlDepth');
    const long = withChart((x) => x.replace('<c:v>x</c:v>', `<c:v>${'A'.repeat(100_000)}</c:v>`));
    skipped(long, { maxXmlTextLength: 50_000 }, 'maxXmlTextLength');
  });

  it('throws for the global ZIP budgets as for every part', () => {
    const big = withChart((x) => x.replace('<c:layout/>', `<c:layout/>${' '.repeat(2_000_000)}`));
    const entry = expectXlvError(() => readWorkbook(big, { limits: { maxEntryBytes: 1_000_000 } }), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxEntryBytes' });
    expect(entry.location.part).toContain('chart1.xml');
    const total = expectXlvError(() => readWorkbook(big, { limits: { maxTotalBytes: 1_000_000 } }), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxTotalBytes' });
    expect(total.location.part).toContain('chart1.xml');
    const ratio = expectXlvError(() => readWorkbook(big, { limits: { maxCompressionRatio: 100 } }), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxCompressionRatio' });
    expect(ratio.location.part).toContain('chart1.xml');
  });

  it('refuses a DTD in a chart part', () => {
    const dtd = withChart((x) => x.replace('?>\n', '?>\n<!DOCTYPE c:chartSpace [<!ENTITY a "aaaa">]>'));
    expectXlvError(() => readWorkbook(dtd), 'XLV_E_XML_DTD_FORBIDDEN');
  });

  it('parses a chart part with a million cached points in linear time', () => {
    const n = 1_000_000;
    const pts = Array.from({ length: n }, (_, i) => `<c:pt idx="${String(i)}"><c:v>${String(i)}</c:v></c:pt>`).join('');
    const huge = withChart((x) => x.replace(/<c:numCache>.*?<\/c:numCache>/, `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${String(n)}"/>${pts}</c:numCache>`));
    const start = performance.now();
    const wb = readWorkbook(huge);
    expect(performance.now() - start).toBeLessThan(10_000);
    expect(wb.warnings).toEqual([]);
    expect(wb.getSheet('Data').charts[0]?.series[0]?.values).toBe('Data!$B$1:$B$2');
  }, 60_000);
});

describe('chart parts: targets and encodings', () => {
  const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0x0d, 0x49, 0x48, 0x44, 0x52, 0xff, 0xfe, 0xc3]);
  const withImage = (): Uint8Array => {
    const wb = createWorkbook();
    const img = wb.addImage(PNG);
    wb.addSheet('Data').set('A1', 'x').set('B1', 1).addImage(img, 'A5:B6').addChart({ type: 'column', anchor: 'D2:J12', series: [{ name: 'Data!A1', values: 'Data!B1:B2' }] });
    return writeWorkbook(wb);
  };
  const raw = (bytes: Uint8Array, name: string, data: Uint8Array): Uint8Array => zipSync({ ...unzipSync(bytes), [name]: data });
  const read = (bytes: Uint8Array): [number, string[][]] => {
    const wb = readWorkbook(bytes);
    return [wb.getSheet('Data').charts.length, wb.warnings.map((w) => [w.code, w.location.part ?? ''])];
  };

  it('skips a chart relationship whose target is not a chart part (CH-01) with XLV_W_INVALID_VALUE', () => {
    const retarget = (target: string): Uint8Array => patch(withImage(), (n, x) => (n === 'xl/drawings/_rels/drawing1.xml.rels' ? x.replace('../charts/chart1.xml"', `${target}"`) : x));
    const png = retarget('../media/image1.png');
    expect(read(png)).toEqual([0, [['XLV_W_INVALID_VALUE', '/xl/media/image1.png']]]);
    expect(readWorkbook(png).warnings[0]!.message).toContain('image/png');
    expect(readWorkbook(png).getSheet('Data').images).toHaveLength(1);
    expect(read(retarget('../worksheets/sheet1.xml'))).toEqual([0, [['XLV_W_INVALID_VALUE', '/xl/worksheets/sheet1.xml']]]);
    expect(read(retarget('/xl/styles.xml'))).toEqual([0, [['XLV_W_INVALID_VALUE', '/xl/styles.xml']]]);
    expect(read(retarget('../charts/missing.xml'))).toEqual([0, [['XLV_W_PART_MISSING', '/xl/charts/missing.xml']]]);
  });

  it('skips a chart part that is not valid UTF-8 or declares another encoding with XLV_W_INVALID_VALUE', () => {
    const invalid = raw(base(), 'xl/charts/chart1.xml', new Uint8Array([...strToU8('<?xml version="1.0"?><c:chartSpace xmlns:c="x">'), 0xff, 0xfe, 0x80]));
    expect(read(invalid)).toEqual([0, [['XLV_W_INVALID_VALUE', '/xl/charts/chart1.xml']]]);
    const latin1 = withChart((x) => x.replace('encoding="UTF-8"', 'encoding="ISO-8859-1"'));
    expect(read(latin1)).toEqual([0, [['XLV_W_INVALID_VALUE', '/xl/charts/chart1.xml']]]);
  });

  it('throws XLV_E_ZIP_CORRUPT for a damaged chart entry, as for every part', () => {
    const stored = zipSync(unzipSync(base()), { level: 0 });
    const at = strFromU8(stored, true).indexOf('<c:plotVisOnly');
    const damaged = stored.slice();
    damaged[at + 3] = 'X'.charCodeAt(0);
    expectXlvError(() => readWorkbook(damaged), 'XLV_E_ZIP_CORRUPT');
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
    expect(performance.now() - start).toBeLessThan(20000);
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
