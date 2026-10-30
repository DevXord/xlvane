// Writer limits and style rules from [MS-OI29500] rev. 25.0 (docs/decisions.md D-1101):
// font size, font name, number format length and count, indent, text rotation, column width,
// record counts of the styles part, sheet names, applyX in cellXfs, bookViews and the order of
// font children.
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { editCells } from '../../packages/xlvane/src/edit';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { writeWorkbookStream } from '../../packages/xlvane/src/stream-write';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError } from '../helpers/expect-error';

function parts(bytes: Uint8Array): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [name, data] of Object.entries(unzipSync(bytes))) out[name] = strFromU8(data);
  return out;
}

async function collect(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

const sheet = () => createWorkbook().addSheet('S');
/** The `i`-th of distinct number format codes that are not built-in formats ("0.00" is built-in id 2). */
const custom = (i: number) => `0.${'0'.repeat(i + 3)}`;

describe('style value ranges ([MS-OI29500])', () => {
  it('font size is 1 to 409.55 points (§2.1.677 c)', () => {
    for (const size of [1, 11, 409.55]) expect(() => sheet().set('A1', 1, { font: { size } })).not.toThrow();
    for (const size of [0.5, 0.99, 409.56, 410]) {
      const error = expectXlvError(() => sheet().set('A1', 1, { font: { size } }), 'XLV_E_INVALID_STYLE', { property: 'font.size' });
      expect(error.hint).toBe('Use a size in points from 1 to 409.55, such as 11.');
    }
  });

  it('font name is 1 to 31 characters (§2.1.711)', () => {
    expect(() => sheet().set('A1', 1, { font: { name: 'x'.repeat(31) } })).not.toThrow();
    const error = expectXlvError(() => sheet().set('A1', 1, { font: { name: 'x'.repeat(32) } }), 'XLV_E_INVALID_STYLE', { property: 'font.name' });
    expect(error.hint).toBe('Use a font name of 1 to 31 characters, such as "Arial".');
  });

  it('a number format code is shorter than 255 characters (§2.1.712 b)', () => {
    expect(() => sheet().set('A1', 1, { numFmt: `0.${'0'.repeat(252)}` })).not.toThrow();
    const error = expectXlvError(() => sheet().set('A1', 1, { numFmt: `0.${'0'.repeat(253)}` }), 'XLV_E_INVALID_STYLE', { property: 'numFmt' });
    expect(error.hint).toContain('1 to 254 characters, such as "#,##0.00"');
  });

  it('indent is 0 to 255 (§2.1.693 b)', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', 'x', { alignment: { horizontal: 'left', indent: 255 } });
    expect(readWorkbook(writeWorkbook(wb)).getSheet('S').cell('A1')?.style.alignment?.indent).toBe(255);
    const error = expectXlvError(() => sheet().set('A1', 1, { alignment: { indent: 256 } }), 'XLV_E_INVALID_STYLE', { property: 'alignment.indent' });
    expect(error.hint).toBe('Use a whole number from 0 to 255, such as 2.');
  });

  it('text rotation is 0 to 180 or 255 (§2.1.693 a)', () => {
    for (const textRotation of [0, 90, 180, 255]) expect(() => sheet().set('A1', 1, { alignment: { textRotation } })).not.toThrow();
    for (const textRotation of [181, 254, 256]) {
      const error = expectXlvError(() => sheet().set('A1', 1, { alignment: { textRotation } }), 'XLV_E_INVALID_STYLE');
      expect(error.hint).toBe('Use a whole number from 0 to 180 or 255, such as 45.');
    }
  });

  it('column width is 0 to 255 (§2.1.598 c) and is written unchanged', () => {
    const wb = createWorkbook();
    wb.addSheet('S').column('A', { width: 8.7109375 }).column('B', { width: 255 }).column('C', { width: 0 });
    const xml = parts(writeWorkbook(wb))['xl/worksheets/sheet1.xml']!;
    expect(xml).toContain('width="8.7109375"');
    expect(xml).toContain('width="255"');
    const error = expectXlvError(() => sheet().column('A', { width: 255.5 }), 'XLV_E_INVALID_VALUE');
    expect(error.hint).toBe('Use a width from 0 to 255 characters.');
  });

  it('sheet names reject U+0003 END OF TEXT (§2.1.578 a)', () => {
    expectXlvError(() => createWorkbook().addSheet('a\u0003b'), 'XLV_E_INVALID_SHEET_NAME');
  });
});

describe('record counts of the styles part ([MS-OI29500])', () => {
  it('at most 206 custom number formats (§2.1.713 g)', () => {
    const fill = (n: number) => {
      const wb = createWorkbook();
      const s = wb.addSheet('S');
      for (let i = 0; i < n; i++) s.set(`A${String(i + 1)}`, 1, { numFmt: custom(i) });
      return wb;
    };
    expect(parts(writeWorkbook(fill(206)))['xl/styles.xml']).toContain('<numFmts count="206">');
    expectXlvError(() => writeWorkbook(fill(207)), 'XLV_E_LIMIT_EXCEEDED', { limit: 'customNumFmts', value: 206 });
  });

  it('at most 65430 cell formats (§2.1.700 a, c) and 65430 fills (§2.1.705 a)', () => {
    const wb = createWorkbook();
    const s = wb.addSheet('S');
    // Each distinct fill colour gives one fill record and one cell format.
    for (let i = 0; i < 65_430; i++) s.set(`A${String(i + 1)}`, 1, { fill: { color: i.toString(16).padStart(6, '0') } });
    const error = expectXlvError(() => writeWorkbook(wb), 'XLV_E_LIMIT_EXCEEDED');
    expect(['fills', 'cellXfs']).toContain(error.details['limit']);
    expect(error.hint).toMatch(/^Use fewer distinct /);
  });

  it('the streaming writer applies the same limits', async () => {
    const rows = function* () {
      for (let i = 0; i < 207; i++) yield { values: [1], styles: [{ numFmt: custom(i) }] };
    };
    const stream = writeWorkbookStream({ sheets: [{ name: 'S', rows: rows() }] });
    await expect(collect(stream)).rejects.toMatchObject({ code: 'XLV_E_LIMIT_EXCEEDED', details: { limit: 'customNumFmts' } });
  });
});

describe('style rules of the writer ([MS-OI29500])', () => {
  const wb = createWorkbook();
  const s = wb.addSheet('S');
  s.set('A1', 'x', { font: { bold: true, italic: true, strike: true, underline: true, size: 12, color: 'FF0000', name: 'Arial' } });
  s.set('A2', 1, { numFmt: '0.000', fill: { color: 'FFFF00' }, border: { top: { style: 'thin' } }, alignment: { wrapText: true } });
  const styles = parts(writeWorkbook(wb))['xl/styles.xml']!;

  it('sets applyX where a cell format differs from the Normal style format (§2.1.721 a–f)', () => {
    const xfs = styles.slice(styles.indexOf('<cellXfs'), styles.indexOf('</cellXfs>'));
    const records = xfs.match(/<xf [^>]*>/g)!;
    expect(records[0]).toBe('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>');
    for (const xf of records.slice(1)) {
      const value = (name: string) => new RegExp(`\\s${name}="([^"]*)"`).exec(xf)?.[1];
      if (value('numFmtId') !== '0') expect(xf).toContain('applyNumberFormat="1"');
      if (value('fontId') !== '0') expect(xf).toContain('applyFont="1"');
      if (value('fillId') !== '0') expect(xf).toContain('applyFill="1"');
      if (value('borderId') !== '0') expect(xf).toContain('applyBorder="1"');
    }
    expect(xfs).toMatch(/applyAlignment="1"><alignment wrapText="1"\/>/);
  });

  it('writes font children in the order b, i, strike, u, sz, color, name (§2.1.706 a, §2.1.674 a)', () => {
    expect(styles).toContain('<font><b/><i/><strike/><u/><sz val="12"/><color rgb="FFFF0000"/><name val="Arial"/></font>');
  });

  it('always writes bookViews/workbookView (§2.1.584 a), also when streaming', async () => {
    expect(parts(writeWorkbook(wb))['xl/workbook.xml']).toMatch(/<bookViews><workbookView[^>]*\/><\/bookViews>/);
    const streamed = parts(await collect(writeWorkbookStream({ sheets: [{ name: 'S', rows: [[1]] }] })));
    expect(streamed['xl/workbook.xml']).toMatch(/<bookViews><workbookView[^>]*\/><\/bookViews>/);
  });
});

describe('editCells and styles outside the limits', () => {
  /** A workbook whose cell A1 has a style outside the Office limits for new styles. */
  function outOfRange(): Uint8Array {
    const wb = createWorkbook();
    wb.addSheet('Data').set('A1', 'x', { font: { name: 'Arial', size: 10 }, alignment: { horizontal: 'left', indent: 3 } });
    const files = unzipSync(writeWorkbook(wb));
    const longName = 'N'.repeat(40);
    const styles = strFromU8(files['xl/styles.xml']!)
      .replace('<sz val="10"/>', '<sz val="0.5"/>')
      .replace('<name val="Arial"/>', `<name val="${longName}"/>`)
      .replace('indent="3"', 'indent="300"');
    files['xl/styles.xml'] = strToU8(styles);
    return zipSync(files);
  }

  it('keeps an existing style that is outside the limits, and checks only the new style', () => {
    const input = outOfRange();
    const before = readWorkbook(input).getSheet('Data').cell('A1')?.style;
    expect(before?.font).toMatchObject({ size: 0.5, name: 'N'.repeat(40) });
    expect(before?.alignment?.indent).toBe(300);
    const edited = editCells(input, [{ sheet: 'Data', cell: 'A1', value: 'y', style: { font: { bold: true } } }]);
    const after = readWorkbook(edited.bytes).getSheet('Data').cell('A1')?.style;
    expect(after?.font).toMatchObject({ bold: true, size: 0.5, name: 'N'.repeat(40) });
    expect(after?.alignment?.indent).toBe(300);
    expectXlvError(() => editCells(input, [{ sheet: 'Data', cell: 'A1', value: 'y', style: { font: { size: 0.5 } } }]), 'XLV_E_INVALID_STYLE');
  });

  it('refuses to add a 207th custom number format', () => {
    const wb = createWorkbook();
    const s = wb.addSheet('Data');
    for (let i = 0; i < 206; i++) s.set(`A${String(i + 1)}`, 1, { numFmt: custom(i) });
    const input = writeWorkbook(wb);
    expect(() => editCells(input, [{ sheet: 'Data', cell: 'B1', value: 1, style: { numFmt: '0.00' } }])).not.toThrow();
    expectXlvError(() => editCells(input, [{ sheet: 'Data', cell: 'B1', value: 1, style: { numFmt: '#,##0.0000_)' } }]), 'XLV_E_LIMIT_EXCEEDED', {
      limit: 'customNumFmts',
    });
  });
});
