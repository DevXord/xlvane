// Style extensions of phase E1 (docs/decisions.md D-E01; docs/facts/sml-write.md W2-60 … W2-63):
// gradient fills, diagonal borders, the remaining font properties, shrinkToFit / justifyLastLine /
// readingOrder, and cell protection - in cell formats, rich text runs and differential formats,
// through the batch writer, the streaming writer, the reader and editCells.
import { strFromU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { editCells } from '../../packages/xlvane/src/edit';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { writeWorkbookStream } from '../../packages/xlvane/src/stream-write';
import { buildXlsx } from '../helpers/xlsx-builder';
import type { CellStyle } from '../../packages/xlvane/src/style';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';

async function collect(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const parts: Uint8Array[] = [];
  for await (const p of stream as unknown as AsyncIterable<Uint8Array>) parts.push(p);
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
}

const STYLE: CellStyle = {
  font: { name: 'Arial', size: 10, bold: true, vertAlign: 'superscript', family: 2, scheme: 'minor', charset: 1, outline: true, shadow: true, condense: true, extend: true },
  fill: { type: 'gradient', gradient: 'linear', degree: 90, stops: [{ position: 0, color: 'FFFFFF' }, { position: 1, color: '4472C4' }] },
  border: { diagonal: { style: 'thin', color: 'FF0000' }, diagonalUp: true, diagonalDown: true, bottom: { style: 'thin' } },
  alignment: { horizontal: 'distributed', shrinkToFit: true, justifyLastLine: true, readingOrder: 'rightToLeft' },
  protection: { locked: false, hidden: true },
};

const stylesXml = (bytes: Uint8Array): string => strFromU8(unzipSync(bytes)['xl/styles.xml']!);

describe('writing the style extensions', () => {
  it('writes every part in schema order with Office’s child order for fonts', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', 'x', STYLE);
    const xml = stylesXml(writeWorkbook(wb));
    expect(xml).toContain(
      '<font><b/><condense/><extend/><outline/><shadow/><vertAlign val="superscript"/><sz val="10"/><name val="Arial"/><family val="2"/><charset val="1"/><scheme val="minor"/></font>',
    );
    expect(xml).toContain('<fill><gradientFill degree="90"><stop position="0"><color rgb="FFFFFFFF"/></stop><stop position="1"><color rgb="FF4472C4"/></stop></gradientFill></fill>');
    expect(xml).toContain('<border diagonalUp="1" diagonalDown="1"><left/><right/><top/><bottom style="thin"></bottom><diagonal style="thin"><color rgb="FFFF0000"/></diagonal></border>');
    expect(xml).toMatch(/applyAlignment="1" applyProtection="1"><alignment horizontal="distributed" justifyLastLine="1" shrinkToFit="1" readingOrder="2"\/><protection locked="0" hidden="1"\/><\/xf>/);
  });

  it('writes a path gradient with its rectangle, and omits default protection', () => {
    const wb = createWorkbook();
    const s = wb.addSheet('S');
    s.set('A1', 1, { fill: { type: 'gradient', gradient: 'path', left: 0.5, right: 0.5, top: 0.5, bottom: 0.5, stops: [{ position: 0, color: 'FFFFFF' }, { position: 1, color: { theme: 4 } }] } });
    s.set('A2', 2, { protection: { locked: true, hidden: false } });
    const xml = stylesXml(writeWorkbook(wb));
    expect(xml).toContain('<gradientFill type="path" left="0.5" right="0.5" top="0.5" bottom="0.5"><stop position="0"><color rgb="FFFFFFFF"/></stop><stop position="1"><color theme="4"/></stop></gradientFill>');
    expect(xml).not.toContain('<protection');
  });

  it('refuses values outside the documented ranges', () => {
    const s = createWorkbook().addSheet('S');
    const bad: unknown[] = [
      { font: { family: 15 } }, // ST_FontFamily: 0 … 14
      { font: { charset: -1 } }, // Part 1 §18.4.1: 0 … 255
      { font: { scheme: 'body' } },
      { fill: { type: 'gradient', gradient: 'angle', stops: [] } },
      { fill: { type: 'gradient', gradient: 'linear', stops: [{ position: Number.NaN, color: 'FFFFFF' }] } },
      { fill: { type: 'gradient', gradient: 'linear', stops: 'none' } },
      { fill: { type: 'pattern', pattern: 'solid' } },
      { border: { diagonalUp: 1 } },
      { alignment: { readingOrder: 'ltr' } },
      { protection: { visible: true } },
    ];
    for (const style of bad) expect(() => s.set('A1', 1, style as never), JSON.stringify(style)).toThrow(/Invalid style/);
  });
});

describe('reading them back', () => {
  it('gives the same style from the batch and the streaming writer', async () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', 'x', STYLE);
    const batch = writeWorkbook(wb);
    const streamed = await collect(writeWorkbookStream({ sheets: [{ name: 'S', rows: [{ values: ['x'], style: STYLE }] }] }));
    for (const bytes of [batch, streamed]) {
      const style = readWorkbook(bytes).getSheet('S').cell('A1')!.style;
      expect(style.font).toEqual({ name: 'Arial', size: 10, bold: true, vertAlign: 'superscript', family: 2, scheme: 'minor', charset: 1, outline: true, shadow: true, condense: true, extend: true });
      expect(style.fill).toEqual({ type: 'gradient', gradient: 'linear', degree: 90, stops: [{ position: 0, color: 'FFFFFFFF' }, { position: 1, color: 'FF4472C4' }] });
      expect(style.border).toEqual({ diagonalUp: true, diagonalDown: true, bottom: { style: 'thin' }, diagonal: { style: 'thin', color: 'FFFF0000' } });
      expect(style.alignment).toEqual({ horizontal: 'distributed', shrinkToFit: true, justifyLastLine: true, readingOrder: 'rightToLeft' });
      expect(style.protection).toEqual({ locked: false, hidden: true });
    }
  });

  it('reads the extended properties in rich text runs, in Office’s run order', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', { richText: [{ text: 'm' }, { text: '2', font: { vertAlign: 'superscript', family: 2, charset: 1, scheme: 'minor', outline: true } }] });
    const bytes = writeWorkbook(wb);
    expect(strFromU8(unzipSync(bytes)['xl/sharedStrings.xml']!)).toContain('<rPr><outline/><vertAlign val="superscript"/><family val="2"/><charset val="1"/><scheme val="minor"/></rPr>');
    expect(readWorkbook(bytes).getSheet('S').cell('A1')!.richText?.[1]?.font).toEqual({ vertAlign: 'superscript', family: 2, charset: 1, scheme: 'minor', outline: true });
  });

  it('reads values written by other applications and ignores invalid ones with a warning', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', 'x', { alignment: { readingOrder: 'leftToRight' }, font: { vertAlign: 'baseline' } });
    const bytes = writeWorkbook(wb);
    expect(readWorkbook(bytes).getSheet('S').cell('A1')!.style).toEqual({ alignment: { readingOrder: 'leftToRight' } });
    const files = unzipSync(bytes);
    const xml = strFromU8(files['xl/styles.xml']!).replace('readingOrder="1"', 'readingOrder="7"');
    files['xl/styles.xml'] = new TextEncoder().encode(xml);
    const read = readWorkbook(zipSync(files));
    expect(read.getSheet('S').cell('A1')!.style).toEqual({});
    expect(read.warnings.some((w) => w.code === 'XLV_W_INVALID_VALUE' && w.message.includes('reading order'))).toBe(true);
  });
});

describe('editCells and differential formats', () => {
  it('editCells merges protection and the new alignment properties into the cell’s format', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', 1, { protection: { locked: false }, alignment: { horizontal: 'left' } });
    const { bytes } = editCells(writeWorkbook(wb), [
      { sheet: 'S', cell: 'A1', value: 2, style: { protection: { locked: true, hidden: true }, alignment: { shrinkToFit: true } } },
    ]);
    const style = readWorkbook(bytes).getSheet('S').cell('A1')!.style;
    expect(style.protection).toEqual({ locked: true, hidden: true });
    expect(style.alignment).toEqual({ horizontal: 'left', shrinkToFit: true });
  });

  it('writes and reads a gradient fill in a differential format', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', 5).conditionalFormat('A1', [{ type: 'cellIs', operator: 'greaterThan', formulas: ['1'], style: { fill: STYLE.fill!, font: { vertAlign: 'subscript' } } }]);
    const bytes = writeWorkbook(wb);
    expect(stylesXml(bytes)).toContain('<dxf><font><vertAlign val="subscript"/></font><fill><gradientFill degree="90">');
    const rule = readWorkbook(bytes).getSheet('S').conditionalFormats[0]!.rules[0]!;
    expect(rule.style?.fill).toEqual({ type: 'gradient', gradient: 'linear', degree: 90, stops: [{ position: 0, color: 'FFFFFFFF' }, { position: 1, color: 'FF4472C4' }] });
    expect(rule.style?.font).toEqual({ vertAlign: 'subscript' });
  });
});

describe('review fixes (D-E06): styles read from a file can be written again', () => {
  it('accepts schema-valid gradients and font families that the old checks refused', () => {
    const s = createWorkbook().addSheet('S');
    for (const style of [
      { font: { family: 10 } },
      { fill: { type: 'gradient', gradient: 'linear', stops: [] } },
      { fill: { type: 'gradient', gradient: 'linear', degree: 400, stops: [{ position: 0, color: 'FFFFFF' }] } },
      { fill: { type: 'gradient', gradient: 'path', left: 2, stops: [{ position: 0, color: 'FFFFFF' }, { position: 2, color: '000000' }] } },
    ]) expect(() => s.set('A1', 1, style as never), JSON.stringify(style)).not.toThrow();
  });

  it('drops a schema-invalid text rotation on read, with a warning, and keeps an indent above 255 (D-1101)', () => {
    const bytes = buildXlsx({
      sheets: [{ name: 'S', data: '<row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1" s="2"><v>2</v></c></row>' }],
      styles:
        '<fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders>' +
        '<cellXfs count="3"><xf/><xf applyAlignment="1"><alignment textRotation="200"/></xf><xf applyAlignment="1"><alignment indent="300" horizontal="left"/></xf></cellXfs>',
    });
    const wb = readWorkbook(bytes);
    const s = wb.getSheet('S');
    expect(s.cell('A1')!.style.alignment).toBeUndefined();
    expect(s.cell('B1')!.style.alignment).toEqual({ horizontal: 'left', indent: 300 });
    expect(wb.warnings.some((w) => w.code === 'XLV_W_INVALID_VALUE')).toBe(true);
  });
});
