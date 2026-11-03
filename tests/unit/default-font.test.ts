// Workbook default font: font 0 of the styles part (W2-12). `createWorkbook({ defaultFont })`
// and `writeWorkbookStream({ defaultFont })` write it and complete the name and size of the other
// fonts from it, as `editCells` already does with the font 0 of the file; `Workbook.defaultFont`
// reads it. Without the option the output stays as before (Calibri 11, D-205).
import { readFileSync } from 'node:fs';
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { editCells } from '../../packages/xlvane/src/edit';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { writeWorkbookStream } from '../../packages/xlvane/src/stream-write';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError } from '../helpers/expect-error';
import { buildXlsx } from '../helpers/xlsx-builder';

function part(bytes: Uint8Array, name: string): string {
  const data = unzipSync(bytes)[name];
  if (!data) throw new Error(`missing ${name}`);
  return strFromU8(data);
}

function fonts(bytes: Uint8Array): string {
  return /<fonts[^>]*>.*<\/fonts>/.exec(part(bytes, 'xl/styles.xml'))?.[0] ?? '';
}

async function collect(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

describe('createWorkbook({ defaultFont })', () => {
  it('writes font 0 from it and completes the name and size of other fonts from it', () => {
    const wb = createWorkbook({ defaultFont: { name: 'Arial', size: 10 } });
    wb.addSheet('S').set('A1', 'x', { font: { bold: true } }).set('A2', 'y', { font: { size: 14 } }).set('A3', 'z', { font: { name: 'Courier New' } });
    const bytes = writeWorkbook(wb);
    expect(fonts(bytes)).toBe(
      '<fonts count="4"><font><sz val="10"/><name val="Arial"/></font><font><b/><sz val="10"/><name val="Arial"/></font><font><sz val="14"/><name val="Arial"/></font><font><sz val="10"/><name val="Courier New"/></font></fonts>',
    );
    const read = readWorkbook(bytes);
    expect(read.defaultFont).toEqual({ name: 'Arial', size: 10 });
    expect(read.getSheet('S').cell('A1')?.style.font).toEqual({ name: 'Arial', size: 10, bold: true });
  });

  it('fills a missing name or size of the default font with Calibri 11', () => {
    const wb = createWorkbook({ defaultFont: { name: 'Arial' } });
    wb.addSheet('S').set('A1', 'x', { font: { italic: true } });
    expect(fonts(writeWorkbook(wb))).toBe('<fonts count="2"><font><sz val="11"/><name val="Arial"/></font><font><i/><sz val="11"/><name val="Arial"/></font></fonts>');
    const sized = createWorkbook({ defaultFont: { size: 9 } });
    sized.addSheet('S');
    expect(fonts(writeWorkbook(sized))).toBe('<fonts count="1"><font><sz val="9"/><name val="Calibri"/></font></fonts>');
  });

  it('writes the other properties of the default font in font 0 only', () => {
    const wb = createWorkbook({ defaultFont: { name: 'Arial', size: 10, color: '333333', family: 2 } });
    wb.addSheet('S').set('A1', 'x', { font: { bold: true } });
    const bytes = writeWorkbook(wb);
    expect(fonts(bytes)).toBe('<fonts count="2"><font><sz val="10"/><color rgb="FF333333"/><name val="Arial"/><family val="2"/></font><font><b/><sz val="10"/><name val="Arial"/></font></fonts>');
    expect(readWorkbook(bytes).defaultFont).toEqual({ name: 'Arial', size: 10, color: 'FF333333', family: 2 });
  });

  it('keeps the output unchanged without the option or with an empty font', () => {
    const make = (options?: Parameters<typeof createWorkbook>[0]): Uint8Array => {
      const wb = createWorkbook(options);
      wb.addSheet('S').set('A1', 'x', { font: { bold: true } });
      return writeWorkbook(wb);
    };
    const plain = make();
    expect(fonts(plain)).toBe('<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>');
    expect(make({ defaultFont: {} })).toEqual(plain);
    expect(make({ defaultFont: { name: 'Calibri', size: 11 } })).toEqual(plain);
    expect(readWorkbook(plain).defaultFont).toEqual({ name: 'Calibri', size: 11 });
  });

  it('rejects an invalid default font with XLV_E_INVALID_STYLE', () => {
    expectXlvError(() => createWorkbook({ defaultFont: { size: 0 } }), 'XLV_E_INVALID_STYLE', { property: 'font.size' });
    expectXlvError(() => createWorkbook({ defaultFont: { name: '' } }), 'XLV_E_INVALID_STYLE', { property: 'font.name' });
    expectXlvError(() => createWorkbook({ defaultFont: 'Arial' as never }), 'XLV_E_INVALID_STYLE');
    expectXlvError(() => createWorkbook({ defaultFont: { nmae: 'Arial' } as never }), 'XLV_E_INVALID_STYLE');
  });

  it('the streaming writer takes the same option', async () => {
    const bytes = await collect(writeWorkbookStream({ defaultFont: { name: 'Arial', size: 10 }, sheets: [{ name: 'S', rows: [{ values: ['x'], style: { font: { bold: true } } }] }] }));
    expect(fonts(bytes)).toBe('<fonts count="2"><font><sz val="10"/><name val="Arial"/></font><font><b/><sz val="10"/><name val="Arial"/></font></fonts>');
    expect(readWorkbook(bytes).defaultFont).toEqual({ name: 'Arial', size: 10 });
  });

  it('editCells completes new fonts from the written default font', () => {
    const wb = createWorkbook({ defaultFont: { name: 'Arial', size: 10 } });
    wb.addSheet('S').set('A1', 'x');
    const edited = editCells(writeWorkbook(wb), [{ sheet: 'S', cell: 'B1', value: 1, style: { font: { bold: true } } }]).bytes;
    expect(readWorkbook(edited).getSheet('S').cell('B1')?.style.font).toEqual({ name: 'Arial', size: 10, bold: true });
  });
});

describe('Workbook.defaultFont', () => {
  it('is font 0 of the styles part', () => {
    // 47813.xlsx (Apache POI corpus, Microsoft Macintosh Excel): <font><sz val="10"/><name val="Verdana"/></font>.
    const bytes = new Uint8Array(readFileSync(new URL('../corpus/poi/47813.xlsx', import.meta.url)));
    expect(readWorkbook(bytes).defaultFont).toEqual({ name: 'Verdana', size: 10 });
  });

  it('is null without a styles part or without fonts', () => {
    expect(readWorkbook(buildXlsx({ sheets: [{ name: 'S', data: '' }] })).defaultFont).toBeNull();
    expect(readWorkbook(buildXlsx({ styles: '<cellXfs count="1"><xf numFmtId="0"/></cellXfs>', sheets: [{ name: 'S', data: '' }] })).defaultFont).toBeNull();
  });
});
