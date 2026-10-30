// Built-in number format ids (docs/decisions.md D-A05): a style's `numFmt` may be the id of a
// built-in "All Languages" format of ECMA-376 Part 1 §18.8.30 (E-41, V-22). It is written as
// xf@numFmtId with no numFmt element; format codes keep D-204 (date codes get a custom id).
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { editCells } from '../../packages/xlvane/src/edit';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { BUILTIN_NUMBER_FORMATS } from '../../packages/xlvane/src/sml/numfmt';
import { writeWorkbookStream } from '../../packages/xlvane/src/stream-write';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError } from '../helpers/expect-error';

function part(bytes: Uint8Array, name: string): string {
  const data = unzipSync(bytes)[name];
  if (!data) throw new Error(`missing ${name}`);
  return strFromU8(data);
}

async function collect(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

const day = new Date(Date.UTC(2024, 0, 31));

describe('built-in number format ids', () => {
  it('writes the id without a numFmt element and reads the built-in code back', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', day, { numFmt: 14 }).set('A2', 0.5, { numFmt: 10 });
    const bytes = writeWorkbook(wb);
    const styles = part(bytes, 'xl/styles.xml');
    expect(styles).not.toContain('<numFmts');
    expect(styles).toContain('<xf numFmtId="14"');
    expect(styles).toContain('<xf numFmtId="10"');
    const s = readWorkbook(bytes).getSheet('S');
    expect(s.cell('A1')?.value).toEqual(day);
    expect(s.cell('A1')?.numFmt).toBe('mm-dd-yy');
    expect(s.cell('A2')?.numFmt).toBe('0.00%');
  });

  it('a Date with a numeric id keeps that id instead of the default date format', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', day, { numFmt: 22 });
    expect(part(writeWorkbook(wb), 'xl/styles.xml')).toContain('<xf numFmtId="22"');
  });

  it('the id and its code string are different styles; the code keeps D-204', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', day, { numFmt: 14 }).set('A2', day, { numFmt: 'mm-dd-yy' });
    const styles = part(writeWorkbook(wb), 'xl/styles.xml');
    expect(styles).toContain('<xf numFmtId="14"');
    expect(styles).toContain('formatCode="mm-dd-yy"');
  });

  it('id 0 means General: no number format', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', 1, { numFmt: 0 });
    const s = readWorkbook(writeWorkbook(wb)).getSheet('S');
    expect(s.cell('A1')?.numFmt).toBeNull();
  });

  it('every listed built-in id is accepted; others throw XLV_E_INVALID_STYLE', () => {
    for (const id of BUILTIN_NUMBER_FORMATS.keys()) expect(() => createWorkbook().addSheet('S').set('A1', 1, { numFmt: id })).not.toThrow();
    for (const id of [5, 8, 23, 27, 50, 81, 164, -1, 1.5, Number.NaN]) {
      const error = expectXlvError(() => createWorkbook().addSheet('S').set('A1', 1, { numFmt: id }), 'XLV_E_INVALID_STYLE', { property: 'numFmt' });
      expect(error.hint).toContain('14');
    }
  });

  it('the streaming writer and editCells write ids the same way', async () => {
    const streamed = await collect(writeWorkbookStream({ sheets: [{ name: 'S', rows: [{ values: [day], style: { numFmt: 14 } }] }] }));
    expect(part(streamed, 'xl/styles.xml')).toContain('numFmtId="14"');
    expect(readWorkbook(streamed).getSheet('S').cell('A1')?.numFmt).toBe('mm-dd-yy');

    const wb = createWorkbook();
    wb.addSheet('S').set('A1', 1);
    const edited = editCells(writeWorkbook(wb), [{ sheet: 'S', cell: 'B1', value: day, style: { numFmt: 14 } }]);
    const styles = part(edited.bytes, 'xl/styles.xml');
    expect(styles).toContain('numFmtId="14"');
    expect(styles).not.toContain('<numFmts');
    const b1 = readWorkbook(edited.bytes).getSheet('S').cell('B1');
    expect(b1?.value).toEqual(day);
    expect(b1?.numFmt).toBe('mm-dd-yy');
  });
});
