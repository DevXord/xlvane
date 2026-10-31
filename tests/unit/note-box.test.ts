// Settings of the box that shows a note: margins, inset mode, locks and anchoring (Part 4
// §19.1.2.22, §19.2.3.17, §19.4.2.37, §19.4.2.38; [MS-OI29500] §2.1.1811, §2.1.1881, §2.1.1886;
// docs/facts/sml-write.md W2-69; docs/decisions.md D-E07).
import { strFromU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { openWorkbookStream } from '../../packages/xlvane/src/stream-read';
import { writeWorkbookStream } from '../../packages/xlvane/src/stream-write';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError } from '../helpers/expect-error';

const vml = (bytes: Uint8Array): string => strFromU8(unzipSync(bytes)['xl/drawings/vmlDrawing1.vml']!);

describe('note boxes', () => {
  it('writes the defaults as before and reads them back', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', 1).note('B2', 'plain');
    const bytes = writeWorkbook(wb);
    const x = vml(bytes);
    expect(x).toContain('o:insetmode="auto"');
    expect(x).toContain('<x:MoveWithCells/><x:SizeWithCells/>');
    expect(x).not.toContain('inset=');
    expect(x).not.toContain('x:Locked');
    expect(readWorkbook(bytes).getSheet('S').notes[0]!.box).toEqual({ autoMargins: true, margins: null, locked: true, lockText: true, anchor: 'absolute' });
  });

  it('writes margins in millimetres, custom inset mode, locks and anchoring, and reads them back', async () => {
    const box = { margins: { left: 2, top: 1.5, right: 2, bottom: 1 }, locked: false, lockText: false, anchor: 'twoCell' } as const;
    const wb = createWorkbook();
    const s = wb.addSheet('S');
    s.note('C3', { text: 'one', box });
    s.note('D4', { text: 'two', box: { anchor: 'oneCell' } });
    const bytes = writeWorkbook(wb);
    const x = vml(bytes);
    expect(x).toContain('o:insetmode="custom"');
    expect(x).toContain('inset="2mm,1.5mm,2mm,1mm"');
    expect(x).toContain('<x:Locked>False</x:Locked>');
    expect(x).toContain('<x:LockText>False</x:LockText>');
    expect(x).toMatch(/<x:ClientData ObjectType="Note"><x:SizeWithCells\/><x:Anchor>/);
    const notes = readWorkbook(bytes).getSheet('S').notes;
    expect(notes[0]!.box).toEqual({ autoMargins: false, margins: { left: 2, top: 1.5, right: 2, bottom: 1 }, locked: false, lockText: false, anchor: 'twoCell' });
    expect(notes[1]!.box).toEqual({ autoMargins: true, margins: null, locked: true, lockText: true, anchor: 'oneCell' });
    // The streaming writer and reader give the same.
    const streamed = new Uint8Array(await new Response(writeWorkbookStream({ sheets: [{ name: 'S', rows: [[1], [2], { values: [3, null, 'v'], notes: [undefined, undefined, { text: 'one', box }] }] }] })).arrayBuffer());
    expect(vml(streamed)).toContain('inset="2mm,1.5mm,2mm,1mm"');
    const stream = await openWorkbookStream(streamed);
    expect((await stream.notes('S'))[0]!.box?.lockText).toBe(false);
  });

  it('reads margins in other units, blank values as the defaults and the inset mode of the shape', () => {
    const wb = createWorkbook();
    wb.addSheet('S').note('A1', 'x');
    const files = unzipSync(writeWorkbook(wb));
    files['xl/drawings/vmlDrawing1.vml'] = new TextEncoder().encode(
      strFromU8(files['xl/drawings/vmlDrawing1.vml']!)
        .replace('o:insetmode="auto"', '')
        .replace('<v:textbox style="mso-direction-alt:auto">', '<v:textbox style="mso-direction-alt:auto" inset="0.1in,,72pt,1cm">')
        .replace('<x:MoveWithCells/><x:SizeWithCells/>', '<x:Locked></x:Locked><x:LockText>f</x:LockText>'),
    );
    const box = readWorkbook(zipSync(files)).getSheet('S').notes[0]!.box;
    expect(box).toEqual({ autoMargins: false, margins: { left: 2.54, top: 1.27, right: 25.4, bottom: 10 }, locked: true, lockText: false, anchor: 'twoCell' });
  });

  it('warns about margins without a unit and about malformed VML, and keeps the notes', () => {
    const wb = createWorkbook();
    wb.addSheet('S').note('A1', 'x');
    const files = unzipSync(writeWorkbook(wb));
    const original = strFromU8(files['xl/drawings/vmlDrawing1.vml']!);
    files['xl/drawings/vmlDrawing1.vml'] = new TextEncoder().encode(original.replace('<v:textbox style="mso-direction-alt:auto">', '<v:textbox inset="1,2,3,4">'));
    const a = readWorkbook(zipSync(files));
    expect(a.getSheet('S').notes[0]!.box?.margins).toBeNull();
    expect(a.warnings.some((w) => w.code === 'XLV_W_INVALID_VALUE')).toBe(true);
    files['xl/drawings/vmlDrawing1.vml'] = new TextEncoder().encode(original.replace('</xml>', ''));
    const b = readWorkbook(zipSync(files));
    expect(b.getSheet('S').notes[0]!.text).toBe('x');
    expect(b.getSheet('S').notes[0]!.box).toBeUndefined();
  });

  it('refuses invalid box settings', () => {
    const s = createWorkbook().addSheet('S');
    expectXlvError(() => s.note('A1', { text: 'x', box: { margins: { left: -1, top: 0, right: 0, bottom: 0 } } }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s.note('A1', { text: 'x', box: { anchor: 'twoCells' as never } }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s.note('A1', { text: 'x', box: { visible: true } as never }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s.note('A1', { text: 'x', box: { margins: { left: 1, top: 1, right: 1 } } as never }), 'XLV_E_INVALID_OPTION');
  });
});
