// Untrusted values given to the writer: prototype keys, inherited properties and getters must
// not change the output or any prototype.
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError } from '../helpers/expect-error';
import { expectPrototypesUntouched } from '../helpers/prototypes';

describe('writer: prototype pollution', () => {
  it('accepts sheet names such as __proto__ and constructor as plain names', () => {
    const wb = createWorkbook();
    for (const name of ['__proto__', 'constructor', 'prototype', 'hasOwnProperty', 'toString']) wb.addSheet(name).set('A1', name);
    const read = readWorkbook(writeWorkbook(wb));
    expect(read.sheets.map((s) => s.cell('A1')?.value)).toEqual(['__proto__', 'constructor', 'prototype', 'hasOwnProperty', 'toString']);
    expectPrototypesUntouched();
  });

  it('rejects __proto__ keys in style objects instead of following them', () => {
    const s = createWorkbook().addSheet('S');
    const style = JSON.parse('{"font": {"bold": true}, "__proto__": {"polluted": "yes"}}') as never;
    expectXlvError(() => s.set('A1', 1, style), 'XLV_E_INVALID_STYLE');
    const font = JSON.parse('{"font": {"__proto__": {"bold": true}}}') as never;
    expectXlvError(() => s.set('A1', 1, font), 'XLV_E_INVALID_STYLE');
    expectPrototypesUntouched();
  });

  it('ignores inherited properties of style and option objects', () => {
    const wb = createWorkbook();
    const s = wb.addSheet('S');
    const inherited = Object.create({ font: { bold: true } }) as object;
    s.set('A1', 1, inherited);
    expect(readWorkbook(writeWorkbook(wb)).getSheet('S').cell('A1')?.style).toEqual({});
    const options = Object.create({ state: 'hidden' }) as object;
    expect(createWorkbook().addSheet('X', options).state).toBe('visible');
  });

  it('reads each style property once, so getters cannot change a validated value', () => {
    const s = createWorkbook().addSheet('S');
    let calls = 0;
    const font = {
      get size(): number {
        calls++;
        return calls === 1 ? 12 : -1;
      },
    };
    s.set('A1', 1, { font });
    expect(calls).toBe(1);
  });

  it('keeps formula text from becoming markup', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', { formula: '"</f><v>1</v>"&A2', result: 'x' });
    expect(readWorkbook(writeWorkbook(wb)).getSheet('S').cell('A1')?.formula?.text).toBe('"</f><v>1</v>"&A2');
  });
});
