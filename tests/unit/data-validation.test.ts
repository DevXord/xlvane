// Data validations (ECMA-376 Part 1 §18.3.1.32, §18.3.1.33, §18.3.1.44, §18.3.1.45, §18.18.18–21;
// [MS-OI29500] §2.1.613, §2.1.614, §2.1.1085; docs/facts/sheet-rules.md SR-01 … SR-06;
// docs/decisions.md D-C01).
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { deleteRows, insertRows } from '../../packages/xlvane/src/rows';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { writeWorkbookStream } from '../../packages/xlvane/src/stream-write';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import type { XlvErrorCode } from '../../packages/xlvane/src/errors';
import { expectXlvError } from '../helpers/expect-error';
import { buildXlsx } from '../helpers/xlsx-builder';

/** Expects an XlvError with `code` whose message, details or hint match `pattern`. */
function expectRefused(fn: () => unknown, code: XlvErrorCode, pattern?: RegExp): void {
  const error = expectXlvError(fn, code);
  if (pattern) expect(`${error.message} ${JSON.stringify(error.details)} ${error.hint}`).toMatch(pattern);
}

const sheetXml = (bytes: Uint8Array, n = 1): string => strFromU8(unzipSync(bytes)[`xl/worksheets/sheet${String(n)}.xml`] as Uint8Array);

describe('reading data validations', () => {
  const file = buildXlsx({
    sheets: [
      {
        name: 'Form',
        data: '<row r="1"><c r="A1"><v>1</v></c></row>',
        after:
          '<dataValidations count="4">' +
          '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" promptTitle="Pick" prompt="One of_x000A_three" sqref="A1:A10 C1"><formula1>"Yes,No,Maybe"</formula1></dataValidation>' +
          '<dataValidation type="whole" operator="greaterThan" errorStyle="warning" error="Too small" sqref="B2"><formula1>5</formula1></dataValidation>' +
          '<dataValidation type="decimal" sqref="B3"><formula1>1.5</formula1><formula2>Limits!$A$1</formula2></dataValidation>' +
          '<dataValidation type="bogus" sqref="B4"><formula1>1</formula1></dataValidation>' +
          '<dataValidation type="custom" sqref="not a range"><formula1>TRUE</formula1></dataValidation>' +
          '</dataValidations>',
      },
    ],
  });

  it('reads every attribute, with schema defaults for the absent ones', () => {
    const wb = readWorkbook(file);
    const dv = wb.getSheet('Form').dataValidations;
    expect(dv).toHaveLength(3);
    expect(dv[0]).toEqual({
      sqref: 'A1:A10 C1',
      type: 'list',
      operator: null,
      formula1: '"Yes,No,Maybe"',
      formula2: null,
      allowBlank: true,
      showDropDown: false,
      showInputMessage: true,
      showErrorMessage: true,
      errorStyle: 'stop',
      imeMode: 'noControl',
      errorTitle: null,
      error: null,
      promptTitle: 'Pick',
      prompt: 'One of\nthree',
    });
    expect(dv[1]).toMatchObject({ type: 'whole', operator: 'greaterThan', errorStyle: 'warning', error: 'Too small', formula1: '5' });
    expect(dv[2]).toMatchObject({ type: 'decimal', operator: 'between', formula1: '1.5', formula2: 'Limits!$A$1' });
    expect(wb.warnings.map((w) => w.code)).toContain('XLV_W_INVALID_VALUE');
  });

  it('sheets without validations have an empty list', () => {
    expect(readWorkbook(buildXlsx({ sheets: [{ name: 'S' }] })).getSheet('S').dataValidations).toEqual([]);
  });

  it('skips a validation whose sqref has more references than Office allows', () => {
    const refs = Array.from({ length: 32_768 }, (_, i) => `A${String(i + 1)}`).join(' ');
    const wb = readWorkbook(buildXlsx({ sheets: [{ name: 'S', after: `<dataValidations count="1"><dataValidation type="whole" sqref="${refs}"><formula1>1</formula1><formula2>2</formula2></dataValidation></dataValidations>` }] }));
    expect(wb.getSheet('S').dataValidations).toEqual([]);
    expect(wb.warnings.some((w) => w.code === 'XLV_W_INVALID_VALUE')).toBe(true);
  });
});

describe('writing data validations', () => {
  it('writes the element after the merges, in schema attribute form, and reads it back', () => {
    const wb = createWorkbook();
    const s = wb.addSheet('Form').set('A1', 'x').merge('C1:D1');
    s.validate('A2:A20', { type: 'list', formula1: '"Yes,No"', allowBlank: true, showErrorMessage: true, error: 'Pick Yes or No' });
    s.validate('B2 B4:B6', { type: 'whole', operator: 'between', formula1: 1, formula2: 10, showInputMessage: true, prompt: '1 to 10' });
    s.validate('E2', { type: 'date', operator: 'greaterThan', formula1: new Date(Date.UTC(2024, 0, 31)) });
    const bytes = writeWorkbook(wb);
    const xml = sheetXml(bytes);
    expect(xml).toMatch(/<\/mergeCells><dataValidations count="3"><dataValidation type="list" allowBlank="1" showErrorMessage="1" error="Pick Yes or No" sqref="A2:A20"><formula1>"Yes,No"<\/formula1><\/dataValidation>/);
    expect(xml).toContain('<dataValidation type="whole" showInputMessage="1" prompt="1 to 10" sqref="B2 B4:B6"><formula1>1</formula1><formula2>10</formula2></dataValidation>');
    expect(xml).toContain('<dataValidation type="date" operator="greaterThan" sqref="E2"><formula1>45322</formula1></dataValidation>');
    const back = readWorkbook(bytes).getSheet('Form').dataValidations;
    expect(back.map((d) => d.sqref)).toEqual(['A2:A20', 'B2 B4:B6', 'E2']);
    expect(back[1]).toMatchObject({ type: 'whole', operator: 'between', formula1: '1', formula2: '10', prompt: '1 to 10', showInputMessage: true });
  });

  it('strips a leading "=" and writes one element per rule; null removes the rules of that range', () => {
    const wb = createWorkbook();
    const s = wb.addSheet('S');
    s.validate('A1', { type: 'custom', formula1: '=ISNUMBER(A1)' }).validate('A1', { type: 'textLength', operator: 'lessThan', formula1: 10 });
    s.validate('B1', { type: 'none', prompt: 'hello', showInputMessage: true }).validate('B1', null);
    const xml = sheetXml(writeWorkbook(wb));
    expect(xml).toContain('<dataValidation type="custom" sqref="A1"><formula1>ISNUMBER(A1)</formula1></dataValidation>');
    expect(xml).toContain('<dataValidations count="2">');
    expect(xml).not.toContain('B1');
  });

  it('enforces the schema enumerations and the formula rules', () => {
    const s = createWorkbook().addSheet('S');
    expectRefused(() => s.validate('A1', { type: 'number' as never, formula1: 1 }), 'XLV_E_INVALID_OPTION', /type/);
    expectRefused(() => s.validate('A1', { type: 'whole', operator: 'bigger' as never, formula1: 1 }), 'XLV_E_INVALID_OPTION', /operator/);
    expectRefused(() => s.validate('A1', { type: 'whole', formula1: 1 }), 'XLV_E_INVALID_OPTION', /formula2/);
    expectRefused(() => s.validate('A1', { type: 'list' }), 'XLV_E_INVALID_OPTION', /formula1/);
    expectRefused(() => s.validate('A1', { type: 'list', operator: 'equal', formula1: '"a"' }), 'XLV_E_INVALID_OPTION', /operator/);
    expectRefused(() => s.validate('A1', { type: 'none', formula1: 1 }), 'XLV_E_INVALID_OPTION', /formula1/);
    expectRefused(() => s.validate('A1', { type: 'list', formula1: '"a"', errorStyle: 'error' as never }), 'XLV_E_INVALID_OPTION', /errorStyle/);
    expectRefused(() => s.validate('A1', { type: 'list', formula1: '"a"', colour: 1 } as never), 'XLV_E_INVALID_OPTION', /colour/);
    expectRefused(() => s.validate('nope', { type: 'list', formula1: '"a"' }), 'XLV_E_INVALID_CELL_REF');
  });

  it('enforces the limits of [MS-OI29500] §2.1.613 and §2.1.1085', () => {
    const s = createWorkbook().addSheet('S');
    const ok = { type: 'list', formula1: '"a"' } as const;
    expectRefused(() => s.validate('A1', { ...ok, prompt: 'x'.repeat(256) }), 'XLV_E_INVALID_OPTION', /255/);
    expectRefused(() => s.validate('A1', { ...ok, error: 'x'.repeat(225) }), 'XLV_E_INVALID_OPTION', /224/);
    expectRefused(() => s.validate('A1', { ...ok, errorTitle: 'x'.repeat(33) }), 'XLV_E_INVALID_OPTION', /32/);
    expectRefused(() => s.validate('A1', { ...ok, promptTitle: 'x'.repeat(33) }), 'XLV_E_INVALID_OPTION', /32/);
    expectRefused(() => s.validate('A1', { type: 'custom', formula1: 'x'.repeat(8193) }), 'XLV_E_INVALID_OPTION', /8192/);
    s.validate('A1', { ...ok, prompt: 'x'.repeat(255), error: 'x'.repeat(224), errorTitle: 'x'.repeat(32), promptTitle: 'x'.repeat(32) });
  });

  it('is written by the streaming writer too', async () => {
    const stream = writeWorkbookStream({ sheets: [{ name: 'S', rows: [['a']], dataValidations: [{ sqref: 'A2:A9', type: 'list', formula1: '"a,b"' }] }] });
    const bytes = new Uint8Array(await new Response(stream).arrayBuffer());
    expect(readWorkbook(bytes).getSheet('S').dataValidations).toMatchObject([{ sqref: 'A2:A9', type: 'list', formula1: '"a,b"' }]);
  });

  it('follows rows inserted and deleted with insertRows / deleteRows', () => {
    const wb = createWorkbook();
    const s = wb.addSheet('S').set('A1', 1);
    s.validate('B2:B5', { type: 'whole', operator: 'greaterThan', formula1: '=A2' });
    const inserted = insertRows(writeWorkbook(wb), { sheet: 'S', at: 3, count: 2 });
    expect(readWorkbook(inserted.bytes).getSheet('S').dataValidations).toMatchObject([{ sqref: 'B2:B7' }]);
    const deleted = deleteRows(inserted.bytes, { sheet: 'S', at: 2, count: 1 });
    expect(readWorkbook(deleted.bytes).getSheet('S').dataValidations).toMatchObject([{ sqref: 'B2:B6' }]);
  });
});
