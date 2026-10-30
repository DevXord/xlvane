// Ported from exceljs v4.4.0 spec/unit/doc/cell.spec.js (MIT, Copyright (c) 2014-2019 Guyon Roche),
// changes: Mocha/Chai → Vitest; the internal sheet mock and `new Cell(...)` are replaced by a
// real worksheet; `cell.merge`/`unmerge` (internal) become `ws.mergeCells`/`ws.unMergeCells`;
// tests of internals (`_mergeCount`, `_upgradeToHyperlink`, constructor validation, `html`)
// are skipped; the note tests
// check `cell.note` instead of the internal `cell.model.comment`, and the box defaults that
// exceljs merges into the model (margins, protection, editAs) are not added (D-A14).
import { beforeEach, describe, expect, it } from 'vitest';
import { ExcelJS, styles } from './helpers';

let ws: any;

describe('Cell', () => {
  beforeEach(() => {
    ws = new ExcelJS.Workbook().addWorksheet('sheet');
  });

  it('stores values', async () => {
    const a1 = ws.getCell('A1');
    expect(a1.address).toBe('A1');
    expect(a1.$col$row).toBe('$A$1');
    expect(a1.type).toBe(ExcelJS.ValueType.Null);

    expect((a1.value = 5)).toBe(5);
    expect(a1.value).toBe(5);
    expect(a1.type).toBe(ExcelJS.ValueType.Number);

    const strValue = 'Hello, World!';
    expect((a1.value = strValue)).toBe(strValue);
    expect(a1.value).toBe(strValue);
    expect(a1.type).toBe(ExcelJS.ValueType.String);

    const dateValue = new Date();
    expect((a1.value = dateValue)).toBe(dateValue);
    expect(a1.value).toBe(dateValue);
    expect(a1.type).toBe(ExcelJS.ValueType.Date);

    let formulaValue: Record<string, unknown> = { formula: 'A2', result: 5 };
    expect((a1.value = formulaValue)).toEqual(formulaValue);
    expect(a1.value).toEqual(formulaValue);
    expect(a1.type).toBe(ExcelJS.ValueType.Formula);

    formulaValue = { formula: 'A3' };
    expect((a1.value = formulaValue)).toEqual(formulaValue);
    expect(a1.value).toEqual({ formula: 'A3' });
    expect(a1.type).toBe(ExcelJS.ValueType.Formula);

    const hyperlinkValue = { hyperlink: 'http://www.link.com', text: 'www.link.com' };
    expect((a1.value = hyperlinkValue)).toEqual(hyperlinkValue);
    expect(a1.value).toEqual(hyperlinkValue);
    expect(a1.type).toBe(ExcelJS.ValueType.Hyperlink);

    expect((a1.value = null)).toBeNull();
    expect(a1.type).toBe(ExcelJS.ValueType.Null);

    // exceljs JSON values: any other object is kept and reported as String (written as JSON text).
    expect((a1.value = { json: 'data' })).toEqual({ json: 'data' });
    expect(a1.value).toEqual({ json: 'data' });
    expect(a1.type).toBe(ExcelJS.ValueType.String);
  });

  it.skip('validates options on construction — not supported: Cell constructor is internal', () => {});

  it('merges', () => {
    const a1 = ws.getCell('A1');
    const a2 = ws.getCell('A2');
    a1.value = 5;
    a2.value = 'Hello, World!';

    ws.mergeCells('A1:A2');
    expect(ws.getCell('A2').value).toBe(5);
    expect(ws.getCell('A2').type).toBe(ExcelJS.ValueType.Merge);
    expect(ws.getCell('A1').isMerged).toBeTruthy();
    expect(ws.getCell('A2').isMerged).toBeTruthy();
    expect(ws.getCell('A2').master).toBe(ws.getCell('A1'));
    expect(ws.getCell('A1').master).toBe(ws.getCell('A1'));

    // assignment of merged cells writes to the master
    ws.getCell('A2').value = 7;
    expect(ws.getCell('A1').value).toBe(7);
    ws.getCell('A2').value = 'Boo!';
    expect(ws.getCell('A1').value).toBe('Boo!');

    ws.unMergeCells('A1:A2');
    expect(ws.getCell('A2').type).toBe(ExcelJS.ValueType.Null);
    expect(ws.getCell('A1').isMerged).toBeFalsy();
    expect(ws.getCell('A2').isMerged).toBeFalsy();
    expect(ws.getCell('A2').master).toBe(ws.getCell('A2'));
    expect(ws.getCell('A1').master).toBe(ws.getCell('A1'));
  });

  it.skip('upgrades from string to hyperlink — not supported: internal _upgradeToHyperlink', () => {});
  it.skip("doesn't upgrade from non-string to hyperlink — not supported: internal _upgradeToHyperlink", () => {});

  it('inherits column styles', () => {
    ws.getColumn(1).style = { font: styles.fonts.arialBlackUI14 };
    expect(ws.getCell('A1').font).toEqual(styles.fonts.arialBlackUI14);
  });

  it('inherits row styles', () => {
    ws.getRow(1).style = { font: styles.fonts.broadwayRedOutline20 };
    expect(ws.getCell('A1').font).toEqual(styles.fonts.broadwayRedOutline20);
  });

  it('has effective types', () => {
    const a1 = new ExcelJS.Workbook().addWorksheet('S').getCell('A1');
    expect(a1.type).toBe(ExcelJS.ValueType.Null);
    expect(a1.effectiveType).toBe(ExcelJS.ValueType.Null);
    a1.value = 5;
    expect(a1.type).toBe(ExcelJS.ValueType.Number);
    expect(a1.effectiveType).toBe(ExcelJS.ValueType.Number);
    a1.value = 'Hello, World!';
    expect(a1.type).toBe(ExcelJS.ValueType.String);
    expect(a1.effectiveType).toBe(ExcelJS.ValueType.String);
    a1.value = new Date();
    expect(a1.type).toBe(ExcelJS.ValueType.Date);
    expect(a1.effectiveType).toBe(ExcelJS.ValueType.Date);
    a1.value = { formula: 'A2', result: 5 };
    expect(a1.type).toEqual(ExcelJS.ValueType.Formula);
    expect(a1.effectiveType).toBe(ExcelJS.ValueType.Number);
    a1.value = { formula: 'A2', result: 'Hello, World!' };
    expect(a1.type).toEqual(ExcelJS.ValueType.Formula);
    expect(a1.effectiveType).toBe(ExcelJS.ValueType.String);
    a1.value = { hyperlink: 'http://www.link.com', text: 'www.link.com' };
    expect(a1.type).toEqual(ExcelJS.ValueType.Hyperlink);
    expect(a1.effectiveType).toBe(ExcelJS.ValueType.Hyperlink);
  });
  it('shares formulas', () => {
    const ws = new ExcelJS.Workbook().addWorksheet('blort');
    const a1 = ws.getCell('A1');
    const b1 = ws.getCell('B1');
    const c1 = ws.getCell('C1');
    a1.value = 1;
    b1.value = { formula: 'A1+1', result: 2 };
    c1.value = { sharedFormula: 'B1', result: 3 };
    expect(b1.type).toBe(ExcelJS.ValueType.Formula);
    expect(b1.formulaType).toBe(ExcelJS.FormulaType.Master);
    expect(c1.type).toBe(ExcelJS.ValueType.Formula);
    expect(c1.formulaType).toBe(ExcelJS.FormulaType.Shared);
    expect(c1.formula).toBe('B1+1');
  });
  it.skip('escapes dangerous html — not supported: cell.html', () => {});

  it('can set comment', () => {
    const a1 = ws.getCell('A1');
    const comment = {
      texts: [{ font: { size: 12, color: { theme: 0 }, name: 'Calibri', family: 2, scheme: 'minor' }, text: 'This is ' }],
      margins: { insetmode: 'auto', inset: [0.13, 0.13, 0.25, 0.25] },
      protection: { locked: 'True', lockText: 'True' },
      editAs: 'twoCells',
    };
    a1.note = comment;
    a1.value = 'test set value';
    expect(a1.note).toEqual(comment);
  });

  it('Cell comments supports setting margins, protection, and position properties (without the model defaults)', () => {
    const a1 = ws.getCell('A1');
    const comment = {
      texts: [{ font: { size: 12, color: { theme: 0 }, name: 'Calibri', family: 2, scheme: 'minor' }, text: 'This is ' }],
      protection: { locked: 'False', lockText: 'True' },
    };
    a1.note = comment;
    a1.value = 'test set value';
    expect(a1.note.texts).toEqual(comment.texts);
    expect(a1.note.protection).toEqual(comment.protection);
  });
});
