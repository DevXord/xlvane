// @xlvane/compat shared formulas, K4 review item S1: `cell.formula` of a `{ sharedFormula: master }`
// dependent moved the master's text with shiftFormula on every read, with no cache, and the
// load-time budget (maxSharedStringChars) is spent only once. A 4 KB file with a 1 000 000-character
// master cost about 200 ms per read of one dependent's formula, every time, and reading the formula
// of each of 250 dependents once took about 47 s (xlsx.load) and 50 s (stream.xlsx.WorkbookReader).
// exceljs caches the translated formula per value (lib/doc/cell.js L861-869 `_getTranslatedFormula`).
// Now the moved text is kept per dependent cell (dropped when the dependent's value, its master's
// formula or either position changes), and the shared-formula text moved in one workbook is limited.
// The tests count the calls of shiftFormula instead of measuring time.
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { buildXlsx } from '../helpers/xlsx-builder';
import { ExcelJS } from '../exceljs/helpers';

const spy = vi.hoisted(() => ({ calls: 0, identity: false }));

vi.mock('xlvane', async (importOriginal) => {
  const mod = await importOriginal<Record<string, unknown>>();
  const real = mod['shiftFormula'] as (text: string, dRow: number, dCol: number) => string | null;
  return {
    ...mod,
    // `identity` returns the text unmoved, so a huge master costs no time in the budget tests.
    shiftFormula: (text: string, dRow: number, dCol: number): string | null => {
      spy.calls++;
      return spy.identity ? text : real(text, dRow, dCol);
    },
  };
});

beforeEach(() => {
  spy.calls = 0;
  spy.identity = false;
});

/** A master in A1 (`text`, shared over A1:A(dependents + 1)) and its dependents below, with B values. */
function shared(text: string, dependents: number): Uint8Array {
  const parts = [`<row r="1"><c r="A1"><f t="shared" ref="A1:A${String(dependents + 1)}" si="0">${text}</f><v>2</v></c><c r="B1"><v>1</v></c></row>`];
  for (let r = 2; r <= dependents + 1; r++) parts.push(`<row r="${String(r)}"><c r="A${String(r)}"><f t="shared" si="0"/><v>${String(2 * r)}</v></c><c r="B${String(r)}"><v>${String(r)}</v></c></row>`);
  return buildXlsx({ sheets: [{ name: 'S', data: parts.join('') }] });
}

async function load(bytes: Uint8Array): Promise<any> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes);
  return wb.getWorksheet('S');
}

/** The error `fn` throws (fails the test when it throws none). */
function thrown(fn: () => unknown): any {
  try {
    fn();
  } catch (error) {
    return error;
  }
  throw new Error('expected an error');
}

describe('K4 S1: the moved formula of a shared-formula dependent is kept', () => {
  it('xlsx.load: reading a dependent formula again does not move the master text again', async () => {
    const sheet = await load(shared('B1*2', 3));
    spy.calls = 0;
    for (let i = 0; i < 5; i++) expect(sheet.getCell('A3').formula).toBe('B3*2');
    expect(spy.calls).toBe(1);
    expect(sheet.getCell('A4').formula).toBe('B4*2');
    expect(spy.calls).toBe(2);
  });

  it('a formula that cannot be moved is not tried again', () => {
    const sheet = new ExcelJS.Workbook().addWorksheet('S');
    sheet.getCell('A1').value = { formula: 'XFD1' };
    sheet.getCell('B1').value = { sharedFormula: 'A1' };
    for (let i = 0; i < 3; i++) expect(sheet.getCell('B1').formula).toBeUndefined();
    expect(spy.calls).toBe(1);
  });

  it('a new value of the dependent gives its own formula', async () => {
    const sheet = await load(shared('B1*2', 3));
    expect(sheet.getCell('A3').formula).toBe('B3*2');
    sheet.getCell('C1').value = { formula: 'D1+1' };
    sheet.getCell('A3').value = { sharedFormula: 'C1', result: 0 };
    // Moved from C1 to A3: two rows down, two columns left.
    expect(sheet.getCell('A3').formula).toBe('B3+1');
    // The same object changed in place.
    (sheet.getCell('A3').value as Record<string, unknown>)['sharedFormula'] = 'A1';
    expect(sheet.getCell('A3').formula).toBe('B3*2');
    sheet.getCell('A3').value = { formula: 'E3' };
    expect(sheet.getCell('A3').formula).toBe('E3');
    sheet.getCell('A3').value = 7;
    expect(sheet.getCell('A3').formula).toBeUndefined();
  });

  it('a new formula of the master gives fresh text (exceljs keeps the old text here)', async () => {
    const sheet = await load(shared('B1*2', 3));
    expect(sheet.getCell('A3').formula).toBe('B3*2');
    sheet.getCell('A1').value = { formula: 'B1*3', result: 3, ref: 'A1:A4', shareType: 'shared' };
    expect(sheet.getCell('A3').formula).toBe('B3*3');
    // The master's value changed in place.
    (sheet.getCell('A1').value as Record<string, unknown>)['formula'] = 'C1*4';
    expect(sheet.getCell('A3').formula).toBe('C3*4');
    expect(sheet.getCell('A2').formula).toBe('C2*4');
    // A master that is itself a dependent.
    sheet.getCell('A1').value = { formula: '$B$1+B1' };
    sheet.getCell('C2').value = { sharedFormula: 'A3' };
    expect(sheet.getCell('C2').formula).toBe('$B$1+D2');
    sheet.getCell('A1').value = { formula: 'B1-1' };
    expect(sheet.getCell('C2').formula).toBe('D2-1');
  });

  it('stream.xlsx.WorkbookReader: a dependent given a { sharedFormula } value keeps its moved formula too', async () => {
    const formulas: unknown[] = [];
    for await (const ws of new ExcelJS.stream.xlsx.WorkbookReader(shared('B1*2', 3), {})) {
      for await (const row of ws) {
        if (row.number === 1) continue;
        // No hidden `formula` here: cell.formula moves the master text the reader keeps.
        const cell = row.getCell(3);
        cell.value = { sharedFormula: 'A1' };
        spy.calls = 0;
        for (let i = 0; i < 4; i++) formulas.push(cell.formula);
        expect(spy.calls).toBe(1);
      }
    }
    expect(formulas).toEqual([...Array<string>(4).fill('D2*2'), ...Array<string>(4).fill('D3*2'), ...Array<string>(4).fill('D4*2')]);
  });
});

describe('K4 S1: the shared-formula text moved in a workbook is limited', () => {
  // 1 000 000 characters, shared with 250 cells: 250 000 000 characters if each dependent's
  // formula is read once.
  const huge = Array.from({ length: 333_333 }, () => 'B1').join('+');

  it('xlsx.load: reading the formula of every dependent stops with XLV_E_LIMIT_EXCEEDED', async () => {
    const sheet = await load(shared(huge, 250));
    spy.identity = true;
    spy.calls = 0;
    const error = thrown(() => {
      for (let r = 2; r <= 251; r++) void sheet.getCell(`A${String(r)}`).formula;
    });
    expect(error.code).toBe('XLV_E_LIMIT_EXCEEDED');
    expect(error.details.limit).toBe('sharedFormulaText');
    expect(typeof error.details.value).toBe('number');
    expect(error.location).toMatchObject({ sheet: 'S' });
    expect(error.hint).toMatch(/result/);
    // Far fewer than the 250 moves, and the work done is within the limit.
    expect(spy.calls).toBeLessThan(50);
    expect(spy.calls * 2 * huge.length).toBeLessThanOrEqual(error.details.value);
    // Reading the moved texts already built costs nothing.
    spy.calls = 0;
    void sheet.getCell('A2').formula;
    expect(spy.calls).toBe(0);
  });

  it('stream.xlsx.WorkbookReader: the same limit for the formulas the reader gives', async () => {
    spy.identity = true;
    let error: any;
    try {
      for await (const ws of new ExcelJS.stream.xlsx.WorkbookReader(shared(huge, 250), {})) {
        for await (const row of ws) void row.getCell(1).formula;
      }
    } catch (e) {
      error = e;
    }
    expect(error?.code).toBe('XLV_E_LIMIT_EXCEEDED');
    expect(error.details.limit).toBe('sharedFormulaText');
    expect(spy.calls).toBeLessThan(50);
  });

  it('stream.xlsx.WorkbookReader: the same limit for a { sharedFormula } value set on a read row', async () => {
    spy.identity = true;
    let error: any;
    try {
      for await (const ws of new ExcelJS.stream.xlsx.WorkbookReader(shared(huge, 250), {})) {
        for await (const row of ws) {
          if (row.number === 1) continue;
          const cell = row.getCell(3);
          cell.value = { sharedFormula: 'A1' };
          void cell.formula;
        }
      }
    } catch (e) {
      error = e;
    }
    expect(error?.code).toBe('XLV_E_LIMIT_EXCEEDED');
    expect(error.details.limit).toBe('sharedFormulaText');
    expect(spy.calls).toBeLessThan(50);
  });

  it('dependents read with the stream reader and those set on its rows share one limit', async () => {
    spy.identity = true;
    let error: any;
    let read = 0;
    try {
      for await (const ws of new ExcelJS.stream.xlsx.WorkbookReader(shared(huge, 250), {})) {
        for await (const row of ws) {
          if (row.number === 1) continue;
          void row.getCell(1).formula;
          const cell = row.getCell(3);
          cell.value = { sharedFormula: 'A1' };
          void cell.formula;
          read++;
        }
      }
    } catch (e) {
      error = e;
    }
    expect(error?.code).toBe('XLV_E_LIMIT_EXCEEDED');
    // Two moves per row: half as many rows as with one kind alone.
    expect(read).toBeLessThan(25);
  });

  it('a workbook within the limit reads every formula: 250 dependents of a 50 000-character master', async () => {
    const master = Array.from({ length: 16_667 }, () => 'B1').join('+');
    const sheet = await load(shared(master, 250));
    spy.identity = true;
    for (let r = 2; r <= 251; r++) expect(sheet.getCell(`A${String(r)}`).formula).toBe(master);
    expect(spy.calls).toBe(250);
  });

  it('formulas no longer than Office writes are counted only against the 512 Mi budget: a 600-character master shared with 40 000 cells (review of D-K04)', async () => {
    const master = Array.from({ length: 200 }, () => 'B1').join('+').slice(0, 599);
    const sheet = await load(shared(master, 40_000));
    spy.identity = true;
    for (let r = 2; r <= 40_001; r++) void sheet.getCell(`A${String(r)}`).formula;
    expect(spy.calls).toBe(40_000);
    // Inserting a row reads every dependent's formula; it works too.
    sheet.insertRow(1, ['x']);
    expect(sheet.getCell('A40002').formula).toBeTypeOf('string');
  }, 120_000);

  it('formulas no longer than Office writes still stop at 512 Mi characters', async () => {
    const master = Array.from({ length: 2731 }, () => 'B1').join('+').slice(0, 8192);
    expect(master.length).toBe(8192);
    // Built in code: a loaded file this size already stops at the load-time budget.
    const sheet = new ExcelJS.Workbook().addWorksheet('S');
    sheet.getCell('A1').value = { formula: master };
    for (let r = 2; r <= 33_001; r++) sheet.getCell(`A${String(r)}`).value = { sharedFormula: 'A1' };
    spy.identity = true;
    const error = thrown(() => {
      for (let r = 2; r <= 33_001; r++) void sheet.getCell(`A${String(r)}`).formula;
    });
    expect(error.details).toMatchObject({ limit: 'sharedFormulaText', value: 512 * 1024 * 1024 });
    expect(spy.calls).toBe(Math.floor((512 * 1024 * 1024) / (2 * 8192)));
  });

  it('the moved formulas are still written when the limit was not reached', async () => {
    const sheet = await load(shared('B1*2', 3));
    for (let r = 2; r <= 4; r++) void sheet.getCell(`A${String(r)}`).formula;
    const back = await new ExcelJS.Workbook().xlsx.load(await sheet.workbook.xlsx.writeBuffer());
    expect(back.getWorksheet('S').getCell('A4').formula).toBe('B4*2');
  });
});
