// xlvane tests for row and column splicing in @xlvane/exceljs (phase B, docs/decisions.md D-B03):
// what xlvane moves like a spreadsheet application where exceljs 4.4.0 does not
// (docs/facts/exceljs.md X4-15, docs/exceljs-compat.md).
import { describe, expect, it } from 'vitest';
import { ExcelJS, expectNotSupported, roundTrip } from './helpers';

const sheet = () => new ExcelJS.Workbook().addWorksheet('S');

describe('spliceRows and insertRow move everything with the row', () => {
  it('moves hidden flags, heights, notes and links with their rows', () => {
    const ws = sheet();
    ws.addRow(['r1']);
    ws.addRow(['r2']);
    ws.getRow(2).hidden = true;
    ws.getRow(2).height = 30;
    ws.getCell('A2').note = 'n2';
    ws.getCell('B2').value = { text: 'site', hyperlink: 'https://example.com' };
    ws.insertRow(2, ['new']);
    expect(ws.getRow(2).hidden).toBe(false);
    expect(ws.getRow(2).height).toBeUndefined();
    expect(ws.getRow(3).hidden).toBe(true);
    expect(ws.getRow(3).height).toBe(30);
    expect(ws.getCell('A3').note).toBe('n2');
    expect(ws.getCell('A2').note).toBeUndefined();
    expect(ws.getCell('B3').hyperlink).toBe('https://example.com');
  });

  it('keeps row objects: a row reference follows its row', () => {
    const ws = sheet();
    ws.addRow(['a']);
    const row = ws.addRow(['b']);
    ws.insertRow(1, ['top']);
    expect(row.number).toBe(3);
    expect(row.getCell(1).address).toBe('A3');
  });

  it('removes rows past the end and rows that were never created', () => {
    const ws = sheet();
    for (let i = 1; i <= 5; i++) ws.addRow([`r${String(i)}`]);
    ws.spliceRows(2, 3);
    expect(ws.getColumn(1).values).toEqual([, 'r1', 'r5']);
    ws.spliceRows(2, 10);
    expect(ws.rowCount).toBe(1);
    ws.spliceRows(4, 2, ['x']);
    expect(ws.getCell('A4').value).toBe('x');
  });

  it('moves and shrinks the autofilter range', () => {
    const ws = sheet();
    ws.autoFilter = 'A1:C5';
    ws.insertRows(3, [[1], [2]]);
    expect(ws.autoFilter).toBe('A1:C7');
    ws.spliceRows(1, 1);
    expect(ws.autoFilter).toBe('A1:C6');
    ws.spliceColumns(2, 1);
    expect(ws.autoFilter).toBe('A1:B6');
    ws.autoFilter = { from: { row: 2, column: 1 }, to: { row: 3, column: 2 } };
    ws.spliceRows(2, 2);
    expect(ws.autoFilter).toBeNull();
  });

  it('keeps defined names in step and writes the result', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    for (let i = 1; i <= 5; i++) ws.addRow([i]);
    wb.definedNames.add('S!$A$1:$A$5', 'block');
    ws.insertRow(3, [99]);
    expect(wb.definedNames.getRanges('block').ranges).toEqual(['S!$A$1:$A$6']);
    const ws2 = (await roundTrip(wb)).getWorksheet('S');
    expect(ws2.getColumn(1).values).toEqual([, 1, 2, 99, 3, 4, 5]);
  });

  it('refuses to push rows beyond row 1048576', () => {
    const ws = sheet();
    ws.getCell('A1048576').value = 'last';
    expect(() => ws.insertRow(1, ['x'])).toThrow(expect.objectContaining({ code: 'XLV_E_INVALID_ARGUMENT' }));
    expect(ws.getCell('A1048576').value).toBe('last');
  });
});

describe('spliceColumns', () => {
  it('keeps falsy inserted values and does not rewrite headers', () => {
    const ws = sheet();
    ws.columns = [{ header: 'H1' }, { header: 'H2', key: 'two' }];
    ws.getCell('B1').value = 'Changed';
    ws.spliceColumns(1, 0, [0, false, '']);
    expect(ws.getCell('A1').value).toBe(0);
    expect(ws.getCell('A2').value).toBe(false);
    expect(ws.getCell('A3').value).toBe('');
    expect(ws.getCell('C1').value).toBe('Changed');
    expect(ws.getColumn('two').number).toBe(3);
  });

  it('moves notes, inserts values past the existing rows, and drops the key of a removed column', () => {
    const ws = sheet();
    ws.columns = [{ key: 'a' }, { key: 'b' }];
    ws.getCell('B1').note = 'b1';
    ws.spliceColumns(1, 1, ['x1', 'x2', 'x3']);
    expect(ws.getCell('B1').note).toBe('b1');
    expect(ws.getCell('A3').value).toBe('x3');
    expect(ws.getColumnKey('a')).toBeUndefined();
    expect(ws.getColumn('b').letter).toBe('B');
  });
});

describe('streamed worksheets', () => {
  it('refuse row and column splices, as exceljs has no such methods there', async () => {
    const writer = new ExcelJS.stream.xlsx.WorkbookWriter({ stream: new WritableStream() });
    const ws = writer.addWorksheet('S');
    ws.addRow([1]);
    await expectNotSupported(() => ws.insertRow(1, [0]));
    await expectNotSupported(() => ws.insertRows(1, [[0]]));
    await expectNotSupported(() => ws.spliceRows(1, 1));
    await expectNotSupported(() => ws.duplicateRow(1, 1));
    await expectNotSupported(() => ws.spliceColumns(1, 1));
    ws.getRow(1).splice(1, 0, 'before');
    expect(ws.getRow(1).values).toEqual([, 'before', 1]);
  });
});

describe('formulas follow splices (xlvane adjustFormula, D-B01)', () => {
  const f = (ws: ReturnType<typeof sheet>, ref: string) => (ws.getCell(ref).value as { formula: string }).formula;

  it('insertRow(5) turns SUM(A1:A10) into SUM(A1:A11) on the same sheet and on other sheets', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    const other = wb.addWorksheet('T');
    for (let i = 1; i <= 10; i++) ws.getCell(`A${String(i)}`).value = i;
    ws.getCell('B1').value = { formula: 'SUM(A1:A10)' };
    other.getCell('A1').value = { formula: 'S!A7+$A$1' };
    ws.insertRow(5, [99]);
    expect(f(ws, 'B1')).toBe('SUM(A1:A11)');
    expect((other.getCell('A1').value as { formula: string }).formula).toBe('S!A8+$A$1');
  });

  it('spliceRows removing a referenced row gives #REF!, removing part of a range shrinks it', () => {
    const ws = sheet();
    ws.getCell('C1').value = { formula: 'A5*2' };
    ws.getCell('C2').value = { formula: 'SUM(A4:A6)' };
    ws.spliceRows(5, 1);
    expect(f(ws, 'C1')).toBe('#REF!*2');
    expect(f(ws, 'C2')).toBe('SUM(A4:A5)');
  });

  it('spliceColumns adjusts column references, including $-absolute ones', () => {
    const ws = sheet();
    ws.getCell('A1').value = { formula: 'C1+$D$1' };
    ws.spliceColumns(2, 0, []);
    expect(f(ws, 'A1')).toBe('D1+$E$1');
  });

  it('duplicateRow moves relative references of the copied formulas with each copy', () => {
    const ws = sheet();
    ws.getCell('A1').value = 1;
    ws.getCell('B1').value = { formula: 'A1*$A$1' };
    ws.duplicateRow(1, 2, true);
    expect(f(ws, 'B2')).toBe('A2*$A$1');
    expect(f(ws, 'B3')).toBe('A3*$A$1');
  });

  it('refuses before changing anything when a formula cannot be adjusted', async () => {
    const ws = sheet();
    ws.getCell('A1').value = 'keep';
    ws.getCell('B1').value = { formula: 'A2#' };
    await expectNotSupported(() => ws.insertRow(1, ['x']));
    expect(ws.getCell('A1').value).toBe('keep');
  });
});

// Regressions found by the independent review of phase B (docs/decisions.md D-B03 addendum).
describe('review fixes: nothing changes when a splice is refused', () => {
  const code = (fn: () => unknown): string => {
    try {
      fn();
    } catch (error) {
      return (error as { code?: string }).code ?? 'NO_CODE';
    }
    return 'NONE';
  };

  it('spliceColumns checks every row for cells pushed beyond XFD before anything moves', () => {
    const ws = sheet();
    ws.getCell('A1').value = 'a';
    ws.getCell('B1').value = 'b';
    ws.mergeCells('A3:B4');
    ws.getCell(2, 16_384).value = 'far';
    expect(code(() => ws.spliceColumns(1, 0, ['x']))).toBe('XLV_E_INVALID_ARGUMENT');
    expect(ws.getCell('A1').value).toBe('a');
    expect(ws.getCell('B1').value).toBe('b');
    expect([...ws._merges.values()]).toEqual([{ top: 3, left: 1, bottom: 4, right: 2 }]);
    expect(ws.getCell(2, 16_384).value).toBe('far');
  });

  it('refuses inserts that would themselves reach beyond the grid', () => {
    const ws = sheet();
    ws.mergeCells('A2:B3');
    expect(code(() => ws.spliceRows(1_048_576, 0, [1], [2]))).toBe('XLV_E_INVALID_ARGUMENT');
    expect(ws.findRow(1_048_576)).toBeUndefined();
    expect(code(() => ws.spliceColumns(16_384, 0, [1], [2]))).toBe('XLV_E_INVALID_ARGUMENT');
    expect([...ws._merges.values()]).toEqual([{ top: 2, left: 1, bottom: 3, right: 2 }]);
    const row = ws.getRow(5);
    expect(code(() => row.splice(16_384, 0, 'a', 'b'))).toBe('XLV_E_INVALID_ARGUMENT');
    expect(ws.findCell(5, 16_384)).toBeUndefined();
  });
});

describe('review fixes: formulas', () => {
  const f = (ws: ReturnType<typeof sheet>, ref: string) => (ws.getCell(ref).value as { formula: string }).formula;

  it('keeps the text of formulas given in inserted rows and columns', () => {
    const ws = sheet();
    ws.getCell('A1').value = 1;
    ws.getCell('A2').value = 2;
    ws.insertRow(2, [{ formula: 'A1+A3' }]);
    expect(f(ws, 'A2')).toBe('A1+A3');
    ws.spliceColumns(1, 0, [null, null, null, null, { formula: 'B1*2' }]);
    expect(f(ws, 'A5')).toBe('B1*2');
  });

  it('duplicateRow(insert = true) copies the adjusted source formula, moved by the copy offset', () => {
    const ws = sheet();
    ws.getCell('B1').value = { formula: 'A5' };
    ws.duplicateRow(1, 1, true);
    // The insertion moves A5 to A6 in the source; the copy one row down refers to A7.
    expect(f(ws, 'B1')).toBe('A6');
    expect(f(ws, 'B2')).toBe('A7');
  });

  it('duplicateRow(insert = false) overwrites rows in place without touching other formulas', () => {
    const ws = sheet();
    for (let i = 1; i <= 3; i++) ws.getCell(`A${String(i)}`).value = i;
    ws.getCell('B1').value = { formula: 'A1*2' };
    ws.getCell('C5').value = { formula: 'A2' };
    ws.getCell('C6').value = { formula: 'SUM(A2:A3)' };
    ws.getCell('C7').value = { formula: 'SUM(A1:A2)' };
    ws.mergeCells('D2:E2');
    ws.duplicateRow(1, 1, false);
    expect(f(ws, 'C5')).toBe('A2');
    expect(f(ws, 'C6')).toBe('SUM(A2:A3)');
    expect(f(ws, 'C7')).toBe('SUM(A1:A2)');
    expect(ws.getCell('A2').value).toBe(1);
    expect(f(ws, 'B2')).toBe('A2*2');
    expect(ws.getCell('A3').value).toBe(3);
    expect(ws._merges.size).toBe(0);
  });

  it('accepts large deletions and insertions when the workbook has formulas', () => {
    const ws = sheet();
    ws.getCell('A1').value = { formula: 'B2' };
    ws.spliceRows(1_000_000, 100_000);
    expect(f(ws, 'A1')).toBe('B2');
  });
});

describe('review fixes: merges, styles and the grid edge', () => {
  it('keeps the own style of each merged cell and never shares style objects', () => {
    const ws = sheet();
    ws.mergeCells('B10:C10');
    ws.getCell('B10').font = { bold: true };
    ws.getCell('C10').border = { right: { style: 'thick' } };
    ws.insertRow(1, [1]);
    expect(ws.getCell('C11').border).toEqual({ right: { style: 'thick' } });
    expect(ws.getCell('B11').font).toEqual({ bold: true });
    expect(ws.getCell('B11').style).not.toBe(ws.getCell('C11').style);
    expect(ws.getCell('C11').font).not.toBe(ws.getCell('B11').font);
  });

  it('gives cells inserted inside a merge a copy of the master style', () => {
    const ws = sheet();
    ws.getCell('A1').font = { italic: true };
    ws.mergeCells('A1:A3');
    ws.insertRow(2, ['x']);
    expect(ws.getCell('A2').font).toEqual({ italic: true });
    expect(ws.getCell('A2').font).not.toBe(ws.getCell('A1').font);
  });

  it('drops a merge whose top-left cell is removed (R-14)', () => {
    const ws = sheet();
    ws.mergeCells('A2:B5');
    ws.spliceRows(2, 1);
    expect(ws._merges.size).toBe(0);
    ws.mergeCells('D2:E3');
    ws.spliceColumns(4, 1);
    expect(ws._merges.size).toBe(0);
  });

  it('stops the autofilter at the last row and column (R-09)', () => {
    const ws = sheet();
    ws.autoFilter = 'A1:C1048576';
    ws.insertRow(1, [0]);
    expect(ws.autoFilter).toBe('A2:C1048576');
    ws.autoFilter = 'A1:XFD1';
    ws.spliceColumns(1, 0, [5]);
    expect(ws.autoFilter).toBe('B1:XFD1');
    // A merge always has cells, so it can never be pushed beyond the grid: such a splice is refused.
    const ws2 = sheet();
    ws2.mergeCells('A1048575:B1048576');
    expect(() => ws2.insertRow(1, [0])).toThrow(expect.objectContaining({ code: 'XLV_E_INVALID_ARGUMENT' }));
    expect([...ws2._merges.values()]).toEqual([{ top: 1_048_575, left: 1, bottom: 1_048_576, right: 2 }]);
  });
});

describe('review fixes: defined names', () => {
  it('match the sheet name ignoring case when rows move', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Data');
    wb.definedNames.add('data!$A$5', 'five');
    ws.insertRow(1, [0]);
    expect(wb.definedNames.getRanges('five').ranges).toEqual(['data!$A$6']);
  });

  it('do not block splices when a formula refers to a sheet-qualified name', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    wb.addWorksheet('Data');
    ws.getCell('A1').value = { formula: 'SUM(Data!Total)+B5' };
    ws.insertRow(3, [1]);
    expect((ws.getCell('A1').value as { formula: string }).formula).toBe('SUM(Data!Total)+B6');
  });
});
