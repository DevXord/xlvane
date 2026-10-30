// Ported from exceljs v4.4.0 spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js
// (MIT, Copyright (c) 2014-2019 Guyon Roche), changes: Mocha/Chai → Vitest with async/await;
// files go to a temporary directory; cells are added with `addRow` instead of
// `worksheet.getCell` (the streaming worksheet subset is `addRow`, row style setters and
// `commit`); fonts without `family`/`outline`; the row style read-back is removed (no row
// default styles after reading, docs/decisions.md D-206); notes and conditional formatting
// expect XLV_E_NOT_SUPPORTED (conditional formatting); notes read back without the font
// properties family/scheme and without the box settings margins, protection, editAs (D-A14);
// tests that need createTestBook, shared formulas, data
// validation, images or zip options are skipped. One test added: names changed after the
// first committed row.
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { ExcelJS, expectNotSupported, styles, withTempDir } from './helpers';

// types/node-modules.d.ts declares only the node:fs members the project uses; createWriteStream
// is https://nodejs.org/api/fs.html#fscreatewritestreampath-options.
async function createWriteStream(path: string, options: { flags: string }): Promise<unknown> {
  const fs = (await import('node:fs')) as unknown as { createWriteStream: (p: string, o: { flags: string }) => unknown };
  return fs.createWriteStream(path, options);
}

async function readBack(file: string): Promise<any> {
  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.readFile(file);
  return wb2;
}

describe('WorkbookWriter', () => {
  it('creates sheets with correct names', () => {
    const wb = new ExcelJS.stream.xlsx.WorkbookWriter();
    const ws1 = wb.addWorksheet('Hello, World!');
    expect(ws1.name).toBe('Hello, World!');
    const ws2 = wb.addWorksheet();
    expect(ws2.name).toMatch(/sheet\d+/);
  });

  describe('Serialise', () => {
    it.skip('xlsx file — not supported: createTestBook fixture (needs data validation)', () => {});
    it.skip('shared formula — not supported: fillFormula', () => {});

    it('auto filter', async () => {
      await withTempDir(async (dir) => {
        const filename = join(dir, 'wb.test.xlsx');
        const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename, useStyles: false });
        const ws = wb.addWorksheet('Hello');
        // As upstream, no row is committed before autoFilter is set.
        ws.addRow([1, 1]);
        ws.addRow([2, 2]);
        ws.addRow([3, 3]);
        ws.autoFilter = 'A1:B1';
        ws.commit();
        await wb.commit();
        expect((await readBack(filename)).getWorksheet('Hello').autoFilter).toBe('A1:B1');
      });
    });

    it.skip('Without styles — not supported: createTestBook fixture', () => {});

    it('serializes row styles and columns properly', async () => {
      await withTempDir(async (dir) => {
        const filename = join(dir, 'wb.test.xlsx');
        const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename, useStyles: true });
        const ws = wb.addWorksheet('blort');
        const colStyle = { font: styles.roundTripFonts.comicSansUdB16, alignment: styles.namedAlignments.middleCentre };
        ws.columns = [{ header: 'A1', width: 10 }, { header: 'B1', style: colStyle }, { header: 'C1', width: 30 }, { header: 'D1' }];
        const row2 = ws.addRow(['A2', 'B2', 'C2']);
        row2.font = styles.roundTripFonts.broadwayRed20;
        row2.commit();
        ws.addRow(['A3', 'B3', 'C3']).commit();
        ws.commit();
        await wb.commit();

        const ws2 = (await readBack(filename)).getWorksheet('blort');
        for (const address of ['A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A3', 'B3', 'C3']) {
          expect(ws2.getCell(address).value).toBe(address);
        }
        expect(ws2.getCell('B1').font).toEqual(styles.roundTripFonts.comicSansUdB16);
        expect(ws2.getCell('B1').alignment).toEqual(styles.namedAlignments.middleCentre);
        expect(ws2.getCell('A2').font).toEqual(styles.roundTripFonts.broadwayRed20);
        expect(ws2.getCell('B2').font).toEqual(styles.roundTripFonts.broadwayRed20);
        expect(ws2.getCell('C2').font).toEqual(styles.roundTripFonts.broadwayRed20);
        expect(ws2.getCell('B3').font).toEqual(styles.roundTripFonts.comicSansUdB16);
        expect(ws2.getCell('B3').alignment).toEqual(styles.namedAlignments.middleCentre);

        expect(ws2.getColumn(2).font).toEqual(styles.roundTripFonts.comicSansUdB16);
        expect(ws2.getColumn(2).alignment).toEqual(styles.namedAlignments.middleCentre);
        expect(ws2.getColumn(2).width).toBe(9);
        expect(ws2.getColumn(4).width).toBe(undefined);
      });
    });

    it('rich text', async () => {
      await withTempDir(async (dir) => {
        const filename = join(dir, 'wb.test.xlsx');
        const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename, useStyles: true });
        const ws = wb.addWorksheet('Hello');
        // Change for xlvane: the cells are added with addRow (see the header).
        ws.addRow([
          {
            richText: [
              { font: { color: { argb: 'FF0000' } }, text: 'red ' },
              { font: { color: { argb: '00FF00' }, bold: true }, text: ' bold green' },
            ],
          },
          'plain text',
        ]).commit();
        ws.commit();
        await wb.commit();
        const ws2 = (await readBack(filename)).getWorksheet('Hello');
        // Change for xlvane: a 6-digit colour is stored with alpha FF, so it reads back with 8
        // digits (CT_Color rgb is 4 bytes, W2-16); exceljs writes and reads the 6 digits as given.
        expect(ws2.getCell('A1').value).toEqual({
          richText: [
            { font: { color: { argb: 'FFFF0000' } }, text: 'red ' },
            { font: { color: { argb: 'FF00FF00' }, bold: true }, text: ' bold green' },
          ],
        });
        expect(ws2.getCell('B1').value).toBe('plain text');
      });
    });

    it('A lot of sheets', async () => {
      await withTempDir(async (dir) => {
        const filename = join(dir, 'wb.test.xlsx');
        const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename });
        const numSheets = 90;
        for (let i = 1; i <= numSheets; i++) wb.addWorksheet(`sheet${String(i)}`).addRow([i]).commit();
        await wb.commit();
        const wb2 = await readBack(filename);
        for (let i = 1; i <= numSheets; i++) {
          const ws2 = wb2.getWorksheet(`sheet${String(i)}`);
          expect(ws2).toBeTruthy();
          expect(ws2.getCell('A1').value).toBe(i);
        }
      });
    });

    it('addRow', async () => {
      await withTempDir(async (dir) => {
        const filename = join(dir, 'wb.test.xlsx');
        const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
          stream: await createWriteStream(filename, { flags: 'w' }),
          useStyles: true,
          useSharedStrings: true,
        });
        const worksheet = workbook.addWorksheet('test');
        const newRow = worksheet.addRow(['hello']);
        newRow.commit();
        worksheet.commit();
        await workbook.commit();
        // Change for xlvane: also check the file content.
        expect((await readBack(filename)).getWorksheet('test').getCell('A1').value).toBe('hello');
      });
    });

    it('defined names', async () => {
      await withTempDir(async (dir) => {
        const filename = join(dir, 'wb.test.xlsx');
        const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename });
        const ws = wb.addWorksheet('blort');
        // Change for xlvane: the cells are added with addRow; as upstream, nothing is committed
        // before the names are set (names are written when the first row is committed, D-A06).
        const row1 = ws.addRow([5, null, null, null, 'grün']);
        row1.getCell(1).name = 'five';
        row1.getCell(5).name = 'greens';
        ws.addRow([null, null, null, null, 'vert']).getCell(5).name = 'greens';
        const row3 = ws.addRow(['drei', 'trois', null, null, 'verde']);
        row3.getCell(1).name = 'threes';
        row3.getCell(2).value = 'san';
        row3.getCell(2).name = 'threes';
        row3.getCell(5).name = 'greens';
        await wb.commit();

        const ws2 = (await readBack(filename)).getWorksheet('blort');
        expect(ws2.getCell('A1').name).toBe('five');
        expect(ws2.getCell('A3').name).toBe('threes');
        expect(ws2.getCell('B3').name).toBe('threes');
        expect(ws2.getCell('E1').name).toBe('greens');
        expect(ws2.getCell('E2').name).toBe('greens');
        expect(ws2.getCell('E3').name).toBe('greens');
      });
    });

    it('names or properties changed after the first committed row — not supported', async () => {
      await withTempDir(async (dir) => {
        const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: join(dir, 'wb.test.xlsx') });
        const ws = wb.addWorksheet('blort');
        ws.addRow([1]).commit();
        ws.getCell('A1').name = 'late';
        await expectNotSupported(() => wb.commit());
      });
    });

    it('does not escape special xml characters', async () => {
      await withTempDir(async (dir) => {
        const filename = join(dir, 'wb.test.xlsx');
        const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename, useSharedStrings: true });
        const ws = wb.addWorksheet('blort');
        const xmlCharacters = 'xml characters: & < > "';
        ws.addRow([xmlCharacters]).commit();
        await wb.commit();
        expect((await readBack(filename)).getWorksheet('blort').getCell('A1').value).toBe(xmlCharacters);
      });
    });

    it.skip('serializes and deserializes dataValidations — not supported: data validation', () => {});
    it.skip('with zip compression option — not supported: zip compression options', () => {});

    for (const [title, box] of [
      ['writes notes', { margins: { insetmode: 'auto', inset: [0.13, 0.13, 0.25, 0.25] }, protection: { locked: 'True', lockText: 'True' }, editAs: 'twoCells' }],
      ['Cell annotation supports setting margins and protection properties', { margins: { insetmode: 'custom', inset: [0.25, 0.25, 0.35, 0.35] }, protection: { locked: 'False', lockText: 'False' }, editAs: 'oneCells' }],
    ] as const) {
      it(`${title} (box settings are not written)`, async () => {
        await withTempDir(async (dir) => {
          const filename = join(dir, 'wb.test.xlsx');
          const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename });
          const ws = wb.addWorksheet('Hello');
          const row = ws.addRow([null, 5, null, 7]);
          row.getCell(2).note = 'five';
          const note = { texts: [{ font: { size: 12, color: { argb: 'FFFF6600' }, name: 'Calibri', scheme: 'minor' }, text: 'seven' }], ...box };
          row.getCell(4).note = note;
          row.commit();
          await wb.commit();
          const ws2 = (await readBack(filename)).getWorksheet('Hello');
          expect(ws2.getCell('B1').value).toBe(5);
          expect(ws2.getCell('B1').note).toBe('five');
          expect(ws2.getCell('D1').value).toBe(7);
          expect(ws2.getCell('D1').note.texts).toEqual([{ font: { size: 12, color: { argb: 'FFFF6600' }, name: 'Calibri' }, text: 'seven' }]);
          expect(ws2.getCell('D1').note.margins).toBeUndefined();
        });
      });
    }
    it.skip('with background image — not supported: images', () => {});
    it.skip('with background image where worksheet is commited in advance — not supported: images', () => {});

    it('with conditional formatting — not supported: conditional formatting', async () => {
      await withTempDir(async (dir) => {
        const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: join(dir, 'wb.test.xlsx') });
        const ws = wb.addWorksheet('Hello');
        await expectNotSupported(() =>
          ws.addConditionalFormatting({
            ref: 'A1:B2',
            rules: [{ type: 'expression', formulae: ['A1>1'], style: { font: { bold: true } } }],
          }),
        );
      });
    });

    it.skip('with conditional formatting that contains numFmt (#1814) — not supported: conditional formatting', () => {});
  });
});

// Added for xlvane (not upstream): exceljs behaviour checked against its source.
describe('WorkbookWriter defaults (added for xlvane)', () => {
  it('writes dates with built-in format 14 also without useStyles, and creator "ExcelJS" (X4-02, X4-03)', async () => {
    await withTempDir(async (dir) => {
      const filename = join(dir, 'wb.test.xlsx');
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename, created: new Date(Date.UTC(2024, 0, 1)) });
      wb.addWorksheet('Dates').addRow([new Date(Date.UTC(2024, 0, 31)), { formula: 'A1', result: new Date(Date.UTC(2024, 0, 31)) }]).commit();
      await wb.commit();
      const { readWorkbook } = await import('xlvane');
      const { readFile } = await import('node:fs/promises');
      const read = readWorkbook(await readFile(filename));
      const s = read.getSheet('Dates');
      expect(s.cell('A1')?.numFmt).toBe('mm-dd-yy');
      expect(s.cell('B1')?.numFmt).toBe('mm-dd-yy');
      expect(read.properties).toMatchObject({ creator: 'ExcelJS', lastModifiedBy: 'ExcelJS', created: new Date(Date.UTC(2024, 0, 1)), modified: new Date(Date.UTC(2024, 0, 1)) });
    });
  });

  it('a rich text run with vertAlign — not supported', async () => {
    const ws = new ExcelJS.Workbook().addWorksheet('S');
    ws.getCell('A1').value = { richText: [{ font: { vertAlign: 'superscript' }, text: '2' }] };
    await expectNotSupported(() => ws.workbook.xlsx.writeBuffer());
  });
});
