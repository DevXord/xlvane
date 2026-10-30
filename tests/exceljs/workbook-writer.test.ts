// Ported from exceljs v4.4.0 spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js
// (MIT, Copyright (c) 2014-2019 Guyon Roche), changes: Mocha/Chai → Vitest with async/await;
// files go to a temporary directory; cells are added with `addRow` instead of
// `worksheet.getCell` (the streaming worksheet subset is `addRow`, row style setters and
// `commit`); fonts without `family`/`outline`; the column and row style and width read-back is
// removed (no row or column default styles after reading, docs/decisions.md D-206); rich text,
// notes and conditional formatting expect XLV_E_NOT_SUPPORTED; tests that need
// createTestBook, shared formulas, defined names, data validation, images or zip options are
// skipped.
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
    it.skip('xlsx file — not supported: createTestBook fixture (needs rich text, notes, data validation)', () => {});
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
      });
    });

    it('rich text — not supported: rich text values on write', async () => {
      await withTempDir(async (dir) => {
        const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: join(dir, 'wb.test.xlsx'), useSharedStrings: true });
        const ws = wb.addWorksheet('Hello');
        await expectNotSupported(async () => {
          ws.addRow([{ richText: [{ font: { bold: true }, text: 'Hello' }, { text: ', World!' }] }]).commit();
          ws.commit();
          await wb.commit();
        });
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

    it.skip('defined names — not supported: defined names', () => {});

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

    it('writes notes — not supported: notes', async () => {
      await withTempDir(async (dir) => {
        const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: join(dir, 'wb.test.xlsx') });
        const ws = wb.addWorksheet('Hello');
        await expectNotSupported(() => {
          ws.addRow(['Hello, World!']).getCell(1).note = 'Here is a nice note';
        });
      });
    });

    it.skip('Cell annotation supports setting margins and protection properties — not supported: notes', () => {});
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
