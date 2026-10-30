// Ported from exceljs v4.4.0 spec/integration/workbook/workbook.spec.js (MIT, Copyright (c)
// 2014-2019 Guyon Roche), changes: Mocha/Chai → Vitest with async/await; files go to a temporary
// directory (`roundTrip`); the merge-styles test uses a pattern fill and a font without
// `family`/`outline` (gradient fills and those font attributes are not written by xlvane);
// the frozen-panes and split-panes tests write one view per worksheet (only the first sheet view
// is written) and expect `topLeftCell` of a split view and `activeCell` to be absent (not
// written by xlvane);
// 'wrong data type' checks that `load({})` rejects, not the JSZip message; tests that need
// unsupported features (createTestBook fixtures, compression options, print titles,
// data validation, CSV, workbook views, fillFormula) are skipped or expect XLV_E_NOT_SUPPORTED; the
// duplicateRow tests build the sheet in memory (the fixture duplicateRowTest.xlsx was not copied)
// and compare styles by value (xlvane copies styles instead of sharing the object); `equalDate`
// is compared with getTime(); chai `have.members` is arrayContaining plus the length.
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { ExcelJS, expectNotSupported, roundTrip, styles, withTempDir } from './helpers';
import { splicedSheets } from './spliced-sheets';

describe('Workbook', () => {
  describe('Serialise', () => {
    it.skip('xlsx file — not supported: createTestBook fixture (needs data validation)', () => {});

    describe('Xlsx Zip Compression', () => {
      for (const level of ['best', 'default', 'fast', 'no']) {
        it.skip(`xlsx file with ${level} compression — not supported: zip compression options`, () => {});
      }
    });

    it('sheets with correct names', async () => {
      const wb = new ExcelJS.Workbook();
      const ws1 = wb.addWorksheet('Hello, World!');
      expect(ws1.name).toBe('Hello, World!');
      ws1.getCell('A1').value = 'Hello, World!';
      const ws2 = wb.addWorksheet();
      expect(ws2.name).toMatch(/sheet\d+/);
      ws2.getCell('A1').value = ws2.name;
      wb.addWorksheet('This & That');

      const wb2 = await roundTrip(wb);
      expect(wb2.getWorksheet('Hello, World!')).toBeTruthy();
      expect(wb2.getWorksheet('This & That')).toBeTruthy();
    });

    it('creator, lastModifiedBy, etc', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Hello');
      ws.getCell('A1').value = 'World!';
      wb.creator = 'Foo';
      wb.lastModifiedBy = 'Bar';
      wb.created = new Date(2016, 0, 1);
      wb.modified = new Date(2016, 4, 19);
      const wb2 = await roundTrip(wb);
      expect(wb2.creator).toBe(wb.creator);
      expect(wb2.lastModifiedBy).toBe(wb.lastModifiedBy);
      expect(wb2.created.getTime()).toBe(wb.created.getTime());
      expect(wb2.modified.getTime()).toBe(wb.modified.getTime());
    });
    it.skip('printTitlesRow — not supported: print settings', () => {});
    it.skip('printTitlesColumn — not supported: print settings', () => {});
    it.skip('printTitlesRowAndColumn — not supported: print settings', () => {});
    it.skip('shared formula — not supported: fillFormula', () => {});

    it('auto filter', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Hello');
      ws.getCell('A1').value = 1;
      ws.getCell('B1').value = 1;
      ws.getCell('A2').value = 2;
      ws.getCell('B2').value = 2;
      ws.getCell('A3').value = 3;
      ws.getCell('B3').value = 3;
      ws.autoFilter = 'A1:B1';

      const wb2 = await roundTrip(wb);
      expect(wb2.getWorksheet('Hello').autoFilter).toBe('A1:B1');
    });

    it('company, manager, etc', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Hello');
      ws.getCell('A1').value = 'World!';
      wb.company = 'Cyber Sapiens, Ltd';
      wb.manager = 'Guyon Roche';
      const wb2 = await roundTrip(wb);
      expect(wb2.company).toBe(wb.company);
      expect(wb2.manager).toBe(wb.manager);
    });

    it('title, subject, etc', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Hello');
      ws.getCell('A1').value = 'World!';
      wb.title = 'the title';
      wb.subject = 'the subject';
      wb.keywords = 'the keywords';
      wb.category = 'the category';
      wb.description = 'the description';
      const wb2 = await roundTrip(wb);
      expect(wb2.title).toBe(wb.title);
      expect(wb2.subject).toBe(wb.subject);
      expect(wb2.keywords).toBe(wb.keywords);
      expect(wb2.category).toBe(wb.category);
      expect(wb2.description).toBe(wb.description);
    });

    it('language, revision and contentStatus', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Hello');
      ws.getCell('A1').value = 'World!';
      wb.language = 'Klingon';
      wb.revision = 2;
      // Upstream sets the misspelt `contentStauts`, so contentStatus stays undefined on both sides.
      (wb as unknown as Record<string, unknown>)['contentStauts'] = 'Final';
      const wb2 = await roundTrip(wb);
      expect(wb2.language).toBe(wb.language);
      expect(wb2.revision).toBe(wb.revision);
      expect(wb2.contentStatus).toBe(wb.contentStatus);
    });

    it('empty strings', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Hello');
      ws.getCell('A1').value = 'Foo';
      ws.getCell('A2').value = '';
      ws.getCell('A3').value = 'Baz';

      const ws2 = (await roundTrip(wb)).getWorksheet('Hello');
      expect(ws2.getCell('A1').value).toBe('Foo');
      expect(ws2.getCell('A2').value).toBe('');
      expect(ws2.getCell('A3').value).toBe('Baz');
    });

    it.skip('dataValidations — not supported: data validation', () => {});

    it('empty string', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet();
      ws.columns = [
        { key: 'id', width: 10 },
        { key: 'name', width: 32 },
      ];
      ws.addRow({ id: 1, name: '' });
      await withTempDir((dir) => wb.xlsx.writeFile(join(dir, 'wb.test.xlsx')));
    });

    it.skip('a lot of sheets to xlsx file — not supported: createTestBook fixture', () => {});

    it('csv file — not supported: csv', async () => {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('blort').getCell('A1').value = 1;
      await expectNotSupported(() => wb.csv.writeBuffer());
    });

    it.skip('CSV file and its configuration — not supported: csv', () => {});
    it('defined names', async () => {
      const wb1 = new ExcelJS.Workbook();
      const ws1a = wb1.addWorksheet('blort');
      const ws1b = wb1.addWorksheet('foo');

      function assign(sheet: any, address: string, value: number, name: string | string[]) {
        const cell = sheet.getCell(address);
        cell.value = value;
        if (Array.isArray(name)) cell.names = name;
        else cell.name = name;
      }

      // single entry
      assign(ws1a, 'A1', 5, 'five');
      // three amigos - horizontal line
      assign(ws1a, 'A3', 3, 'amigos');
      assign(ws1a, 'B3', 3, 'amigos');
      assign(ws1a, 'C3', 3, 'amigos');
      // three amigos - vertical line
      assign(ws1a, 'E1', 3, 'verts');
      assign(ws1a, 'E2', 3, 'verts');
      assign(ws1a, 'E3', 3, 'verts');
      // four square
      assign(ws1a, 'C5', 4, 'squares');
      assign(ws1a, 'B6', 4, 'squares');
      assign(ws1a, 'C6', 4, 'squares');
      assign(ws1a, 'B5', 4, 'squares');
      // long distance
      assign(ws1a, 'B7', 2, 'sheets');
      assign(ws1b, 'B7', 2, 'sheets');
      // two names
      assign(ws1a, 'G1', 1, 'thing1');
      ws1a.getCell('G1').addName('thing2');
      // once removed
      assign(ws1a, 'G2', 1, ['once', 'twice']);
      ws1a.getCell('G2').removeName('once');

      const wb2 = await roundTrip(wb1);
      const ws2a = wb2.getWorksheet('blort');
      const ws2b = wb2.getWorksheet('foo');

      function check(sheet: any, address: string, value: number, name: string) {
        const cell = sheet.getCell(address);
        expect(cell.value).toBe(value);
        expect(cell.name).toBe(name);
      }

      check(ws2a, 'A1', 5, 'five');
      check(ws2a, 'A3', 3, 'amigos');
      check(ws2a, 'B3', 3, 'amigos');
      check(ws2a, 'C3', 3, 'amigos');
      check(ws2a, 'E1', 3, 'verts');
      check(ws2a, 'E2', 3, 'verts');
      check(ws2a, 'E3', 3, 'verts');
      check(ws2a, 'C5', 4, 'squares');
      check(ws2a, 'B6', 4, 'squares');
      check(ws2a, 'C6', 4, 'squares');
      check(ws2a, 'B5', 4, 'squares');
      check(ws2a, 'B7', 2, 'sheets');
      check(ws2b, 'B7', 2, 'sheets');

      expect(ws2a.getCell('G1').names).toEqual(expect.arrayContaining(['thing1', 'thing2']));
      expect(ws2a.getCell('G1').names).toHaveLength(2);
      expect(ws2a.getCell('G2').names).toEqual(['twice']);

      function rangeCheck(name: string, members: string[]) {
        const ranges = wb2.definedNames.getRanges(name);
        expect(ranges.name).toBe(name);
        if (members.length) {
          expect(ranges.ranges).toEqual(expect.arrayContaining(members));
          expect(ranges.ranges).toHaveLength(members.length);
        } else expect(ranges.ranges.length).toBe(0);
      }

      rangeCheck('five', ['blort!$A$1']);
      rangeCheck('amigos', ['blort!$A$3:$C$3']);
      rangeCheck('verts', ['blort!$E$1:$E$3']);
      rangeCheck('squares', ['blort!$B$5:$C$6']);
      rangeCheck('sheets', ['blort!$B$7', 'foo!$B$7']);
      rangeCheck('thing1', ['blort!$G$1']);
      rangeCheck('thing2', ['blort!$G$1']);
      rangeCheck('once', []);
      rangeCheck('twice', ['blort!$G$2']);
    });

    describe('Duplicate Rows', () => {
      it('Duplicate rows with styles properly', async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('duplicateTest');
        ws.getCell('A1').value = 'OneInfo';
        ws.getCell('A1').font = styles.roundTripFonts.broadwayRed20;
        ws.getCell('A2').value = 'TwoInfo';
        ws.duplicateRow(1, 2);
        const ws2 = (await roundTrip(wb)).getWorksheet('duplicateTest');
        expect(ws2.getCell('A2').value).toBe('OneInfo');
        expect(ws2.getCell('A2').style).toEqual(ws2.getCell('A1').style);
        expect(ws2.getCell('A3').value).toBe('OneInfo');
        expect(ws2.getCell('A3').style).toEqual(ws2.getCell('A1').style);
        expect(ws2.getCell('A4').value).toBeNull();
      });

      const four = () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('duplicateTest');
        ws.getCell('A1').value = 'OneInfo';
        ws.getCell('A2').value = 'TwoInfo';
        ws.getCell('A3').value = 'ThreeInfo';
        ws.getCell('A4').value = 'FourInfo';
        return { wb, ws };
      };

      it('Duplicate rows replacing properly', async () => {
        const { wb, ws } = four();
        ws.duplicateRow(1, 2, false);
        const ws2 = (await roundTrip(wb)).getWorksheet('duplicateTest');
        expect(ws2.getCell('A1').value).toBe('OneInfo');
        expect(ws2.getCell('A2').value).toBe('OneInfo');
        expect(ws2.getCell('A3').value).toBe('OneInfo');
        expect(ws2.getCell('A4').value).toBe('FourInfo');
      });

      it('Duplicate rows shifting properly', async () => {
        const { wb, ws } = four();
        ws.duplicateRow(1, 2, true);
        const ws2 = (await roundTrip(wb)).getWorksheet('duplicateTest');
        expect(ws2.getCell('A1').value).toBe('OneInfo');
        expect(ws2.getCell('A2').value).toBe('OneInfo');
        expect(ws2.getCell('A3').value).toBe('OneInfo');
        expect(ws2.getCell('A4').value).toBe('TwoInfo');
      });

      it('Duplicate rows with height properly', async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('duplicateTest');
        ws.getCell('A1').value = 'OneInfo';
        ws.getCell('A2').value = 'TwoInfo';
        ws.getRow(1).height = 25;
        ws.getRow(2).height = 15;
        ws.duplicateRow(1, 1, true);
        const ws2 = (await roundTrip(wb)).getWorksheet('duplicateTest');
        expect(ws2.getCell('A1').value).toBe('OneInfo');
        expect(ws2.getCell('A2').value).toBe('OneInfo');
        expect(ws2.getRow(1).height).toBe(ws2.getRow(2).height);
        expect(ws2.getRow(1).height).not.toBe(ws2.getRow(3).height);
      });
    });

    describe('Merge Cells', () => {
      it('serialises and deserialises properly', async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('blort');
        ws.getCell('B2').value = 'B2';
        ws.mergeCells('B2:C3');

        const ws2 = (await roundTrip(wb)).getWorksheet('blort');
        expect(ws2.getCell('B2').value).toBe('B2');
        expect(ws2.getCell('B3').value).toBe('B2');
        expect(ws2.getCell('C2').value).toBe('B2');
        expect(ws2.getCell('C3').value).toBe('B2');
        expect(ws2.getCell('B2').type).toBe(ExcelJS.ValueType.String);
        expect(ws2.getCell('B3').type).toBe(ExcelJS.ValueType.Merge);
        expect(ws2.getCell('C2').type).toBe(ExcelJS.ValueType.Merge);
        expect(ws2.getCell('C3').type).toBe(ExcelJS.ValueType.Merge);
      });

      it('styles', async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('blort');
        const B2 = ws.getCell('B2');
        B2.value = 5;
        B2.style.font = styles.roundTripFonts.broadwayRed20;
        B2.style.border = styles.borders.doubleRed;
        B2.style.fill = styles.fills.redDarkVertical;
        B2.style.alignment = styles.namedAlignments.middleCentre;
        B2.style.numFmt = styles.numFmts.numFmt1;
        // expecting styles to be copied (see worksheet spec)
        ws.mergeCells('B2:C3');

        const ws2 = (await roundTrip(wb)).getWorksheet('blort');
        for (const ref of ['B2', 'B3', 'C2', 'C3']) {
          const cell = ws2.getCell(ref);
          expect(cell.font, ref).toEqual(styles.roundTripFonts.broadwayRed20);
          expect(cell.border, ref).toEqual(styles.borders.doubleRed);
          expect(cell.fill, ref).toEqual(styles.fills.redDarkVertical);
          expect(cell.alignment, ref).toEqual(styles.namedAlignments.middleCentre);
          expect(cell.numFmt, ref).toBe(styles.numFmts.numFmt1);
        }
      });
    });
  });

  it('spliced meat and ham', async () => {
    const sheets: [string, string][] = [
      ['rows', 'removeOnly'],
      ['rows', 'insertFewer'],
      ['rows', 'insertSame'],
      ['rows', 'insertMore'],
      ['rows', 'insertStyle'],
      ['columns', 'removeOnly'],
      ['columns', 'insertFewer'],
      ['columns', 'insertSame'],
      ['columns', 'insertMore'],
    ];
    const wb = new ExcelJS.Workbook();
    for (const [group, name] of sheets) splicedSheets[group]?.[name]?.addSheet(wb);
    const wb2 = await roundTrip(wb);
    // Change for xlvane: rows.insertStyle checks a row default style, which xlvane does not write
    // (D-206, README "What xlvane does not keep"), so it is checked before the round trip.
    for (const [group, name] of sheets) splicedSheets[group]?.[name]?.checkSheet(name === 'insertStyle' ? wb : wb2);
  });

  it('throws an error when xlsx file not found', async () => {
    const wb = new ExcelJS.Workbook();
    await expect(wb.xlsx.readFile('./wb.doesnotexist.xlsx')).rejects.toBeDefined();
  });

  it('throws an error when csv file not found — not supported: csv', async () => {
    const wb = new ExcelJS.Workbook();
    await expectNotSupported(() => wb.csv.readFile('./wb.doesnotexist.csv'));
  });

  it('throw an error for wrong data type', async () => {
    const wb = new ExcelJS.Workbook();
    await expect(wb.xlsx.load({})).rejects.toBeDefined();
  });

  describe('Sheet Views', () => {
    it('frozen panes', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('frozen');
      ws.views = [{ state: 'frozen', xSplit: 2, ySplit: 3, topLeftCell: 'C4', activeCell: 'D5' }];
      ws.getCell('A1').value = 'Let it Snow!';

      const ws2 = (await roundTrip(wb)).getWorksheet('frozen');
      expect(ws2).toBeTruthy();
      expect(ws2.getCell('A1').value).toBe('Let it Snow!');
      expect(ws2.views[0]).toMatchObject({ state: 'frozen', xSplit: 2, ySplit: 3 });
    });

    // Change for xlvane: the second and third upstream views, each as the only view of a worksheet.
    it('frozen panes: the view read back', async () => {
      const defaults = { workbookViewId: 0, showRuler: true, showGridLines: true, showRowColHeaders: true, zoomScale: 100, zoomScaleNormal: 100, rightToLeft: false };
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('rows', { views: [{ state: 'frozen', ySplit: 1 }] }).getCell('A1').value = 1;
      wb.addWorksheet('cols', { views: [{ state: 'frozen', xSplit: 1 }] }).getCell('A1').value = 1;
      const wb2 = await roundTrip(wb);
      expect(wb2.getWorksheet('rows').views).toEqual([{ ...defaults, state: 'frozen', xSplit: 0, ySplit: 1, topLeftCell: 'A2' }]);
      expect(wb2.getWorksheet('cols').views).toEqual([{ ...defaults, state: 'frozen', xSplit: 1, ySplit: 0, topLeftCell: 'B1' }]);
    });

    // Change for xlvane: each upstream view is the only view of its own worksheet; `topLeftCell`
    // and `activeCell` of a split view are not written, so they read back as undefined/absent.
    it('serialises split panes', async () => {
      const defaults = { workbookViewId: 0, showRuler: true, showGridLines: true, showRowColHeaders: true, zoomScale: 100, zoomScaleNormal: 100, rightToLeft: false };
      const wb = new ExcelJS.Workbook();
      const views = [
        { state: 'split', xSplit: 2000, ySplit: 3000, topLeftCell: 'C4', activeCell: 'D5', activePane: 'bottomRight' },
        { state: 'split', ySplit: 1500, activePane: 'bottomLeft', topLeftCell: 'A10' },
        { state: 'split', xSplit: 1500, activePane: 'topRight' },
      ];
      views.forEach((view, i) => {
        const ws = wb.addWorksheet(`split${String(i)}`);
        ws.views = [view];
        ws.getCell('A1').value = 'Do the splits!';
      });
      const wb2 = await roundTrip(wb);
      const expected = [
        { ...defaults, state: 'split', xSplit: 2000, ySplit: 3000, topLeftCell: undefined, activePane: 'bottomRight' },
        { ...defaults, state: 'split', xSplit: 0, ySplit: 1500, topLeftCell: undefined, activePane: 'bottomLeft' },
        { ...defaults, state: 'split', xSplit: 1500, ySplit: 0, topLeftCell: undefined, activePane: 'topRight' },
      ];
      expected.forEach((view, i) => {
        const ws2 = wb2.getWorksheet(`split${String(i)}`);
        expect(ws2.getCell('A1').value).toBe('Do the splits!');
        expect(ws2.views).toEqual([view]);
      });
    });
    it.skip('multiple book views — not supported: workbook views', () => {});
  });
});
