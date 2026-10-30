// xlvane-specific tests of the streaming WorkbookWriter: sheet layout is written before the
// rows, so layout changes after the first committed row must throw rather than be lost (D-507).
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { readWorkbookFile } from '../../packages/xlvane/src/node';
import { ExcelJS, expectNotSupported, withTempDir } from './helpers';

describe('stream.xlsx.WorkbookWriter layout', () => {
  it('writes layout set before the first commit', async () => {
    await withTempDir(async (dir) => {
      const filename = join(dir, 'a.xlsx');
      const writer = new ExcelJS.stream.xlsx.WorkbookWriter({ filename });
      const sheet = writer.addWorksheet('S', { views: [{ state: 'frozen', ySplit: 1 }] });
      sheet.columns = [{ header: 'n', key: 'n', width: 12 }, { header: 'sq', key: 'sq' }];
      sheet.autoFilter = 'A1:B1';
      for (let n = 1; n <= 3; n++) sheet.addRow({ n, sq: n * n }).commit();
      sheet.commit();
      await writer.commit();
      const s = (await readWorkbookFile(filename)).getSheet('S');
      expect(s.cell('B4')?.value).toBe(9);
      expect(s.autoFilter).toBe('A1:B1');
      expect(s.freeze).toEqual({ rows: 1, cols: 0 });
      expect(s.columns[0]).toMatchObject({ min: 1, width: 12 });
    });
  });

  it('throws XLV_E_NOT_SUPPORTED for an autoFilter set after rows were committed', async () => {
    await withTempDir(async (dir) => {
      const writer = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: join(dir, 'b.xlsx') });
      const sheet = writer.addWorksheet('S');
      sheet.addRow([1, 2]).commit();
      sheet.autoFilter = 'A1:B1';
      await expectNotSupported(() => sheet.commit());
      await expect(writer.commit()).rejects.toMatchObject({ code: 'XLV_E_NOT_SUPPORTED' });
    });
  });

  it('throws XLV_E_NOT_SUPPORTED for merges added after rows were committed', async () => {
    const writer = new ExcelJS.stream.xlsx.WorkbookWriter({ stream: { write: (_c: unknown, cb: () => void) => cb(), end: (cb: () => void) => cb() } });
    const sheet = writer.addWorksheet('S');
    sheet.addRow([1, 2]).commit();
    await expectNotSupported(() => sheet.mergeCells('A2:B2'));
  });
});
