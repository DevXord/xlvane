// LibreOffice re-saves a file with pictures written by xlvane and keeps them with their anchors
// (D-D03, TD-39). The images come from the Apache POI corpus (tests/corpus/LICENSES.md).
// Skipped when `soffice` is not installed, unless XLV_REQUIRE_LIBREOFFICE=1 (set in CI).
import { spawnSync } from 'node:child_process';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { readFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';

const SOFFICE = process.env['XLV_SOFFICE'] ?? 'soffice';
const available = spawnSync(SOFFICE, ['--version'], { encoding: 'utf8', timeout: 60_000 }).status === 0;
const required = process.env['XLV_REQUIRE_LIBREOFFICE'] === '1';
const media = (ext: string): Uint8Array =>
  Object.entries(unzipSync(readFileSync(new URL('../corpus/poi/WithDrawing.xlsx', import.meta.url)))).find(([n]) => n.startsWith('xl/media/') && n.endsWith(ext))![1];

describe.skipIf(!available && !required)('LibreOffice and pictures', () => {
  it('keeps the pictures, their images and their cell anchors when it re-saves the file', { timeout: 200_000 }, async () => {
    const dir = await mkdtemp(join(tmpdir(), 'xlvane-lo-images-'));
    try {
      const wb = createWorkbook();
      const png = wb.addImage(media('.png'));
      const jpeg = wb.addImage(media('.jpeg'));
      wb.addSheet('Pictures').set('A1', 'Logo').addImage(png, 'B2:D6', { description: 'Logo' }).addImage(jpeg, { from: 'F3', to: 'H10', editAs: 'oneCell' });
      const file = join(dir, 'pictures.xlsx');
      await writeFile(file, writeWorkbook(wb));
      const profile = pathToFileURL(join(dir, 'profile')).href;
      const result = spawnSync(SOFFICE, [`-env:UserInstallation=${profile}`, '--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', join(dir, 'out'), file], {
        encoding: 'utf8',
        timeout: 180_000,
      });
      expect(result.status, result.stderr).toBe(0);
      const back = readWorkbook(await readFile(join(dir, 'out', 'pictures.xlsx')));
      const images = back.getSheet('Pictures').images;
      expect(images.map((i) => i.image?.format)).toEqual(['png', 'jpeg']);
      expect(images[0]?.image?.bytes).toEqual(media('.png'));
      // LibreOffice 24.2.7.2 writes the bottom-right corner as an offset inside the last covered
      // cell (D6 + 611 280 EMU) instead of the start of the next one (E7 + 0); both are the same
      // place within a few EMU of rounding (TD-39). Positions in default-size cells, in cells:
      const at = (m: { col: number; row: number; colOffset: number; rowOffset: number }): [number, number] => [m.col + m.colOffset / 609_600, m.row + m.rowOffset / 190_500];
      const corners = images.map((i) => (i.anchor.type === 'twoCell' ? [...at(i.anchor.from), ...at(i.anchor.to)] : []));
      const expected = [
        [1, 1, 4, 6],
        [5, 2, 7, 9],
      ];
      corners.forEach((c, k) => c.forEach((v, n) => expect(v).toBeCloseTo(expected[k]![n]!, 1)));
      expect(images.map((i) => i.anchor.type === 'twoCell' && i.anchor.editAs)).toEqual(['twoCell', 'oneCell']);
      expect(images[0]?.description).toBe('Logo');
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});
