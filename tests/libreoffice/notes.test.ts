// LibreOffice re-saves a file with notes written by xlvane and keeps them, and xlvane reads the
// notes LibreOffice writes (D-A12). Skipped when `soffice` is not installed, unless
// XLV_REQUIRE_LIBREOFFICE=1 (set in CI).
import { spawnSync } from 'node:child_process';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { writeWorkbookStream } from '../../packages/xlvane/src/stream-write';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';

const SOFFICE = process.env['XLV_SOFFICE'] ?? 'soffice';
const available = spawnSync(SOFFICE, ['--version'], { encoding: 'utf8', timeout: 60_000 }).status === 0;
const required = process.env['XLV_REQUIRE_LIBREOFFICE'] === '1';

function resave(dir: string, file: string): void {
  const profile = pathToFileURL(join(dir, 'profile')).href;
  const result = spawnSync(SOFFICE, [`-env:UserInstallation=${profile}`, '--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', join(dir, 'out'), file], {
    encoding: 'utf8',
    timeout: 180_000,
  });
  expect(result.status, result.stderr).toBe(0);
}

describe.skipIf(!available && !required)('LibreOffice and notes', () => {
  it('keeps the notes of the batch and the streaming writer when it re-saves the file', { timeout: 300_000 }, async () => {
    const dir = await mkdtemp(join(tmpdir(), 'xlvane-lo-notes-'));
    try {
      const wb = createWorkbook();
      wb.addSheet('Budget')
        .set('B2', 1250)
        .note('B2', 'Estimate from March')
        .note('D4', { runs: [{ text: 'Ann: ', font: { bold: true } }, { text: 'empty cell' }], author: 'Ann' });
      await writeFile(join(dir, 'batch.xlsx'), writeWorkbook(wb));
      const streamed = new Uint8Array(
        await new Response(writeWorkbookStream({ sheets: [{ name: 'Data', rows: [{ values: ['id'], notes: ['Identifier'] }, [1]] }] })).arrayBuffer(),
      );
      await writeFile(join(dir, 'stream.xlsx'), streamed);
      resave(dir, join(dir, 'batch.xlsx'));
      resave(dir, join(dir, 'stream.xlsx'));

      const budget = readWorkbook(await readFile(join(dir, 'out', 'batch.xlsx'))).getSheet('Budget');
      expect(budget.cell('B2')?.value).toBe(1250);
      const byRef = new Map(budget.notes.map((n) => [n.ref, n]));
      expect(byRef.get('B2')?.text).toBe('Estimate from March');
      expect(byRef.get('D4')?.text).toBe('Ann: empty cell');
      expect(budget.cell('B2')?.note?.text).toBe('Estimate from March');
      const data = readWorkbook(await readFile(join(dir, 'out', 'stream.xlsx'))).getSheet('Data');
      expect(data.notes.map((n) => [n.ref, n.text])).toEqual([['A1', 'Identifier']]);
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it('keeps a note with a custom box, and xlvane reads the box LibreOffice writes (D-E07)', { timeout: 300_000 }, async () => {
    const dir = await mkdtemp(join(tmpdir(), 'xlvane-lo-note-box-'));
    try {
      const wb = createWorkbook();
      wb.addSheet('Box').set('A1', 1).note('A1', { text: 'Editable', box: { margins: { left: 2, top: 1.5, right: 2, bottom: 1 }, lockText: false } });
      await writeFile(join(dir, 'box.xlsx'), writeWorkbook(wb));
      resave(dir, join(dir, 'box.xlsx'));
      const out = await readFile(join(dir, 'out', 'box.xlsx'));
      const note = readWorkbook(out).getSheet('Box').notes[0]!;
      expect(note.text).toBe('Editable');
      // Observed with LibreOffice 24.2.7.2 (W2-69): it keeps the note and writes its own box, with
      // no inset, its own inset mode and the text locked again; xlvane reads that box.
      expect(note.box).toEqual({ autoMargins: false, margins: null, locked: true, lockText: true, anchor: 'absolute' });
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});
