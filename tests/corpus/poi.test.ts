// The Apache POI corpus (tests/corpus/poi, Apache-2.0; docs/decisions.md D-901). For every file:
// - reading gives the outcome recorded in poi/expected.json (create or update it with
//   `npx vitest run tests/corpus -u` and review the diff): "ok" with the warning codes, or the
//   error code; nothing but an XlvError may ever be thrown, and reading takes at most 5 seconds;
// - for readable files: inspect works, the streaming reader returns the same cells as the batch
//   reader, readTable and toCSV work on the first worksheet, and editing one empty cell gives a
//   file that reads back with exactly that one change according to diff (files whose invalid
//   cell references were skipped are refused by editCells instead, D-903);
// - inserting one row before row 2 of the first worksheet and deleting it again gives a workbook
//   equal to the original according to diff, or a refusal with XLV_E_NOT_SUPPORTED (D-B04).
import { readdirSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { toCSV } from '../../packages/xlvane/src/csv';
import { diff } from '../../packages/xlvane/src/diff';
import { editCells } from '../../packages/xlvane/src/edit';
import { XlvError } from '../../packages/xlvane/src/errors';
import { inspect } from '../../packages/xlvane/src/inspect';
import type { Cell, Workbook } from '../../packages/xlvane/src/model';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { openWorkbookStream } from '../../packages/xlvane/src/stream-read';
import { deleteRows, insertRows } from '../../packages/xlvane/src/rows';
import { readTable } from '../../packages/xlvane/src/table';

const dir = new URL('./poi/', import.meta.url);
const licenses = readFileSync(fileURLToPath(new URL('./LICENSES.md', import.meta.url)), 'utf8');
const files = readdirSync(dir).filter((n) => /\.xlsx$/i.test(n)).sort();
const MAX_READ_MS = 5000;

type Outcome = { ok: true; warnings: string[] } | { ok: false; error: string };

function outcome(bytes: Uint8Array): { outcome: Outcome; workbook: Workbook | null } {
  try {
    const workbook = readWorkbook(bytes);
    return { outcome: { ok: true, warnings: [...new Set(workbook.warnings.map((w) => w.code))].sort() }, workbook };
  } catch (error) {
    if (!(error instanceof XlvError)) throw error;
    return { outcome: { ok: false, error: error.code }, workbook: null };
  }
}

function cellKey(c: Cell): string {
  const v = c.value instanceof Date ? c.value.toISOString() : c.value;
  return `${c.ref}|${c.type}|${JSON.stringify(v)}|${c.formula?.text ?? ''}`;
}

/** A cell address in the first worksheet that holds nothing, past the used columns. */
function emptyCell(workbook: Workbook): { sheet: string; cell: string } | null {
  const sheet = workbook.sheets.find((s) => s.kind === 'worksheet');
  if (!sheet) return null;
  let maxCol = 0;
  for (const c of sheet.cells()) maxCol = Math.max(maxCol, c.col);
  if (maxCol >= 16000) return null;
  let col = maxCol + 1;
  let letters = '';
  while (col > 0) {
    const r = (col - 1) % 26;
    letters = String.fromCharCode(65 + r) + letters;
    col = Math.floor((col - 1) / 26);
  }
  return { sheet: sheet.name, cell: `${letters}1` };
}

describe('corpus: Apache POI', () => {
  it('every file is listed in LICENSES.md, with LICENSE and NOTICE next to it', () => {
    expect(files.length).toBe(340);
    for (const name of files) expect(licenses, name).toContain(`| poi/${name} |`);
    expect(readdirSync(dir)).toEqual(expect.arrayContaining(['LICENSE', 'NOTICE', 'expected.json']));
  });

  it('reads every file with the recorded outcome', { timeout: 300_000 }, async () => {
    const outcomes: Record<string, Outcome> = {};
    const slow: string[] = [];
    for (const name of files) {
      const started = performance.now();
      outcomes[name] = outcome(readFileSync(new URL(name, dir))).outcome;
      const ms = performance.now() - started;
      if (ms > MAX_READ_MS) slow.push(`${name}: ${ms.toFixed(0)} ms`);
    }
    expect(slow).toEqual([]);
    await expect(`${JSON.stringify(outcomes, null, 2)}\n`).toMatchFileSnapshot('./poi/expected.json');
  });

  it('readable files: inspect, streaming, readTable, toCSV and a one-cell edit agree', { timeout: 600_000 }, async () => {
    const problems: string[] = [];
    for (const name of files) {
      const bytes = readFileSync(new URL(name, dir));
      const { workbook } = outcome(bytes);
      if (!workbook) continue;
      try {
        inspect(workbook);
        const stream = await openWorkbookStream(bytes);
        for (const sheet of workbook.sheets) {
          if (sheet.kind !== 'worksheet') continue;
          const batch = [...sheet.cells()].map(cellKey).sort();
          const streamed: string[] = [];
          for await (const row of stream.rows(sheet.name)) for (const c of row.cells) streamed.push(cellKey(c));
          streamed.sort();
          if (JSON.stringify(batch) !== JSON.stringify(streamed)) problems.push(`${name} [${sheet.name}]: streaming gives ${String(streamed.length)} cells, batch ${String(batch.length)}`);
        }
        const first = workbook.sheets.find((s) => s.kind === 'worksheet' && s.state === 'visible');
        if (first) {
          readTable(workbook, { sheet: first.name });
          toCSV(workbook, { sheet: first.name });
        }
        const target = emptyCell(workbook);
        if (target && workbook.warnings.some((w) => w.code === 'XLV_W_INVALID_CELL_REF')) {
          // Sheets with skipped invalid references cannot be edited safely (D-903).
          let code = '';
          try {
            editCells(bytes, [{ ...target, value: 'xlvane-edit' }]);
          } catch (error) {
            code = error instanceof XlvError ? error.code : 'CRASH';
          }
          if (code !== 'XLV_E_INVALID_CELL_REF') problems.push(`${name}: editing a sheet with invalid references gave ${code || 'no error'}`);
        } else if (target) {
          const edited = editCells(bytes, [{ ...target, value: 'xlvane-edit' }]);
          const d = diff(workbook, readWorkbook(edited.bytes));
          const changes = d.sheets.flatMap((s) => s.cells.map((c) => `${s.name}!${c.ref}:${c.kind}`));
          if (changes.join() !== `${target.sheet}!${target.cell}:added` || d.sheetsAdded.length + d.sheetsRemoved.length > 0) {
            problems.push(`${name}: the edit of ${target.sheet}!${target.cell} gives ${changes.slice(0, 5).join(', ') || 'no change'}`);
          }
        }
      } catch (error) {
        // Errors on readable files are recorded rather than thrown so that one file does not hide others.
        problems.push(`${name}: ${error instanceof XlvError ? error.code : 'CRASH'} ${error instanceof Error ? error.message.slice(0, 160) : String(error)}`);
      }
    }
    expect(problems).toEqual([]);
  });
  it('readable files: inserting and deleting a row gives the original back, or a refusal', { timeout: 600_000 }, async () => {
    const problems: string[] = [];
    let checked = 0;
    for (const name of files) {
      const bytes = readFileSync(new URL(name, dir));
      const { workbook } = outcome(bytes);
      if (!workbook || workbook.warnings.some((w) => w.code === 'XLV_W_INVALID_CELL_REF')) continue;
      const sheet = workbook.sheets.find((s) => s.kind === 'worksheet');
      if (!sheet) continue;
      try {
        const inserted = insertRows(bytes, { sheet: sheet.name, at: 2 }).bytes;
        const back = deleteRows(inserted, { sheet: sheet.name, at: 2 }).bytes;
        const d = diff(workbook, readWorkbook(back));
        const changes = d.sheets.flatMap((s) => s.cells.map((c) => `${s.name}!${c.ref}:${c.kind}`));
        // Reversed ranges such as O2:O1 come back written in order (O1:O2, R-15).
        if (!d.equal && !(name === 'bug60858.xlsx' && changes.join() === 'Arbeitszeitcheck_2.3!F2:changed,Arbeitszeitcheck_2.3!G2:changed')) problems.push(`${name}: ${changes.slice(0, 3).join(', ')}`);
        checked++;
      } catch (error) {
        const code = error instanceof XlvError ? error.code : 'CRASH';
        if (code !== 'XLV_E_NOT_SUPPORTED') problems.push(`${name}: ${code} ${error instanceof Error ? error.message.slice(0, 160) : String(error)}`);
      }
    }
    expect(problems).toEqual([]);
    // 223 files on 2026-10-30; the others have tables, drawings or similar parts (D-B04).
    expect(checked).toBeGreaterThan(200);
  });
});
