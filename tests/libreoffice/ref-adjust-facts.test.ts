// Re-observes how LibreOffice Calc adjusts references when rows or columns are inserted or
// deleted (docs/facts/ref-adjust.md). Files written by xlvane are changed through LibreOffice's UNO
// API (ref-adjust-facts.py: XTableRows/XTableColumns insertByIndex/removeByIndex), saved as XLSX
// and read back; the formulas, merges and defined names must equal the recorded observations
// (ref-adjust-facts.cases.ts). Skipped when soffice or Python UNO is missing, unless
// XLV_REQUIRE_LIBREOFFICE=1.
import { spawnSync } from 'node:child_process';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import type { Workbook } from '../../packages/xlvane/src/model';
import { COL_FORMULAS, OBSERVED, OTHER_FORMULAS, ROW_FORMULAS } from './ref-adjust-facts.cases';

const SOFFICE = process.env['XLV_SOFFICE'] ?? 'soffice';
const hasSoffice = spawnSync(SOFFICE, ['--version'], { encoding: 'utf8', timeout: 60_000 }).status === 0;
const hasUno = spawnSync('python3', ['-c', 'import uno'], { encoding: 'utf8', timeout: 60_000 }).status === 0;
const available = hasSoffice && hasUno;
const required = process.env['XLV_REQUIRE_LIBREOFFICE'] === '1';
const SCRIPT = fileURLToPath(new URL('./ref-adjust-facts.py', import.meta.url));

/** Where the formulas of ROW_FORMULAS sit on sheet Main: row 1 from column C, then row 30. */
function rowCell(i: number): { col: string; row: number } {
  return { col: String.fromCharCode(67 + (i % 20)), row: i < 20 ? 1 : 30 };
}

function inputs(): Record<string, Uint8Array> {
  const rows = createWorkbook();
  const main = rows.addSheet('Main');
  const other = rows.addSheet('Other');
  for (let r = 2; r <= 20; r++) main.set(`B${String(r)}`, r);
  ROW_FORMULAS.forEach((f, i) => {
    const { col, row } = rowCell(i);
    main.set(`${col}${String(row)}`, { formula: f });
  });
  OTHER_FORMULAS.forEach((f, i) => other.set(`${String.fromCharCode(65 + i)}1`, { formula: f }));
  rows.defineName('RangeAll', 'Main!$B$2:$B$10').defineName('RangeGone', 'Main!$B$5:$B$6').defineName('CellGone', 'Main!$B$5').defineName('RangePart', 'Main!$B$4:$B$6');
  main.merge('D5:E8').merge('F5:F6').merge('G4:G5').merge('H4:H8').merge('I6:I7');

  const cols = createWorkbook();
  const cmain = cols.addSheet('Main');
  cols.addSheet('Other');
  for (let c = 2; c <= 20; c++) cmain.set(`${String.fromCharCode(64 + c)}2`, c);
  COL_FORMULAS.forEach((f, i) => cmain.set(`A${String(i + 3)}`, { formula: f }));
  cols.defineName('RangeAll', 'Main!$B$2:$J$2').defineName('RangeGone', 'Main!$E$2:$F$2');
  cmain.merge('E4:H5').merge('E6:F6').merge('D7:E7');
  return { 'row.xlsx': writeWorkbook(rows), 'col.xlsx': writeWorkbook(cols) };
}

const JOBS = [
  { input: 'row.xlsx', output: 'row-insert-Main.xlsx', sheet: 'Main', axis: 'row', op: 'insert', at: 5, count: 2 },
  { input: 'row.xlsx', output: 'row-delete-Main.xlsx', sheet: 'Main', axis: 'row', op: 'delete', at: 5, count: 2 },
  { input: 'row.xlsx', output: 'row-insert-Other.xlsx', sheet: 'Other', axis: 'row', op: 'insert', at: 5, count: 2 },
  { input: 'col.xlsx', output: 'col-insert-Main.xlsx', sheet: 'Main', axis: 'col', op: 'insert', at: 5, count: 2 },
  { input: 'col.xlsx', output: 'col-delete-Main.xlsx', sheet: 'Main', axis: 'col', op: 'delete', at: 5, count: 2 },
] as const;

describe.skipIf(!available && !required)('LibreOffice: reference adjustment facts', () => {
  let dir = '';
  const results = new Map<string, Workbook>();

  beforeAll(async () => {
    dir = await mkdtemp(join(tmpdir(), 'xlvane-refadj-'));
    for (const [name, bytes] of Object.entries(inputs())) await writeFile(join(dir, name), bytes);
    await writeFile(join(dir, 'jobs.json'), JSON.stringify(JOBS));
    const run = spawnSync('python3', [SCRIPT, SOFFICE, dir, join(dir, 'jobs.json')], { encoding: 'utf8', timeout: 600_000 });
    if (run.status !== 0) throw new Error(`ref-adjust-facts.py failed: ${run.stderr}${run.error?.message ?? ''}`);
    for (const job of JOBS) results.set(job.output, readWorkbook(await readFile(join(dir, job.output))));
  }, 650_000);

  afterAll(async () => {
    if (dir) await rm(dir, { recursive: true, force: true });
  });

  it('formulas are adjusted as recorded (R-01 … R-13)', () => {
    const problems: string[] = [];
    const formulaAt = (wb: Workbook, sheet: string, ref: string): string | null => wb.getSheet(sheet).cell(ref)?.formula?.text ?? null;
    const counters = new Map<string, number>();
    for (const c of OBSERVED) {
      const key = `${c.axis}|${c.op}|${c.sheet}|${c.formulaSheet}`;
      const i = counters.get(key) ?? 0;
      counters.set(key, i + 1);
      const wb = results.get(`${c.axis}-${c.op}-${c.sheet}.xlsx`);
      if (!wb) throw new Error(`no result for ${key}`);
      let ref: string;
      if (c.axis === 'col') ref = `A${String(i + 3)}`;
      else if (c.formulaSheet === 'Other') ref = `${String.fromCharCode(65 + i)}1`;
      else {
        const { col, row } = rowCell(i);
        const moved = c.sheet === 'Main' && row >= 5 ? row + (c.op === 'insert' ? 2 : -2) : row;
        ref = `${col}${String(moved)}`;
      }
      const seen = formulaAt(wb, c.formulaSheet, ref);
      if (seen !== c.observed) problems.push(`${key} ${c.formula}: recorded ${c.observed}, LibreOffice now gives ${String(seen)}`);
    }
    expect(problems).toEqual([]);
  });

  it('defined names are adjusted like formulas (R-11)', () => {
    const names = (file: string): string[] => (results.get(file)?.definedNames ?? []).map((n) => `${n.name}=${n.refersTo}`).sort();
    expect(names('row-insert-Main.xlsx')).toEqual(['CellGone=Main!$B$7', 'RangeAll=Main!$B$2:$B$12', 'RangeGone=Main!$B$7:$B$8', 'RangePart=Main!$B$4:$B$8']);
    expect(names('row-delete-Main.xlsx')).toEqual(['CellGone=#REF!', 'RangeAll=Main!$B$2:$B$8', 'RangeGone=#REF!', 'RangePart=Main!$B$4:$B$4']);
    expect(names('row-insert-Other.xlsx')).toEqual(['CellGone=Main!$B$5', 'RangeAll=Main!$B$2:$B$10', 'RangeGone=Main!$B$5:$B$6', 'RangePart=Main!$B$4:$B$6']);
    expect(names('col-insert-Main.xlsx')).toEqual(['RangeAll=Main!$B$2:$L$2', 'RangeGone=Main!$G$2:$H$2']);
    expect(names('col-delete-Main.xlsx')).toEqual(['RangeAll=Main!$B$2:$H$2', 'RangeGone=#REF!']);
  });

  it('merged ranges grow, move, shrink or are dropped (R-14)', () => {
    const merges = (file: string): string[] => [...(results.get(file)?.getSheet('Main').merges ?? [])].sort();
    expect(merges('row-insert-Main.xlsx')).toEqual(['D7:E10', 'F7:F8', 'G4:G7', 'H4:H10', 'I8:I9']);
    // D5:E8 loses its top-left cell and is dropped; G4:G5 and I6:I7 shrink to one cell and are dropped.
    expect(merges('row-delete-Main.xlsx')).toEqual(['H4:H6']);
    expect(merges('col-insert-Main.xlsx')).toEqual(['D7:G7', 'G4:J5', 'G6:H6']);
    expect(merges('col-delete-Main.xlsx')).toEqual([]);
  });
});
