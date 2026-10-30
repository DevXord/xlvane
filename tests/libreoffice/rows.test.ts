// insertRows / deleteRows against LibreOffice (docs/decisions.md D-B04): the same workbook is
// changed once by LibreOffice's own row insertion or deletion (ref-adjust-facts.py, UNO API) and
// once by xlvane; formulas on both sheets, merged ranges and defined names must come out the same.
// Skipped when soffice or Python UNO is missing, unless XLV_REQUIRE_LIBREOFFICE=1.
import { spawnSync } from 'node:child_process';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import type { Workbook } from '../../packages/xlvane/src/model';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { deleteRows, insertRows } from '../../packages/xlvane/src/rows';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { OTHER_FORMULAS, ROW_FORMULAS } from './ref-adjust-facts.cases';

const SOFFICE = process.env['XLV_SOFFICE'] ?? 'soffice';
const hasSoffice = spawnSync(SOFFICE, ['--version'], { encoding: 'utf8', timeout: 60_000 }).status === 0;
const hasUno = spawnSync('python3', ['-c', 'import uno'], { encoding: 'utf8', timeout: 60_000 }).status === 0;
const required = process.env['XLV_REQUIRE_LIBREOFFICE'] === '1';
const SCRIPT = fileURLToPath(new URL('./ref-adjust-facts.py', import.meta.url));

// 3-D references that cover the changed sheet are refused by xlvane (D-B01), so they are left out.
const MAIN = ROW_FORMULAS.filter((f) => !f.includes(':Other!'));

function source(): Uint8Array {
  const wb = createWorkbook();
  const main = wb.addSheet('Main');
  const other = wb.addSheet('Other');
  for (let r = 2; r <= 20; r++) main.set(`B${String(r)}`, r);
  MAIN.forEach((f, i) => main.set(`${String.fromCharCode(67 + (i % 20))}${i < 20 ? '1' : '30'}`, { formula: f }));
  OTHER_FORMULAS.forEach((f, i) => other.set(`${String.fromCharCode(65 + i)}1`, { formula: f }));
  wb.defineName('RangeAll', 'Main!$B$2:$B$10').defineName('RangeGone', 'Main!$B$5:$B$6').defineName('CellGone', 'Main!$B$5').defineName('RangePart', 'Main!$B$4:$B$6');
  main.merge('D5:E8').merge('F5:F6').merge('G4:G5').merge('H4:H8').merge('I6:I7');
  return writeWorkbook(wb);
}

/** Formulas, merges and names, with sheet names written without quotes (LibreOffice drops needless quotes). */
function summary(wb: Workbook): Record<string, string[]> {
  const unquote = (s: string): string => s.replace(/'(Main|Other)'!/g, '$1!');
  const out: Record<string, string[]> = {};
  for (const name of ['Main', 'Other']) {
    out[name] = [...wb.getSheet(name).cells()].filter((c) => c.formula !== null).map((c) => `${c.ref}=${unquote(c.formula?.text ?? '')}`);
  }
  out['merges'] = [...wb.getSheet('Main').merges].sort();
  out['names'] = wb.definedNames.map((n) => `${n.name}=${n.refersTo}`).sort();
  return out;
}

describe.skipIf(!(hasSoffice && hasUno) && !required)('LibreOffice: insertRows and deleteRows', () => {
  let dir = '';
  const lo = new Map<string, Workbook>();

  beforeAll(async () => {
    dir = await mkdtemp(join(tmpdir(), 'xlvane-rows-'));
    await writeFile(join(dir, 'in.xlsx'), source());
    const jobs = [
      { input: 'in.xlsx', output: 'insert.xlsx', sheet: 'Main', axis: 'row', op: 'insert', at: 5, count: 2 },
      { input: 'in.xlsx', output: 'delete.xlsx', sheet: 'Main', axis: 'row', op: 'delete', at: 5, count: 2 },
    ];
    await writeFile(join(dir, 'jobs.json'), JSON.stringify(jobs));
    const run = spawnSync('python3', [SCRIPT, SOFFICE, dir, join(dir, 'jobs.json')], { encoding: 'utf8', timeout: 600_000 });
    if (run.status !== 0) throw new Error(`ref-adjust-facts.py failed: ${run.stderr}${run.error?.message ?? ''}`);
    for (const job of jobs) lo.set(job.op, readWorkbook(await readFile(join(dir, job.output))));
  }, 650_000);

  afterAll(async () => {
    if (dir) await rm(dir, { recursive: true, force: true });
  });

  it('insertRows gives the formulas, merges and names LibreOffice gives', () => {
    const ours = readWorkbook(insertRows(source(), { sheet: 'Main', at: 5, count: 2 }).bytes);
    expect(summary(ours)).toEqual(summary(lo.get('insert') as Workbook));
  });

  it('deleteRows gives the formulas, merges and names LibreOffice gives', () => {
    const ours = readWorkbook(deleteRows(source(), { sheet: 'Main', at: 5, count: 2 }).bytes);
    expect(summary(ours)).toEqual(summary(lo.get('delete') as Workbook));
  });
});
