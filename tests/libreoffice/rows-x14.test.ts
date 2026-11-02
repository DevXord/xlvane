// insertRows / deleteRows against LibreOffice for conditional formats with cell thresholds
// (cfvo@val, SR-12) and the [MS-XLSX] extension (x14 rules with xm:f and xm:sqref; SR-40 …
// SR-43): the same workbook is changed once by LibreOffice (ref-adjust-facts.py, UNO API) and once
// by xlvane; ranges, rule types and thresholds must come out the same (R-18 … R-20, D-F04).
// GUIDs, priorities and the data bar lengths that LibreOffice rewrites (SR-45) are not compared.
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

const SOFFICE = process.env['XLV_SOFFICE'] ?? 'soffice';
const hasSoffice = spawnSync(SOFFICE, ['--version'], { encoding: 'utf8', timeout: 60_000 }).status === 0;
const hasUno = spawnSync('python3', ['-c', 'import uno'], { encoding: 'utf8', timeout: 60_000 }).status === 0;
const required = process.env['XLV_REQUIRE_LIBREOFFICE'] === '1';
const SCRIPT = fileURLToPath(new URL('./ref-adjust-facts.py', import.meta.url));

function source(): Uint8Array {
  const wb = createWorkbook();
  const main = wb.addSheet('Main');
  const report = wb.addSheet('Report');
  for (let r = 1; r <= 20; r++) main.set(`B${String(r)}`, r).set(`D${String(r)}`, r).set(`E${String(r)}`, r).set(`G${String(r)}`, r);
  main.set('C12', 8).set('C15', 4);
  main.conditionalFormat('B2:B20', [{ type: 'colorScale', colorScale: { cfvo: [{ type: 'min' }, { type: 'formula', value: '$C$12' }], colors: ['F8696B', '63BE7B'] } }]);
  main.conditionalFormat('D2:D20', [{ type: 'dataBar', dataBar: { cfvo: [{ type: 'min' }, { type: 'formula', value: '$C$15' }], color: '638EC6', gradient: false, borderColor: '0000FF', axisColor: '000000' } }]);
  main.conditionalFormat('E5:E10 G3:G6', [{ type: 'iconSet', iconSet: { iconSet: '3Stars', cfvo: [{ type: 'percent', value: 0 }, { type: 'formula', value: '$C$12' }, { type: 'percent', value: 67 }] } }]);
  main.conditionalFormat('H5:H6', [{ type: 'iconSet', iconSet: { iconSet: '3Triangles', cfvo: [{ type: 'percent', value: 0 }, { type: 'percent', value: 33 }, { type: 'percent', value: 67 }] } }]);
  report.set('A1', 1);
  report.conditionalFormat('A1:A3', [{ type: 'iconSet', iconSet: { iconSet: '3Stars', cfvo: [{ type: 'percent', value: 0 }, { type: 'formula', value: 'Main!$C$15' }, { type: 'percent', value: 67 }] } }]);
  return writeWorkbook(wb);
}

/** Ranges, rule types, icon sets and thresholds, with sheet names written without quotes. */
function summary(wb: Workbook): Record<string, string[]> {
  const out: Record<string, string[]> = {};
  for (const name of ['Main', 'Report']) {
    out[name] = wb.getSheet(name).conditionalFormats.map((f) => {
      const r = f.rules[0]!;
      const cfvo = (r.colorScale ?? r.dataBar ?? r.iconSet)?.cfvo ?? [];
      return `${f.sqref} ${r.type}${r.iconSet ? ` ${r.iconSet.iconSet}` : ''} ${cfvo.map((v) => `${v.type}:${String(v.value).replace(/'(Main)'!/g, '$1!')}`).join(',')}`;
    });
  }
  return out;
}

describe.skipIf(!(hasSoffice && hasUno) && !required)('LibreOffice: row changes and extension conditional formats', () => {
  let dir = '';
  const lo = new Map<string, Workbook>();
  const jobs = [
    { input: 'in.xlsx', output: 'insert.xlsx', sheet: 'Main', axis: 'row', op: 'insert', at: 5, count: 2 },
    { input: 'in.xlsx', output: 'delete.xlsx', sheet: 'Main', axis: 'row', op: 'delete', at: 5, count: 2 },
    { input: 'in.xlsx', output: 'delete12.xlsx', sheet: 'Main', axis: 'row', op: 'delete', at: 12, count: 1 },
  ];

  beforeAll(async () => {
    dir = await mkdtemp(join(tmpdir(), 'xlvane-rows-x14-'));
    await writeFile(join(dir, 'in.xlsx'), source());
    await writeFile(join(dir, 'jobs.json'), JSON.stringify(jobs));
    const run = spawnSync('python3', [SCRIPT, SOFFICE, dir, join(dir, 'jobs.json')], { encoding: 'utf8', timeout: 600_000 });
    if (run.status !== 0) throw new Error(`ref-adjust-facts.py failed: ${run.stderr}${run.error?.message ?? ''}`);
    for (const job of jobs) lo.set(job.output, readWorkbook(await readFile(join(dir, job.output))));
  }, 650_000);

  afterAll(async () => {
    if (dir) await rm(dir, { recursive: true, force: true });
  });

  for (const job of jobs) {
    it(`${job.op} ${String(job.count)} row(s) at ${String(job.at)} gives what LibreOffice gives`, () => {
      const change = job.op === 'insert' ? insertRows : deleteRows;
      const ours = readWorkbook(change(source(), { sheet: job.sheet, at: job.at, count: job.count }).bytes);
      expect(summary(ours)).toEqual(summary(lo.get(job.output) as Workbook));
    });
  }
});
