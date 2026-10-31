// Makes the LibreOffice part of the corpus (docs/decisions.md D-902): the content described in
// tests/corpus/REQUESTED.md is written with xlvane to a temporary file, opened by LibreOffice and
// saved again by LibreOffice's own XLSX export, so the committed files are written by LibreOffice.
// Usage: npm run build -w packages/xlvane && node scripts/make-lo-corpus.mjs
import { spawnSync } from 'node:child_process';
import { copyFileSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createWorkbook, writeWorkbook } from '../packages/xlvane/dist/index.js';

const root = fileURLToPath(new URL('..', import.meta.url));
const soffice = process.env.XLV_SOFFICE ?? 'soffice';
const d = (y, m, day, h = 0, min = 0) => new Date(Date.UTC(y, m - 1, day, h, min));

function basicTypes() {
  const wb = createWorkbook();
  const s = wb.addSheet('Data');
  ['Text', 'Number', 'Date', 'DateTime', 'Boolean', 'Formula'].forEach((h, i) => s.set(`${'ABCDEF'[i]}1`, h));
  s.set('A2', 'Alice').set('B2', 1234.5).set('C2', d(2024, 1, 31)).set('D2', d(2024, 1, 31, 13, 45), { numFmt: 'yyyy-mm-dd hh:mm' }).set('E2', true).set('F2', { formula: 'B2*2', result: 2469 });
  s.set('A3', '  spaces  ').set('B3', -0.001).set('C3', d(1900, 3, 1)).set('D3', d(2024, 2, 29), { numFmt: 'yyyy-mm-dd hh:mm' }).set('E3', false).set('F3', { formula: '1/0', result: { error: '#DIV/0!' } });
  s.set('A4', 'first\nsecond').set('B4', 1e20, { numFmt: '0.00E+00' }).set('C4', d(9999, 12, 31)).set('F4', { formula: 'A2&"!"', result: 'Alice!' });
  s.set('A5', 'zażółć 😀').set('B5', 0.1, { numFmt: '0.00%' }).set('F5', { formula: 'SUM(B2:B4)', result: 1234.499 + 1e20 });
  wb.addSheet('Second').set('A1', { formula: 'Data!B2', result: 1234.5 });
  return wb;
}

function formulasMergedHidden() {
  const wb = createWorkbook();
  const f = wb.addSheet('Formulas');
  for (let i = 1; i <= 10; i++) f.set(`A${i}`, i).set(`B${i}`, { formula: `A${i}*2`, result: i * 2 });
  f.set('C1', { formula: 'SUM(B1:B10)', result: 110 });
  const r = wb.addSheet('Report');
  r.set('A1', 'Title').merge('A1:D1');
  ['Region', 'Q1', 'Q2', 'Q3'].forEach((h, i) => r.set(`${'ABCD'[i]}2`, h));
  [['North', 1, 2, 3], ['South', 4, 5, 6], ['East', 7, 8, 9]].forEach((row, j) => row.forEach((v, i) => r.set(`${'ABCD'[i]}${String(j + 3)}`, v)));
  wb.addSheet('Hidden', { state: 'hidden' }).set('A1', 'hidden value');
  return wb;
}

function date1904() {
  const wb = createWorkbook({ dateSystem: '1904' });
  wb.addSheet('Dates').set('A1', d(1904, 1, 1)).set('A2', d(2024, 1, 31));
  return wb;
}

const work = mkdtempSync(join(tmpdir(), 'xlvane-lo-corpus-'));
const out = join(root, 'tests/corpus/libreoffice');
mkdirSync(out, { recursive: true });
try {
  for (const [name, wb] of [['basic-types', basicTypes()], ['formulas-merged-hidden', formulasMergedHidden()], ['date1904', date1904()]]) {
    mkdirSync(join(work, 'in'), { recursive: true });
    writeFileSync(join(work, 'in', `${name}.xlsx`), writeWorkbook(wb));
    const r = spawnSync(soffice, [`-env:UserInstallation=file://${join(work, 'profile')}`, '--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', join(work, 'out'), join(work, 'in', `${name}.xlsx`)], { encoding: 'utf8', timeout: 120_000 });
    if (r.status !== 0) throw new Error(`LibreOffice failed for ${name}: ${r.stderr}`);
    copyFileSync(join(work, 'out', `${name}.xlsx`), join(out, `${name}.xlsx`));
    console.log(`libreoffice/${name}.xlsx`);
  }
  const version = spawnSync(soffice, ['--version'], { encoding: 'utf8' }).stdout.trim();
  console.log(version);
} finally {
  rmSync(work, { recursive: true, force: true });
}
