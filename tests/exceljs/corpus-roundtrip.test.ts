// Load-then-save: every file of the POI corpus that @xlvane/compat loads must also be written by
// `xlsx.writeBuffer()` and read back (pre-release audit, docs/decisions.md D-R01). Files the layer
// refuses to load (encrypted, legacy formats, files beyond the limits, …) are outside this test;
// tests/corpus/poi.test.ts covers how the core reads them. Corpus licences: tests/corpus/LICENSES.md.
import { readdirSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { ExcelJS } from './helpers';

const DIR = fileURLToPath(new URL('../corpus/poi', import.meta.url));
const FILES = readdirSync(DIR).filter((f) => /\.xls[xm]$/.test(f)).sort();

describe('POI corpus: files that load also save', () => {
  it('has files to check', () => {
    expect(FILES.length).toBeGreaterThan(100);
  });

  it('writes and re-reads every loadable file', { timeout: 300_000 }, async () => {
    const failures: string[] = [];
    let loaded = 0;
    for (const file of FILES) {
      const bytes = await readFile(join(DIR, file));
      const wb = new ExcelJS.Workbook();
      try {
        await wb.xlsx.load(bytes);
      } catch {
        continue;
      }
      loaded++;
      try {
        const out = await wb.xlsx.writeBuffer();
        await new ExcelJS.Workbook().xlsx.load(out);
      } catch (error) {
        const e = error as { code?: string; message?: string; location?: unknown };
        failures.push(`${file}: ${String(e.code)} ${String(e.message)} ${JSON.stringify(e.location ?? {})}`);
      }
    }
    expect(loaded).toBeGreaterThan(250);
    expect(failures, failures.join('\n')).toEqual([]);
  });
});
