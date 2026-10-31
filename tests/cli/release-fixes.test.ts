// CLI fixes from testing the packed package in outside projects before 1.0.0 (docs/decisions.md
// D-R08): --dry-run of an in-place edit, "Try:" suggestions that fix the right argument, a CLI
// hint for a JSON object that is not an array, and bidirectional-text controls in output.
import { copyFile, readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { commandLine, escapeControls } from '../../packages/xlvane/src/cli/output';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { cli, withFixtures } from './helpers';

const tryLine = (stderr: string): string | undefined => /^Try: +(.*)$/m.exec(stderr)?.[1];

describe('edit --dry-run in place (M1)', () => {
  it('fails like the real run when neither --output nor --force is given', async () => {
    await withFixtures(async (dir) => {
      const before = await readFile(join(dir, 'report.xlsx'));
      const dry = await cli(['edit', 'report.xlsx', '--set', 'Sales!B2=1', '--dry-run'], { cwd: dir });
      const real = await cli(['edit', 'report.xlsx', '--set', 'Sales!B2=1'], { cwd: dir });
      expect(real.code).toBe(3);
      expect(dry.code).toBe(real.code);
      expect(dry.stderr).toContain('XLV_E_FILE_EXISTS');
      const json = JSON.parse((await cli(['edit', 'report.xlsx', '--set', 'Sales!B2=1', '--dry-run', '--json'], { cwd: dir })).stdout) as { ok: boolean; error: { code: string } };
      expect(json.ok).toBe(false);
      expect(json.error.code).toBe('XLV_E_FILE_EXISTS');
      expect(await readFile(join(dir, 'report.xlsx'))).toEqual(before);
    });
  });

  it('with --force reports the replacement and writes nothing', async () => {
    await withFixtures(async (dir) => {
      const before = await readFile(join(dir, 'report.xlsx'));
      const r = await cli(['edit', 'report.xlsx', '--set', 'Sales!B2=1', '--dry-run', '--force'], { cwd: dir });
      expect(r.code).toBe(0);
      expect(r.stdout).toContain('Dry run: would replace report.xlsx');
      expect(await readFile(join(dir, 'report.xlsx'))).toEqual(before);
    });
  });
});

describe('"Try:" suggestions (M2)', () => {
  it('a wrong sheet inside --set is fixed inside --set, not with an extra --sheet', async () => {
    await withFixtures(async (dir) => {
      const r = await cli(['edit', 'report.xlsx', '--set', 'Sales9!B2=1', '-o', 'x.xlsx'], { cwd: dir });
      expect(r.stderr).toContain('XLV_E_SHEET_NOT_FOUND');
      expect(tryLine(r.stderr)).toBe(commandLine(['edit', 'report.xlsx', '--set', 'Sales!B2=1', '-o', 'x.xlsx']));
      const json = JSON.parse((await cli(['edit', 'report.xlsx', '--set=Sales9!B2=1', '-o', 'x.xlsx', '--json'], { cwd: dir })).stdout) as { try: string | null };
      expect(json.try).toBe(commandLine(['edit', 'report.xlsx', '--set=Sales!B2=1', '-o', 'x.xlsx', '--json']));
      // The suggested command works.
      expect((await cli(['edit', 'report.xlsx', '--set', 'Sales!B2=1', '-o', 'x.xlsx'], { cwd: dir })).code).toBe(0);
    });
  });

  it('a quoted sheet name in --set stays quoted', async () => {
    await withFixtures(async (dir) => {
      const wb = createWorkbook();
      wb.addSheet('My Sales').set('A1', 1);
      await writeFile(join(dir, 'sp.xlsx'), writeWorkbook(wb));
      const r = await cli(['edit', 'sp.xlsx', '--set', "'My Sale'!A1=2", '-o', 'x.xlsx'], { cwd: dir });
      expect(tryLine(r.stderr)).toBe(commandLine(['edit', 'sp.xlsx', '--set', "'My Sales'!A1=2", '-o', 'x.xlsx']));
    });
  });

  it('--sheet is still corrected when the sheet came from --sheet', async () => {
    await withFixtures(async (dir) => {
      const r = await cli(['read', 'report.xlsx', '--sheet', 'Sale'], { cwd: dir });
      expect(tryLine(r.stderr)).toBe('xlvane read report.xlsx --sheet Sales');
    });
  });

  it('every --set that names the wrong sheet is fixed', async () => {
    await withFixtures(async (dir) => {
      const r = await cli(['edit', 'report.xlsx', '--set', 'Sales9!B2=1', '--set', 'Sales9!B3=1', '-o', 'x.xlsx'], { cwd: dir });
      expect(tryLine(r.stderr)).toBe(commandLine(['edit', 'report.xlsx', '--set', 'Sales!B2=1', '--set', 'Sales!B3=1', '-o', 'x.xlsx']));
    });
  });

  it('a missing file is replaced only where it was read, and never by another argument of the command', async () => {
    await withFixtures(async (dir) => {
      await copyFile(join(dir, 'report.xlsx'), join(dir, 'sp.xlsx'));
      const diff = await cli(['diff', 'sp.xlsx', 'sp4.xlsx'], { cwd: dir });
      expect(diff.stderr).toContain('File not found: "sp4.xlsx"');
      expect(diff.stderr).not.toContain('Did you mean "sp.xlsx"');
      expect(tryLine(diff.stderr)).not.toBe('xlvane diff sp.xlsx sp.xlsx');
      const edit = await cli(['edit', 'sp4.xlsx', '--set', 'Sales!B2=1', '-o', 'sp4.xlsx'], { cwd: dir });
      expect(edit.stderr).toContain('Did you mean "sp.xlsx"');
      expect(tryLine(edit.stderr)).toBe(commandLine(['edit', 'sp.xlsx', '--set', 'Sales!B2=1', '-o', 'sp4.xlsx']));
    });
  });

  it('only files the command can read are suggested', async () => {
    await withFixtures(async (dir) => {
      await writeFile(join(dir, 'long.csv'), 'a,b\r\n1,2\r\n');
      const r = await cli(['validate', 'long.xlsx', '--columns', 'a:number'], { cwd: dir });
      expect(r.stderr).toContain('File not found: "long.xlsx"');
      expect(r.stderr).not.toContain('long.csv');
      expect(tryLine(r.stderr) ?? '').not.toContain('.csv');
      // The other way round: a CSV input gets a CSV suggestion.
      const c = await cli(['convert', 'sale.csv', 'x.xlsx'], { cwd: dir });
      expect(tryLine(c.stderr)).toBe('xlvane convert sales.csv x.xlsx');
    });
  });
});

describe('create with a JSON object that is not an array (L1)', () => {
  it('names the property that holds the rows and how to extract it', async () => {
    await withFixtures(async (dir) => {
      await writeFile(join(dir, 'wrapped.json'), JSON.stringify({ total: 2, rows: [{ a: 1 }, { a: 2 }] }));
      const r = await cli(['create', 'wrapped.json', 'out.xlsx'], { cwd: dir });
      expect(r.code).toBe(3);
      expect(r.stderr).not.toContain('e.g. data.rows');
      expect(r.stderr).toContain('"rows"');
      expect(r.stderr).toContain("jq '.rows' wrapped.json > rows.json");
      await writeFile(join(dir, 'plain.json'), JSON.stringify({ a: 1 }));
      const p = await cli(['create', 'plain.json', 'out.xlsx'], { cwd: dir });
      expect(p.stderr).not.toContain('e.g. data.rows');
      expect(p.stderr).toContain('array of objects');
    });
  });
});

describe('bidirectional text controls (L3)', () => {
  it('are escaped in human-readable output', () => {
    for (const c of ['‪', '‫', '‬', '‭', '‮', '⁦', '⁧', '⁨', '⁩', '‎', '‏', '؜']) {
      expect(escapeControls(`a${c}b`)).toBe(`a\\u${c.charCodeAt(0).toString(16).padStart(4, '0')}b`);
    }
  });

  it('are exactly the characters with the Unicode property Bidi_Control', () => {
    const bidi = /\p{Bidi_Control}/u;
    for (let c = 0; c <= 0xffff; c++) {
      const ch = String.fromCharCode(c);
      if (c < 0x20 || (c >= 0x7f && c <= 0x9f)) continue;
      expect(escapeControls(ch) !== ch, `U+${c.toString(16)}`).toBe(bidi.test(ch));
    }
  });

  it('do not reach the terminal from cells, sheet names and notes; --json keeps them', async () => {
    await withFixtures(async (dir) => {
      const wb = createWorkbook();
      wb.addSheet('Data‮').set('A1', 'Name').set('A2', 'abc‮gnp.exe').note('A2', 'x⁧y');
      await writeFile(join(dir, 'bidi.xlsx'), writeWorkbook(wb));
      for (const argv of [['read', 'bidi.xlsx'], ['inspect', 'bidi.xlsx']]) {
        const r = await cli(argv, { cwd: dir });
        expect(r.stdout + r.stderr).not.toMatch(/[‪-‮⁦-⁩]/);
        expect(r.stdout).toContain('\\u202e');
      }
      const json = await cli(['read', 'bidi.xlsx', '--json'], { cwd: dir });
      expect(json.stdout).toContain('‮');
    });
  });
});
