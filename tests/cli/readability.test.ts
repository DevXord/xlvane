// Readable and safe console output (docs/decisions.md D-R02): control characters from files are
// escaped, TERM=dumb has no colour, summaries go to stderr, warnings are noted, and the options
// added before 1.0.0 (--limit all, --range checked first, --max-columns) behave as documented.
import { chmod, stat, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { escapeControls, termWidth, terminalSafe, useColor } from '../../packages/xlvane/src/cli/output';
import { buildXlsx, inline, row } from '../helpers/xlsx-builder';
import { cli, withFixtures } from './helpers';

/** Any C0 control except tab and line feed, DEL, or C1. */
const CONTROL = /[\u0000-\u0008\u000b-\u001f\u007f-\u009f]/;
/** Every escape sequence the output may hold in colour: the SGR codes of Output. */
const OWN_SGR = /\u001b\[(?:1|2|22|31|32|33|36|39)m/g;

/** A cell whose text tries to set the window title and clear the screen (XML escapes _xHHHH_). */
const hostile = (): Uint8Array =>
  buildXlsx({
    sheets: [{ name: 'S', data: row(1, inline('A1', 'Name'), inline('B1', 'Note')) + row(2, inline('A2', 'x'), inline('B2', '_x001B_]0;pwned_x0007__x001B_[2J_x009B_')) }],
  });

/** A shared formula that cannot be expanded (contains "@"), which gives a warning on read. */
const withWarning = (): Uint8Array =>
  buildXlsx({
    sheets: [
      {
        name: 'S',
        data:
          row(1, inline('A1', 'Name'), inline('B1', 'Value')) +
          row(2, inline('A2', 'a'), '<c r="B2"><f t="shared" ref="B2:B3" si="0">@A2</f><v>1</v></c>') +
          row(3, inline('A3', 'b'), '<c r="B3"><f t="shared" si="0"/><v>2</v></c>'),
      },
    ],
  });

describe('terminal control characters', () => {
  it('are escaped as \\xHH, except tab and line feed', () => {
    expect(escapeControls('a\u001b[2Jb\u0007\u009b\u007f\tc\n')).toBe('a\\x1b[2Jb\\x07\\x9b\\x7f\tc\n');
    expect(terminalSafe('\u001b[1mok\u001b[22m \u001b[2J', true)).toBe('\u001b[1mok\u001b[22m \\x1b[2J');
    expect(terminalSafe('\u001b[1mok\u001b[22m', false)).toBe('\\x1b[1mok\\x1b[22m');
  });

  it('from cells never reach the terminal; --json keeps the text', async () => {
    await withFixtures(async (dir) => {
      await writeFile(join(dir, 'hostile.xlsx'), hostile());
      for (const command of [['read', 'hostile.xlsx'], ['inspect', 'hostile.xlsx']]) {
        const plain = await cli(command, { cwd: dir });
        expect(plain.code).toBe(0);
        expect(plain.stdout).not.toMatch(CONTROL);
        expect(plain.stdout).toContain('\\x1b]0;pwned\\x07');
        const coloured = await cli(command, { cwd: dir, env: {}, tty: true });
        expect(coloured.stdout.replace(OWN_SGR, '')).not.toMatch(CONTROL);
      }
      const json = await cli(['read', 'hostile.xlsx', '--json'], { cwd: dir });
      expect(JSON.parse(json.stdout).result.rows[0].Note).toBe('\u001b]0;pwned\u0007\u001b[2J\u009b');
    });
  });

  it('from arguments are escaped in error messages', async () => {
    await withFixtures(async (dir) => {
      const r = await cli(['read', 'report.xlsx', '--sheet', '\u001b[2JSales'], { cwd: dir, env: {}, tty: true });
      expect(r.code).toBe(3);
      expect(r.stderr.replace(OWN_SGR, '')).not.toMatch(CONTROL);
      expect(r.stderr).toContain('\\x1b[2JSales');
    });
  });
});

describe('terminal detection', () => {
  it('TERM=dumb has no colour unless FORCE_COLOR says so', () => {
    expect(useColor({ TERM: 'dumb' }, true)).toBe(false);
    expect(useColor({ TERM: 'dumb', FORCE_COLOR: '1' }, true)).toBe(true);
    expect(useColor({ TERM: 'xterm-256color' }, true)).toBe(true);
  });

  it('a terminal that reports width 0 gets COLUMNS or 100', () => {
    expect(termWidth({}, 0, true)).toBe(100);
    expect(termWidth({ COLUMNS: '72' }, 0, true)).toBe(72);
    expect(termWidth({ COLUMNS: '72' }, undefined, false)).toBe(100);
    expect(termWidth({}, 120, true)).toBe(120);
  });

  it('JSON is compact when piped and indented on a terminal', async () => {
    await withFixtures(async (dir) => {
      const piped = await cli(['read', 'report.xlsx', '--json'], { cwd: dir });
      expect(piped.stdout.trimEnd().split('\n')).toHaveLength(1);
      const tty = await cli(['read', 'report.xlsx', '--json'], { cwd: dir, tty: true });
      expect(tty.stdout.split('\n').length).toBeGreaterThan(5);
      expect(JSON.parse(tty.stdout)).toEqual(JSON.parse(piped.stdout));
    });
  });
});

describe('read', () => {
  it('--limit all shows every row, --range is checked before reading', async () => {
    await withFixtures(async (dir) => {
      const all = await cli(['read', 'report.xlsx', '--limit', 'all', '--json'], { cwd: dir });
      expect(JSON.parse(all.stdout).result).toMatchObject({ totalRows: 3, truncated: false, warnings: [] });
      const one = await cli(['read', 'report.xlsx', '--limit', '1'], { cwd: dir });
      expect(one.stderr).toContain('showing 1 of 3 rows. Use --limit all to see every row.');
      const range = await cli(['read', 'missing.xlsx', '--range', 'A1:ZZ'], { cwd: dir });
      expect(range.code).toBe(2);
      expect(range.stderr).toContain('The option "--range" must be a range such as A1:D20, not "A1:ZZ".');
    });
  });

  it('notes warnings on stderr and lists them in --json', async () => {
    await withFixtures(async (dir) => {
      await writeFile(join(dir, 'warn.xlsx'), withWarning());
      const text = await cli(['read', 'warn.xlsx'], { cwd: dir });
      expect(text.code).toBe(0);
      expect(text.stderr).toMatch(/^note: 1 warning while reading warn\.xlsx \(XLV_W_SHARED_FORMULA_NOT_EXPANDED\); run "xlvane validate warn\.xlsx" for details\.$/m);
      expect(text.stdout).not.toContain('note:');
      const json = await cli(['read', 'warn.xlsx', '--json'], { cwd: dir });
      expect(json.stderr).toBe('');
      expect(JSON.parse(json.stdout).result.warnings.map((w: { code: string }) => w.code)).toEqual(['XLV_W_SHARED_FORMULA_NOT_EXPANDED']);
      const validate = await cli(['validate', 'warn.xlsx'], { cwd: dir });
      expect(validate.code).toBe(0);
      expect(validate.stdout).toContain('Warnings (1):');
      expect(validate.stdout).toContain('OK with 1 warning: warn.xlsx. Use --strict to fail on warnings.');
    });
  });

  it('refuses standard input with a usage error', async () => {
    await withFixtures(async (dir) => {
      const r = await cli(['read', '-'], { cwd: dir });
      expect(r.code).toBe(2);
      expect(r.stderr).toContain('Reading from standard input ("-") is not supported.');
    });
  });
});

describe('writing and converting', () => {
  it('--force keeps the permissions of the replaced file', async () => {
    if (process.platform === 'win32') return;
    await withFixtures(async (dir) => {
      const out = join(dir, 'out.csv');
      await writeFile(out, 'old');
      await chmod(out, 0o640);
      const r = await cli(['convert', 'report.xlsx', 'out.csv', '--force'], { cwd: dir });
      expect(r.code).toBe(0);
      expect((await stat(out)).mode & 0o777).toBe(0o640);
    });
  });

  it('points out a CSV that uses semicolons', async () => {
    await withFixtures(async (dir) => {
      await writeFile(join(dir, 'semi.csv'), 'Region;Revenue\r\nNorth;1200\r\n');
      const r = await cli(['convert', 'semi.csv', 'semi.xlsx'], { cwd: dir });
      expect(r.code).toBe(0);
      expect(r.stderr).toContain('Try: xlvane convert semi.csv semi.xlsx --delimiter \';\'');
      const fine = await cli(['convert', 'semi.csv', 'semi2.xlsx', '--delimiter', ';'], { cwd: dir });
      expect(fine.stderr).toBe('');
    });
  });
});

describe('validate and inspect', () => {
  it('suggests the closest column type', async () => {
    await withFixtures(async (dir) => {
      const r = await cli(['validate', 'report.xlsx', '--columns', 'Revenue:numbr'], { cwd: dir });
      expect(r.code).toBe(2);
      expect(r.stderr).toContain('Did you mean "Revenue:number"?');
    });
  });

  it('--max-columns limits the columns and says how to see the rest', async () => {
    await withFixtures(async (dir) => {
      const r = await cli(['inspect', 'report.xlsx', '--sheet', 'Sales', '--max-columns', '1'], { cwd: dir });
      expect(r.stdout).toContain('… 2 more columns; use --max-columns 3 to see them');
      const bad = await cli(['inspect', 'report.xlsx', '--max-columns', '0'], { cwd: dir });
      expect(bad.code).toBe(2);
    });
  });

  it('hides built-in hidden names unless --verbose', async () => {
    await withFixtures(async (dir) => {
      const names = '<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">S!$A$1:$B$2</definedName><definedName name="Rate">S!$B$2</definedName></definedNames>';
      await writeFile(join(dir, 'names.xlsx'), buildXlsx({ sheets: [{ name: 'S', data: row(1, inline('A1', 'a')) }], workbookAfterSheets: names }));
      const plain = await cli(['inspect', 'names.xlsx'], { cwd: dir });
      expect(plain.stdout).toContain('Rate');
      expect(plain.stdout).not.toContain('_FilterDatabase');
      const verbose = await cli(['inspect', 'names.xlsx', '--verbose'], { cwd: dir });
      expect(verbose.stdout).toContain('_xlnm._FilterDatabase');
    });
  });
});
