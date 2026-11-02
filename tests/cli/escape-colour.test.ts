// With colours on, the final filter keeps the colour sequences the CLI writes itself, so text
// from a file holding the same bytes (ESC [32m from _x001B_) passed through raw. Text from files
// is now escaped where it is inserted (1.1.1 security review, D-H01).
import { writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { buildXlsx } from '../helpers/xlsx-builder';
import { cli, withFixtures } from './helpers';

const E = '_x001B_[32m';
const raw = 'Q\u001b[32m';
const escaped = 'Q\\x1b[32m';

function hostile(cell: string): Uint8Array {
  return buildXlsx({
    sheets: [
      {
        name: `SQ${E}Z`,
        data: `<row r="1"><c r="A1" t="inlineStr"><is><t>hQ${E}Z</t></is></c><c r="B1" t="inlineStr"><is><t>b</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>${cell}</t></is></c><c r="B2" t="e"><v>#N/A</v></c></row>`,
      },
      { name: 'Two', data: '' },
    ],
    workbookAfterSheets: `<definedNames><definedName name="NQ">'SQ${E}Z'!$A$1</definedName></definedNames>`,
  });
}

describe('text from files with colours on', () => {
  const env = { FORCE_COLOR: '1' };
  for (const argv of [
    ['inspect', 'h.xlsx'],
    ['inspect', 'h.xlsx', '--sheet', `SQ\u001b[32mZ`],
    ['read', 'h.xlsx'],
    ['read', 'h.xlsx', '--sheet', `SQ\u001b[32mZ`],
    ['read', 'h.xlsx', '--sheet', 'nope'],
    ['diff', 'h.xlsx', 'h2.xlsx'],
    ['validate', 'h.xlsx'],
    ['convert', 'h.xlsx', 'out.csv'],
    ['convert', 'h.xlsx', 'out.csv', '--sheet', 'nope'],
  ]) {
    it(`xlvane ${argv.join(' ')} escapes it`, async () => {
      await withFixtures(async (dir) => {
        await writeFile(join(dir, 'h.xlsx'), hostile(`cQ${E}Z`));
        await writeFile(join(dir, 'h2.xlsx'), hostile(`dQ${E}Z`));
        const r = await cli(argv, { cwd: dir, env, tty: true });
        const all = r.stdout + r.stderr;
        expect(all).not.toContain(raw);
        // validate and a successful convert print nothing from the file.
        if (argv[0] !== 'validate' && !(argv[0] === 'convert' && argv.length === 3)) expect(all).toContain(escaped);
      });
    });
  }
});

describe('--json', () => {
  it('writes DEL, C1 and bidirectional characters from the file as \\u escapes, with the same value', async () => {
    await withFixtures(async (dir) => {
      await writeFile(join(dir, 'h.xlsx'), hostile('x_x009B_31m_x202E__x2028__x007F_y'));
      for (const tty of [false, true]) {
        const r = await cli(['read', 'h.xlsx', '--json', '--header', 'none'], { cwd: dir, tty });
        expect(r.stdout).not.toMatch(/[\u007f-\u009f\u202e\u2028]/);
        expect(r.stdout).toContain('x\\u009b31m\\u202e\\u2028\\u007fy');
        const parsed = JSON.parse(r.stdout) as { result: unknown };
        expect(JSON.stringify(parsed.result)).toContain('x\u009b31m\u202e\u2028\u007fy');
      }
    });
  });
});
