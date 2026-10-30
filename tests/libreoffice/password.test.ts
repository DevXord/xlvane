// Sheet passwords and LibreOffice 24.2.7.2 (docs/facts/sheet-rules.md SR-55, SR-58; D-C08):
// - LibreOffice protecting a sheet with "secret" writes the legacy verifier, which xlvane reads
//   and verifies (the hex digits are the high byte first);
// - LibreOffice accepts the strong SHA-512 verifier xlvane writes: unprotecting with a wrong
//   password keeps the sheet protected, with the right one it does not.
// Skipped when soffice or Python UNO is missing, unless XLV_REQUIRE_LIBREOFFICE=1.
import { spawnSync } from 'node:child_process';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { legacyPasswordHash, verifyPassword } from '../../packages/xlvane/src/password';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';

const SOFFICE = process.env['XLV_SOFFICE'] ?? 'soffice';
const hasSoffice = spawnSync(SOFFICE, ['--version'], { encoding: 'utf8', timeout: 60_000 }).status === 0;
const hasUno = spawnSync('python3', ['-c', 'import uno'], { encoding: 'utf8', timeout: 60_000 }).status === 0;
const required = process.env['XLV_REQUIRE_LIBREOFFICE'] === '1';
const SCRIPT = fileURLToPath(new URL('./password-facts.py', import.meta.url));

describe.skipIf(!required && !(hasSoffice && hasUno))('LibreOffice and sheet passwords', () => {
  it('LibreOffice writes the legacy verifier, and accepts xlvane strong verifiers', { timeout: 600_000 }, async () => {
    const dir = await mkdtemp(join(tmpdir(), 'xlvane-pw-'));
    try {
      const made = spawnSync('python3', [SCRIPT, SOFFICE, dir, 'secret'], { encoding: 'utf8', timeout: 300_000 });
      expect(made.status, made.stderr).toBe(0);
      const sheet = readWorkbook(await readFile(join(dir, 'protected.xlsx'))).sheets[0]!;
      expect(sheet.protection?.password?.toUpperCase()).toBe(legacyPasswordHash('secret'));
      expect(verifyPassword(sheet.protection!, 'secret')).toBe(true);
      expect(verifyPassword(sheet.protection!, 'Secret')).toBe(false);

      const wb = createWorkbook();
      wb.addSheet('S').set('A1', 1).protect({ password: 'secret' });
      await writeFile(join(dir, 'xlvane.xlsx'), writeWorkbook(wb));
      const checked = spawnSync('python3', [SCRIPT, SOFFICE, dir, 'secret', 'xlvane.xlsx'], { encoding: 'utf8', timeout: 300_000 });
      expect(checked.status, checked.stderr).toBe(0);
      expect(JSON.parse(checked.stdout.trim().split('\n').pop()!)).toEqual({ afterWrong: true, afterRight: false });
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});
