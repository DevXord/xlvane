// xlvane migrate follows no symbolic links (a link loop made the scan exponential, and a link
// could make it read files outside the project), and --force changes the exceljs value in the
// dependency section that JSON.parse reads, not the first "exceljs" key anywhere in the file
// (1.1.1 security review, D-H01).
import { mkdir, readFile, symlink, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { cli, withFixtures } from './helpers';

const NEW = 'npm:@xlvane/compat@^1';

async function forced(dir: string, packageJson: string): Promise<{ text: string; parsed: Record<string, Record<string, string>> }> {
  await mkdir(join(dir, 'app'), { recursive: true });
  await writeFile(join(dir, 'app', 'package.json'), packageJson);
  const r = await cli(['migrate', 'app', '--force'], { cwd: dir });
  expect(r.code).toBe(0);
  const text = await readFile(join(dir, 'app', 'package.json'), 'utf8');
  return { text, parsed: JSON.parse(text) as Record<string, Record<string, string>> };
}

describe('migrate --force changes the right value', () => {
  it('leaves "exceljs" in overrides before the dependencies as it is', async () => {
    await withFixtures(async (dir) => {
      const { text, parsed } = await forced(dir, '{\n  "overrides": { "exceljs": "4.4.0" },\n  "dependencies": { "exceljs": "4.4.0" }\n}\n');
      expect(parsed['overrides']?.['exceljs']).toBe('4.4.0');
      expect(parsed['dependencies']?.['exceljs']).toBe(NEW);
      expect(text).toBe(`{\n  "overrides": { "exceljs": "4.4.0" },\n  "dependencies": { "exceljs": "${NEW}" }\n}\n`);
    });
  });

  it('changes the value JSON.parse reads when a section or key is repeated, or a key is escaped', async () => {
    await withFixtures(async (dir) => {
      const { parsed } = await forced(dir, '{"dependencies":{"exceljs":"1.0.0"},"devDependencies":{"exc\\u0065ljs":"4.4.0","exceljs":"4.3.0"},"dependencies":{"a":"1","exceljs":"4.4.0"}}');
      expect(parsed['dependencies']).toEqual({ a: '1', exceljs: NEW });
      expect(parsed['devDependencies']).toEqual({ exceljs: NEW });
    });
  });

  it('leaves a value that only looks the same in a string or another object alone', async () => {
    await withFixtures(async (dir) => {
      const { parsed } = await forced(dir, '{"description":"\\"exceljs\\": \\"4.4.0\\"","x":{"dependencies":{"exceljs":"4.4.0"}},"dependencies":{"exceljs":"4.4.0"}}');
      expect(parsed['description']).toBe('"exceljs": "4.4.0"');
      expect((parsed['x'] as unknown as Record<string, Record<string, string>>)['dependencies']?.['exceljs']).toBe('4.4.0');
      expect(parsed['dependencies']?.['exceljs']).toBe(NEW);
    });
  });
});

describe('migrate follows no symbolic links', () => {
  it('a directory linking to itself and a file linking outside the project are skipped', async () => {
    await withFixtures(async (dir) => {
      await mkdir(join(dir, 'outside'));
      await writeFile(join(dir, 'outside', 'secret.js'), "import ExcelJS from 'exceljs';\nnew ModelContainer();\n");
      await mkdir(join(dir, 'app', 'src'), { recursive: true });
      await writeFile(join(dir, 'app', 'package.json'), '{"dependencies":{"exceljs":"4.4.0"}}');
      await writeFile(join(dir, 'app', 'src', 'a.js'), "import ExcelJS from 'exceljs';\n");
      for (let i = 0; i < 8; i++) await symlink('..', join(dir, 'app', 'src', `loop${String(i)}`));
      await symlink(join(dir, 'outside', 'secret.js'), join(dir, 'app', 'src', 'link.js'));
      const start = performance.now();
      const r = await cli(['migrate', 'app', '--json'], { cwd: dir });
      expect(performance.now() - start).toBeLessThan(3000);
      const result = (JSON.parse(r.stdout) as { result: { filesScanned: number; findings: unknown[] } }).result;
      expect(result.filesScanned).toBe(1);
      expect(result.findings).toEqual([]);
    });
  });
});
