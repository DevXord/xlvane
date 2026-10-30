// Runs inspect() on every workbook in the corpus and compares the result with the JSON snapshot
// stored next to the file (`<name>.inspect.json`; create or update with `npx vitest run -u`).
// Files that xlvane must reject (encrypted workbooks, legacy .xls) snapshot their error instead.
// Every file must be listed in tests/corpus/LICENSES.md; files still to be supplied are listed in
// tests/corpus/REQUESTED.md.
import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { XlvError } from '../../packages/xlvane/src/errors';
import { inspect } from '../../packages/xlvane/src/inspect';

const SOURCES = ['excel', 'libreoffice', 'google-sheets', 'other'] as const;
const root = new URL('./', import.meta.url);
const licenses = readFileSync(fileURLToPath(new URL('LICENSES.md', root)), 'utf8');

for (const source of SOURCES) {
  const dir = new URL(`${source}/`, root);
  const files = existsSync(dir)
    ? readdirSync(dir)
        .filter((name) => /\.(xlsx|xlsm|xltx|xltm|xls)$/i.test(name))
        .sort()
    : [];
  describe(`corpus: ${source}`, () => {
    if (files.length === 0) {
      it.todo(`no ${source} files yet (see tests/corpus/REQUESTED.md)`);
      return;
    }
    for (const name of files) {
      it(name, async () => {
        expect(licenses, `${source}/${name} must be listed in tests/corpus/LICENSES.md`).toContain(`${source}/${name}`);
        let result: unknown;
        try {
          result = inspect(readFileSync(new URL(name, dir)));
        } catch (error) {
          if (!(error instanceof XlvError)) throw error;
          result = { error: error.toJSON() };
        }
        await expect(`${JSON.stringify(result, null, 2)}\n`).toMatchFileSnapshot(`./${source}/${name}.inspect.json`);
      });
    }
  });
}
