// Bundles tests/runtimes/smoke.ts into one self-contained ES module (xlvane, @xlvane/exceljs and
// fflate included, no node: imports) for the runtime runners in tests/runtimes/ (D-904).
// Usage: node scripts/build-smoke.mjs [outDir]   (default node_modules/.cache/xlvane-runtimes)
import { readFile, writeFile } from 'node:fs/promises';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { build } from 'tsdown';

const root = fileURLToPath(new URL('..', import.meta.url));

/** Builds the bundle and returns the path of smoke.js. */
export async function buildSmoke(outDir = join(root, 'node_modules/.cache/xlvane-runtimes')) {
  // Node-only code paths of @xlvane/exceljs import xlvane/node dynamically; it is replaced with a
  // stub so that the bundle stays a single file without node: imports.
  await build({
    config: false,
    entry: { smoke: join(root, 'tests/runtimes/smoke.ts') },
    outDir,
    format: 'esm',
    platform: 'neutral',
    target: 'es2022',
    dts: false,
    clean: false,
    logLevel: 'warn',
    alias: {
      xlvane: join(root, 'packages/xlvane/src/index.ts'),
      'xlvane/node': join(root, 'tests/runtimes/node-stub.ts'),
      '@xlvane/exceljs': join(root, 'packages/exceljs/src/index.ts'),
    },
    deps: { alwaysBundle: [/./], onlyBundle: false },
    outputOptions: { codeSplitting: false },
  });
  const file = join(outDir, 'smoke.js');
  const code = await readFile(file, 'utf8');
  const nodeImports = code.match(/(?:from|import)\s*\(?\s*['"](node:[^'"]+)['"]/g);
  if (nodeImports) throw new Error(`The smoke bundle imports Node.js modules: ${nodeImports.join(', ')}`);
  const bare = code.match(/^\s*import\s[^;]*from\s*['"][^./][^'"]*['"]/gm);
  if (bare) throw new Error(`The smoke bundle is not self-contained: ${bare.join(', ')}`);
  return file;
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  buildSmoke(process.argv[2] ? resolve(process.argv[2]) : undefined).then(
    (file) => console.log(file),
    (error) => {
      console.error(error);
      process.exitCode = 1;
    },
  );
}
