// Bundles tests/runtimes/smoke.ts into one self-contained ES module (xlvane, @xlvane/compat and
// fflate included, no node: imports) for the runtime runners in tests/runtimes/ (D-904, D-R01).
// Usage: node scripts/build-smoke.mjs [outDir]   (default node_modules/.cache/xlvane-runtimes)
import { readFile, writeFile } from 'node:fs/promises';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { build } from 'tsdown';

const root = fileURLToPath(new URL('..', import.meta.url));

/**
 * Builds the bundle and returns the path of smoke.js.
 *
 * - default: from the sources (the src folders of the packages), for every runtime;
 * - `{ packages: true }`: from the built packages as a user's bundler sees them (package exports,
 *   dist files, browser platform, no aliases), writing `smoke-packages.js`. It checks that
 *   @xlvane/compat bundles for browsers and Cloudflare Workers without node: imports (D-R01).
 *   Run `npm run build` first.
 */
export async function buildSmoke(outDir = join(root, 'node_modules/.cache/xlvane-runtimes'), { packages = false } = {}) {
  const name = packages ? 'smoke-packages' : 'smoke';
  await build({
    config: false,
    entry: { [name]: join(root, 'tests/runtimes/smoke.ts') },
    outDir,
    format: 'esm',
    platform: packages ? 'browser' : 'neutral',
    target: 'es2022',
    dts: false,
    clean: false,
    logLevel: 'warn',
    // No alias for xlvane/node: @xlvane/compat loads it only through a non-literal specifier
    // when a file is read or written by name, so it never enters the bundle.
    alias: packages
      ? {}
      : {
          xlvane: join(root, 'packages/xlvane/src/index.ts'),
          '@xlvane/compat': join(root, 'packages/compat/src/index.ts'),
        },
    deps: { alwaysBundle: [/./], onlyBundle: false },
    outputOptions: { codeSplitting: false },
  });
  const file = join(outDir, `${name}.js`);
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
