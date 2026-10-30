// `npm run check:pkg`: builds every package and checks it the way a consumer sees it.
//
// 1. publint --strict (warnings count as errors)
// 2. attw --pack <dir> --profile esm-only (types resolve for ESM consumers; CJS is not supported)
// 3. Every file in dist loads with require(): Node.js throws ERR_REQUIRE_ASYNC_MODULE for a
//    module graph with top-level await (docs/decisions.md, F-NODE-01), which CLAUDE.md forbids.
// 4. Size budget of xlvane: gzip size of all dist/*.js files (docs/decisions.md D-208). The
//    budget is stated in packages/xlvane/README.md; tests/policy/size-budget.test.ts keeps the
//    two in sync.
import { spawnSync } from 'node:child_process';
import { readdirSync, readFileSync } from 'node:fs';
import { gzipSync } from 'node:zlib';
import { createRequire } from 'node:module';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('..', import.meta.url));
const packages = ['packages/xlvane', 'packages/exceljs', 'packages/mcp'];
const failures = [];
/** Size budget of packages/xlvane/dist/*.js, gzip level 9, in bytes. Same number as in the README. */
const SIZE_BUDGET_BYTES = 100_000;

function run(label, command) {
  console.log(`\n> ${label}`);
  // shell: true so that the npm/npx .cmd shims also run on Windows.
  const result = spawnSync(command, { cwd: root, stdio: 'inherit', shell: true });
  if (result.status !== 0) failures.push(label);
}

run('build', 'npm run build');
if (failures.length === 0) {
  for (const dir of packages) {
    run(`publint ${dir}`, `npx publint --strict ${dir}`);
    run(`attw ${dir}`, `npx attw --pack ${dir} --profile esm-only`);
  }
  const require = createRequire(import.meta.url);
  for (const dir of packages) {
    const dist = join(root, dir, 'dist');
    for (const file of readdirSync(dist).filter((f) => f.endsWith('.js'))) {
      const label = `require(esm) ${dir}/dist/${file}`;
      try {
        require(join(dist, file));
        console.log(`\n> ${label}: ok`);
      } catch (error) {
        console.error(`\n> ${label}: ${error.code ?? ''} ${error.message}`);
        failures.push(label);
      }
    }
  }
}

if (failures.length === 0) {
  const dist = join(root, 'packages/xlvane/dist');
  let gzipped = 0;
  for (const file of readdirSync(dist).filter((f) => f.endsWith('.js'))) {
    gzipped += gzipSync(readFileSync(join(dist, file)), { level: 9 }).length;
  }
  const label = `size budget xlvane: ${gzipped} of ${SIZE_BUDGET_BYTES} bytes (gzip)`;
  console.log(`\n> ${label}`);
  if (gzipped > SIZE_BUDGET_BYTES) failures.push(label);
}

if (failures.length > 0) {
  console.error(`\ncheck:pkg failed:\n  ${failures.join('\n  ')}`);
  process.exit(1);
}
console.log('\ncheck:pkg passed');
