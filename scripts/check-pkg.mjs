// `npm run check:pkg`: builds every package and checks it the way a consumer sees it.
//
// 1. publint --strict (warnings count as errors)
// 2. attw --pack <dir> --profile esm-only (types resolve for ESM consumers; CJS is not supported)
// 3. Every file in dist loads with require(): Node.js throws ERR_REQUIRE_ASYNC_MODULE for a
//    module graph with top-level await (docs/decisions.md, F-NODE-01), which CLAUDE.md forbids.
//    The CLI entry (xlvane/dist/cli.js) runs when loaded, so it is checked by running
//    `--version` instead; it imports the same chunks, which are loaded above.
// 4. Size budgets of xlvane (gzip): the library (every dist/*.js except cli.js, D-208, D-408)
//    and the CLI (cli.js, D-608). Both are stated in packages/xlvane/README.md;
//    tests/policy/size-budget.test.ts keeps the numbers in sync.
// 5. The xlvane tarball contains docs/, the Agent Skill and llms-full.txt (D-707).
import { spawnSync } from 'node:child_process';
import { readdirSync, readFileSync } from 'node:fs';
import { gzipSync } from 'node:zlib';
import { createRequire } from 'node:module';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('..', import.meta.url));
const packages = ['packages/xlvane', 'packages/exceljs'];
const failures = [];
/** Size budget of the library files in packages/xlvane/dist, gzip level 9, in bytes. Same number as in the README. */
const SIZE_BUDGET_BYTES = 175_000;
/** Size budget of packages/xlvane/dist/cli.js, gzip level 9, in bytes. Same number as in the README. */
const CLI_SIZE_BUDGET_BYTES = 40_000;
const CLI_FILE = 'cli.js';

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
      if (file === CLI_FILE) {
        // The command entry point runs when loaded, so it is run with --version instead.
        const flag = '--version';
        const label = `${dir}/dist/${CLI_FILE} ${flag}`;
        const r = spawnSync(process.execPath, [join(dist, CLI_FILE), flag], { encoding: 'utf8' });
        console.log(`\n> ${label}: ${r.status === 0 ? `ok (${r.stdout.trim().split('\n')[0]})` : `exit ${r.status} ${r.stderr}`}`);
        if (r.status !== 0) failures.push(label);
        continue;
      }
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
  // 5. The xlvane tarball carries the docs, the Agent Skill and llms-full.txt (D-707).
  const pack = spawnSync('npm', ['pack', '--dry-run', '--json'], { cwd: join(root, 'packages/xlvane'), encoding: 'utf8', shell: true });
  const packed = pack.status === 0 ? JSON.parse(pack.stdout)[0].files.map((f) => f.path) : [];
  for (const required of ['llms-full.txt', 'skills/xlvane/SKILL.md', 'docs/errors.md', 'docs/cli-json.md', 'docs/cli-json/common.schema.json', 'dist/cli.js']) {
    const label = `xlvane tarball contains ${required}`;
    console.log(`\n> ${label}: ${packed.includes(required) ? 'ok' : 'missing'}`);
    if (!packed.includes(required)) failures.push(label);
  }
}

if (failures.length === 0) {
  const dist = join(root, 'packages/xlvane/dist');
  let gzipped = 0;
  for (const file of readdirSync(dist).filter((f) => f.endsWith('.js') && f !== CLI_FILE)) {
    gzipped += gzipSync(readFileSync(join(dist, file)), { level: 9 }).length;
  }
  const label = `size budget xlvane: ${gzipped} of ${SIZE_BUDGET_BYTES} bytes (gzip)`;
  console.log(`\n> ${label}`);
  if (gzipped > SIZE_BUDGET_BYTES) failures.push(label);
  const cli = gzipSync(readFileSync(join(dist, CLI_FILE)), { level: 9 }).length;
  const cliLabel = `size budget xlvane CLI: ${cli} of ${CLI_SIZE_BUDGET_BYTES} bytes (gzip)`;
  console.log(`\n> ${cliLabel}`);
  if (cli > CLI_SIZE_BUDGET_BYTES) failures.push(cliLabel);
}

if (failures.length > 0) {
  console.error(`\ncheck:pkg failed:\n  ${failures.join('\n  ')}`);
  process.exit(1);
}
console.log('\ncheck:pkg passed');
