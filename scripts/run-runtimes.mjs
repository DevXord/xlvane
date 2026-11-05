// `npm run test:runtimes`: builds the portable smoke bundle once and runs it in Bun, Deno,
// Chromium (Playwright) and workerd (Cloudflare Workers). Node.js runs it in `npm test`
// (tests/runtimes/node.test.ts). A runtime whose program is missing counts as a failure unless
// it is listed in XLV_SKIP_RUNTIMES (comma-separated: bun,deno,browser,workerd). D-904.
import { spawnSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildSmoke } from './build-smoke.mjs';
import { runInBrowser } from '../tests/runtimes/browser.mjs';
import { runInWorkerd } from '../tests/runtimes/workerd.mjs';

const root = fileURLToPath(new URL('..', import.meta.url));
const runner = join(root, 'tests/runtimes/run-bundle.mjs');
const skip = new Set((process.env.XLV_SKIP_RUNTIMES ?? '').split(',').map((s) => s.trim()).filter(Boolean));

function bunPath() {
  const candidates = [process.env.BUN_INSTALL ? join(process.env.BUN_INSTALL, 'bin/bun') : '', join(process.env.HOME ?? '', '.bun/bin/bun')];
  for (const c of candidates) if (c && existsSync(c)) return c;
  return 'bun';
}

/** Runs a command that prints the results as JSON; returns them or throws. */
function viaProcess(command, args) {
  const r = spawnSync(command, args, { cwd: root, encoding: 'utf8', timeout: 300_000 });
  if (r.error) throw new Error(`${command}: ${r.error.message}`);
  const line = r.stdout.trim().split('\n').pop() ?? '';
  try {
    return JSON.parse(line);
  } catch {
    throw new Error(`${command} exited with ${r.status}:\n${r.stdout}\n${r.stderr}`);
  }
}

const RUNTIMES = {
  bun: (bundle) => viaProcess(bunPath(), ['run', '--no-install', runner, bundle]),
  // Deno: no config, lock file, npm or remote modules; read access to the bundle directory only.
  deno: (bundle) =>
    viaProcess(join(root, 'node_modules/.bin/deno'), ['run', '--no-config', '--no-lock', '--no-npm', '--no-remote', '--no-prompt', `--allow-read=${dirname(bundle)}`, runner, bundle]),
  browser: (bundle) => runInBrowser(bundle),
  workerd: (bundle) => runInWorkerd(bundle),
};

async function main() {
  // The source bundle runs everywhere; the bundle made from the built packages (as a user's
  // bundler would make it, D-R01) runs in the two runtimes without Node.js modules.
  const build = spawnSync('npm', ['run', 'build', '-w', 'packages/xlvane', '-w', 'packages/exceljs'], { cwd: root, encoding: 'utf8', shell: process.platform === 'win32' });
  if (build.status !== 0) throw new Error(`npm run build failed:\n${build.stdout}\n${build.stderr}`);
  const bundle = await buildSmoke();
  const packaged = await buildSmoke(undefined, { packages: true });
  const runs = [...Object.entries(RUNTIMES).map(([name, run]) => [name, run, bundle]), ['browser (packages)', RUNTIMES.browser, packaged, 'browser'], ['workerd (packages)', RUNTIMES.workerd, packaged, 'workerd']];
  const rows = [];
  let failed = false;
  for (const [name, run, file, key = name] of runs) {
    if (skip.has(key)) {
      rows.push([name, 'skipped (XLV_SKIP_RUNTIMES)', '']);
      continue;
    }
    try {
      const results = await run(file);
      const bad = results.filter((r) => !r.ok);
      if (bad.length > 0 || results.length === 0) failed = true;
      rows.push([name, `${results.length - bad.length}/${results.length} passed`, bad.map((r) => `${r.name}: ${r.error ?? ''}`).join('\n    ')]);
    } catch (error) {
      failed = true;
      rows.push([name, 'error', String(error instanceof Error ? error.message : error)]);
    }
  }
  for (const [name, status, detail] of rows) console.log(`${name.padEnd(18)} ${status}${detail ? `\n    ${detail}` : ''}`);
  if (failed) process.exitCode = 1;
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
