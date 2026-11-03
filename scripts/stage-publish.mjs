// Stages the packages of a release on npm (docs/decisions.md D-I07). Usage, in the release
// workflow after scripts/check-release-version.mjs:
//   node scripts/stage-publish.mjs vX.Y.Z
// For each package: a version already on npm is skipped, so running the workflow again after a
// failure finishes only what is missing; a package that does not exist on npm yet gets a warning,
// because npm sets up a trusted publisher only for an existing package (F-NPM-05, D-I03); a failed
// package does not stop the others, and the script fails at the end. Only E404 from `npm view`
// counts as "not on npm"; any other error fails the package (D-J01).
// TLOG_CREATE_ENTRY_ERROR is tried again, up to three tries 30 s apart: npm's sigstore client waits
// 5 s for the transparency log and retries, and when the first request was only slow, the retry
// gets 409 "an equivalent entry already exists" and npm stops before anything is staged (npm
// 11.21.1, sigstore/dist/config.js L25-26 and L105, @sigstore/sign/dist/witness/tlog/client.js
// L34-56; seen twice for @xlvane/compat 1.3.0, D-K05). A new try signs again, so it gets a new entry.
// XLV_STAGE_RETRY_MS (tests only): the wait between tries.
// XLV_STAGE_NPM (tests only): a script run with Node.js instead of npm.
import { spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('..', import.meta.url));
const version = (process.argv[2] ?? '').replace(/^v/, '');
const PACKAGES = ['packages/xlvane', 'packages/compat', 'packages/mcp'];
const npm = (args) =>
  process.env.XLV_STAGE_NPM
    ? spawnSync(process.execPath, [process.env.XLV_STAGE_NPM, ...args], { cwd: root, encoding: 'utf8' })
    : spawnSync('npm', args, { cwd: root, encoding: 'utf8' });

const TRIES = 3;
const RETRY_MS = Number(process.env.XLV_STAGE_RETRY_MS ?? 30_000);
const failed = [];
const notFound = (r) => /\bE404\b/.test(`${r.stdout ?? ''}${r.stderr ?? ''}`);
const registryError = (name, r) => {
  const code = /\bE\d{3}\b|\bE[A-Z]+\b/.exec(`${r.stdout ?? ''}${r.stderr ?? ''}`)?.[0] ?? `exit ${String(r.status)}`;
  console.log(`::error::${name}: the npm registry could not be asked (${code}); the other packages were still tried.`);
  failed.push(name);
};
for (const dir of PACKAGES) {
  const { name } = JSON.parse(readFileSync(`${root}/${dir}/package.json`, 'utf8'));
  const published = npm(['view', `${name}@${version}`, 'version']);
  if (published.status === 0 && published.stdout.trim() !== '') {
    console.log(`${name} ${version} is already on npm; not staged again.`);
    continue;
  }
  // Only E404 means "not there": any other failure (an outage, a rate limit) must not be taken
  // for a missing package, or the job would pass with nothing staged (D-J01).
  if (published.status !== 0 && !notFound(published)) {
    registryError(name, published);
    continue;
  }
  const exists = npm(['view', name, 'name']);
  if (exists.status !== 0) {
    if (notFound(exists)) console.log(`::warning::${name} is not on npm yet: publish its first version from your machine (docs/RELEASING.md).`);
    else registryError(name, exists);
    continue;
  }
  let staged = false;
  for (let attempt = 1; attempt <= TRIES; attempt++) {
    const r = npm(['stage', 'publish', '-w', dir]);
    process.stdout.write(r.stdout ?? '');
    process.stderr.write(r.stderr ?? '');
    if (r.status === 0) {
      staged = true;
      break;
    }
    if (attempt === TRIES || !/\bTLOG_CREATE_ENTRY_ERROR\b/.test(`${r.stdout ?? ''}${r.stderr ?? ''}`)) break;
    console.log(`::warning::${name}: the transparency log did not answer in time (TLOG_CREATE_ENTRY_ERROR); trying again in ${String(RETRY_MS / 1000)} s (try ${String(attempt + 1)} of ${String(TRIES)}).`);
    Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, RETRY_MS);
  }
  if (!staged) {
    console.log(`::error::${name} ${version} could not be staged; the other packages were still tried.`);
    failed.push(name);
  }
}
if (failed.length > 0) {
  console.error(`Not staged: ${failed.join(', ')}. Fix the cause and run the workflow again; packages already on npm are skipped.`);
  process.exit(1);
}
