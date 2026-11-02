// Stages the packages of a release on npm (docs/decisions.md D-I07). Usage, in the release
// workflow after scripts/check-release-version.mjs:
//   node scripts/stage-publish.mjs vX.Y.Z
// For each package: a version already on npm is skipped, so running the workflow again after a
// failure finishes only what is missing; a package that does not exist on npm yet gets a warning,
// because npm sets up a trusted publisher only for an existing package (F-NPM-05, D-I03); a failed
// package does not stop the others, and the script fails at the end.
// XLV_STAGE_NPM (tests only): a script run with Node.js instead of npm.
import { spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('..', import.meta.url));
const version = (process.argv[2] ?? '').replace(/^v/, '');
const PACKAGES = ['packages/xlvane', 'packages/compat', 'packages/mcp'];
const npm = (args, inherit = false) =>
  process.env.XLV_STAGE_NPM
    ? spawnSync(process.execPath, [process.env.XLV_STAGE_NPM, ...args], { cwd: root, encoding: 'utf8', stdio: inherit ? 'inherit' : 'pipe' })
    : spawnSync('npm', args, { cwd: root, encoding: 'utf8', stdio: inherit ? 'inherit' : 'pipe' });

const failed = [];
for (const dir of PACKAGES) {
  const { name } = JSON.parse(readFileSync(`${root}/${dir}/package.json`, 'utf8'));
  const published = npm(['view', `${name}@${version}`, 'version']);
  if (published.status === 0 && published.stdout.trim() !== '') {
    console.log(`${name} ${version} is already on npm; not staged again.`);
    continue;
  }
  if (npm(['view', name, 'name']).status !== 0) {
    console.log(`::warning::${name} is not on npm yet: publish its first version from your machine (docs/RELEASING.md).`);
    continue;
  }
  if (npm(['stage', 'publish', '-w', dir], true).status !== 0) {
    console.log(`::error::${name} ${version} could not be staged; the other packages were still tried.`);
    failed.push(name);
  }
}
if (failed.length > 0) {
  console.error(`Not staged: ${failed.join(', ')}. Fix the cause and run the workflow again; packages already on npm are skipped.`);
  process.exit(1);
}
