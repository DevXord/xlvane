// Checks that a release tag (vX.Y.Z) matches the version of every published package and that
// @xlvane/exceljs depends on that version of xlvane (docs/decisions.md D-1002). Usage:
//   node scripts/check-release-version.mjs v1.0.0
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('..', import.meta.url));
const tag = process.argv[2] ?? '';
const m = /^v(\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?)$/.exec(tag);
const read = (dir) => JSON.parse(readFileSync(`${root}/${dir}/package.json`, 'utf8'));
const xlvane = read('packages/xlvane');
const exceljs = read('packages/exceljs');
const problems = [];
if (!m) problems.push(`"${tag}" is not a release tag of the form vX.Y.Z.`);
else {
  for (const p of [xlvane, exceljs]) if (p.version !== m[1]) problems.push(`${p.name} has version ${p.version}, the tag is ${tag}.`);
}
if (exceljs.dependencies?.xlvane !== `^${xlvane.version}`) problems.push(`@xlvane/exceljs depends on xlvane "${exceljs.dependencies?.xlvane}", expected "^${xlvane.version}".`);
for (const p of [xlvane, exceljs]) {
  if (p.private) problems.push(`${p.name} is marked private.`);
  if (p.publishConfig?.provenance !== true || p.publishConfig?.access !== 'public') problems.push(`${p.name} needs publishConfig { access: "public", provenance: true }.`);
}
if (problems.length > 0) {
  console.error(problems.join('\n'));
  process.exit(1);
}
console.log(`Release ${tag}: xlvane and @xlvane/exceljs ${xlvane.version}.`);
