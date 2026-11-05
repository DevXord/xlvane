// Checks that a release tag (vX.Y.Z) matches the version of every published package and that
// @xlvane/exceljs and @xlvane/mcp depend on that version of xlvane (docs/decisions.md D-1002, D-I03). Usage:
//   node scripts/check-release-version.mjs v1.0.0
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('..', import.meta.url));
const tag = process.argv[2] ?? '';
// No prerelease suffix: npm refuses to publish a prerelease without --tag, which the release
// workflow does not pass (npm 11.21.1 lib/commands/publish.js; D-K03).
const m = /^v(\d+\.\d+\.\d+)$/.exec(tag);
const read = (dir) => JSON.parse(readFileSync(`${root}/${dir}/package.json`, 'utf8'));
const xlvane = read('packages/xlvane');
const exceljs = read('packages/exceljs');
const mcp = read('packages/mcp');
const serverJson = JSON.parse(readFileSync(`${root}/packages/mcp/server.json`, 'utf8'));
const problems = [];
if (!m) problems.push(`"${tag}" is not a release tag of the form vX.Y.Z.`);
else {
  for (const p of [xlvane, exceljs, mcp]) if (p.version !== m[1]) problems.push(`${p.name} has version ${p.version}, the tag is ${tag}.`);
}
// packages/mcp/server.json is published to the MCP Registry with the same version (D-I06).
if (serverJson.version !== mcp.version || serverJson.packages?.[0]?.version !== mcp.version) problems.push(`packages/mcp/server.json has version ${serverJson.version}, @xlvane/mcp has ${mcp.version}.`);
if (serverJson.name !== mcp.mcpName) problems.push(`packages/mcp/server.json is named ${serverJson.name}, but @xlvane/mcp has mcpName ${mcp.mcpName}.`);
for (const p of [exceljs, mcp]) if (p.dependencies?.xlvane !== `^${xlvane.version}`) problems.push(`${p.name} depends on xlvane "${p.dependencies?.xlvane}", expected "^${xlvane.version}".`);
for (const p of [xlvane, exceljs, mcp]) {
  if (p.private) problems.push(`${p.name} is marked private.`);
  if (p.publishConfig?.provenance !== true || p.publishConfig?.access !== 'public') problems.push(`${p.name} needs publishConfig { access: "public", provenance: true }.`);
}
if (problems.length > 0) {
  console.error(problems.join('\n'));
  process.exit(1);
}
console.log(`Release ${tag}: xlvane, @xlvane/exceljs and @xlvane/mcp ${xlvane.version}.`);
