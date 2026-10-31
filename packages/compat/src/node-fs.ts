// What `#node-fs` resolves to under the `node` condition (package.json "imports"). Node.js does
// not accept a `node:` URL as an "imports" target, so the builtin is re-exported from here (D-R07).
// @ts-expect-error -- this package has no Node.js types; the module is loaded only in Node.js, Bun and Deno.
export { readFile, writeFile } from 'node:fs/promises';
