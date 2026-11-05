import { defineConfig } from 'tsdown';

export default defineConfig({
  entry: ['src/index.ts', 'src/no-node.ts', 'src/node-fs.ts'],
  // Resolved by the runtime through package.json "imports": xlvane/node and dist/node-fs.js
  // (node:fs/promises) under the node condition (Node.js, Bun, Deno), dist/no-node.js elsewhere (browsers, Workers).
  deps: { neverBundle: ['#xlvane-node', '#node-fs', 'node:fs/promises'] },
  format: 'esm',
  platform: 'neutral',
  target: 'es2022',
  dts: true,
  clean: true,
});
