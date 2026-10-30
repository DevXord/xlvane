import { defineConfig } from 'tsdown';

export default defineConfig({
  // cli: the `xlvane` command (package.json "bin"), not part of the library API.
  entry: { index: 'src/index.ts', node: 'src/node.ts', cli: 'src/cli/bin.ts' },
  format: 'esm',
  platform: 'neutral',
  target: 'es2022',
  dts: true,
  clean: true,
  // src/node.ts imports Node.js built-ins; they stay imports (platform 'neutral' does not know them).
  deps: { neverBundle: [/^node:/] },
});
