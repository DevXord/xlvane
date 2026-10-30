import { defineConfig } from 'tsdown';

export default defineConfig({
  entry: ['src/index.ts', 'src/node.ts'],
  format: 'esm',
  platform: 'neutral',
  target: 'es2022',
  dts: true,
  clean: true,
  // src/node.ts imports Node.js built-ins; they stay imports (platform 'neutral' does not know them).
  deps: { neverBundle: [/^node:/] },
});
