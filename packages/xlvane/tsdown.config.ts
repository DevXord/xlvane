import { defineConfig } from 'tsdown';

export default defineConfig({
  entry: ['src/index.ts', 'src/node.ts'],
  format: 'esm',
  platform: 'neutral',
  target: 'es2022',
  dts: true,
  clean: true,
});
