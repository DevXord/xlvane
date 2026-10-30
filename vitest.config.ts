import { fileURLToPath } from 'node:url';
import { defineConfig } from 'vitest/config';

// `npm test`: unit, security, policy and corpus tests. Property tests in tests/fuzz run here
// too, with a small number of runs; `npm run fuzz` runs them with many more.
export default defineConfig({
  resolve: {
    // In-process tests load xlvane from its sources, not from dist: tests/exceljs/alias.test.ts
    // rebuilds dist while other test files run.
    alias: [
      { find: /^xlvane\/node$/, replacement: fileURLToPath(new URL('packages/xlvane/src/node.ts', import.meta.url)) },
      { find: /^xlvane$/, replacement: fileURLToPath(new URL('packages/xlvane/src/index.ts', import.meta.url)) },
    ],
  },
  test: {
    include: ['tests/**/*.test.ts'],
    // tests/stream/memory.test.ts calls gc() to measure retained memory (Vitest option
    // `execArgv`: arguments for the node processes that run test files).
    execArgv: ['--expose-gc'],
    env: {
      XLV_FUZZ_RUNS: process.env['XLV_FUZZ_RUNS'] ?? '100',
    },
  },
});
