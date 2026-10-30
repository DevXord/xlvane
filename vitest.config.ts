import { defineConfig } from 'vitest/config';

// `npm test`: unit, security, policy and corpus tests. Property tests in tests/fuzz run here
// too, with a small number of runs; `npm run fuzz` runs them with many more.
export default defineConfig({
  test: {
    include: ['tests/**/*.test.ts'],
    env: {
      XLV_FUZZ_RUNS: '100',
    },
  },
});
