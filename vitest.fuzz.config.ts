import { defineConfig } from 'vitest/config';

// `npm run fuzz`: the property-based tests only, with many runs per property.
export default defineConfig({
  test: {
    include: ['tests/fuzz/**/*.test.ts'],
    testTimeout: 600_000,
    env: {
      XLV_FUZZ_RUNS: '20000',
    },
  },
});
