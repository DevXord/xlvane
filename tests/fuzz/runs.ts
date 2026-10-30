// Number of runs per property, from XLV_FUZZ_RUNS (100 for `npm test`, 20000 for `npm run fuzz`).

const DEFAULT_RUNS = 100;

function readRuns(): number {
  const raw = process.env['XLV_FUZZ_RUNS'];
  if (raw === undefined || raw === '') return DEFAULT_RUNS;
  const runs = Number(raw);
  if (!Number.isSafeInteger(runs) || runs < 1) {
    throw new Error(`XLV_FUZZ_RUNS must be a positive whole number, got "${raw}".`);
  }
  return runs;
}

/** Runs per property; pass as `{ numRuns: NUM_RUNS }` to `fc.assert`. */
export const NUM_RUNS: number = readRuns();
