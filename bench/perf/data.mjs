// Deterministic benchmark data: value(row, col) is computed from a hash of the position and a
// fixed seed, so every variant (and the validator) sees the same values without keeping a copy.
// Rows and columns are 0-based here; the sheet holds them from A1, without a header row.

const SEED = 0x5eed2026;
const VOCABULARY = ['North', 'South', 'East', 'West', 'Central', 'pending', 'shipped', 'returned', 'Alice', 'Bob', 'zażółć', 'Ünïcödé'];
const DAY_MS = 86_400_000;
const BASE_DATE = Date.UTC(2020, 0, 1);

/** mulberry32 step on a 32-bit state; returns a float in [0, 1). */
function hash01(n) {
  let t = (n + SEED) >>> 0;
  t = Math.imul(t ^ (t >>> 15), t | 1);
  t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
  return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
}

/** Column kinds: 0 unique id, then a repeating mix of integer, float, text, date and boolean. */
export function kindOf(col) {
  if (col === 0) return 'id';
  return ['int', 'float', 'text', 'date', 'bool'][(col - 1) % 5];
}

export function value(row, col, cols) {
  const h = hash01(row * cols + col);
  switch (kindOf(col)) {
    case 'id':
      return `ID-${String(row + 1).padStart(7, '0')}`;
    case 'int':
      return Math.floor(h * 1_000_000);
    case 'float':
      return Math.round(h * 1_000_000) / 100;
    case 'text':
      return VOCABULARY[Math.floor(h * VOCABULARY.length)];
    case 'date':
      return new Date(BASE_DATE + Math.floor(h * 2000) * DAY_MS);
    default:
      return h < 0.5;
  }
}

export function rowValues(row, cols) {
  const out = new Array(cols);
  for (let c = 0; c < cols; c++) out[c] = value(row, c, cols);
  return out;
}

/** Positions checked by every variant: corners, the middle, and one cell of each kind. */
export function samplePositions(rows, cols) {
  const r = Math.floor(rows / 2);
  const out = [
    [0, 0],
    [rows - 1, cols - 1],
    [r, 0],
  ];
  for (let c = 1; c <= Math.min(5, cols - 1); c++) out.push([r, c]);
  return out;
}

/** Compares a value read back with the expected one (dates by time, floats exactly). */
export function same(got, expected) {
  if (expected instanceof Date) return got instanceof Date && got.getTime() === expected.getTime();
  return got === expected;
}

export function describe(v) {
  return v instanceof Date ? `Date(${v.toISOString()})` : JSON.stringify(v);
}
