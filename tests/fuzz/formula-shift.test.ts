// Property tests for the shared-formula tokenizer: hostile text never throws and stays fast,
// and shifting a well-formed formula there and back again restores it.
import fc from 'fast-check';
import { describe, expect, it } from 'vitest';
import { columnLetters } from '../../packages/xlvane/src/sml/cell-ref';
import { compileFormula, shiftFormula } from '../../packages/xlvane/src/sml/formula-shift';
import { NUM_RUNS } from './runs';

const RUNS = { numRuns: NUM_RUNS };

const formulaChars = fc.constantFrom(
  'A', 'B', 'x', 'Z', 'R', 'C', '1', '0', '9', '$', ':', '!', "'", '"', '[', ']', '#', '(', ')', ',', ';',
  ' ', '+', '-', '*', '/', '^', '&', '=', '<', '>', '%', '@', '{', '}', '.', 'E', '_', '\\', '?', 'ż', '~',
);
const formulaLike = fc.string({ unit: formulaChars, maxLength: 60 });

const cellRef = fc
  .record({
    col: fc.integer({ min: 1, max: 16_384 }),
    row: fc.integer({ min: 1, max: 1_048_576 }),
    colAbs: fc.boolean(),
    rowAbs: fc.boolean(),
  })
  .map(({ col, row, colAbs, rowAbs }) => `${colAbs ? '$' : ''}${columnLetters(col)}${rowAbs ? '$' : ''}${String(row)}`);
const atom = fc.oneof(
  cellRef,
  fc.tuple(cellRef, cellRef).map(([a, b]) => `${a}:${b}`),
  fc.tuple(cellRef, fc.constantFrom("'My Sheet'!", 'Data!', '[1]Ext!')).map(([r, s]) => s + r),
  fc.constantFrom('TAX_RATE', '"A1 text"', '#N/A', '1.5E-3', 'TRUE', 'SUM(1:3)', 'COUNT(A:C)', 'Table1[Col]'),
  fc.integer({ min: 0, max: 1000 }).map(String),
);
const formula = fc.array(atom, { minLength: 1, maxLength: 6 }).chain((atoms) =>
  fc.array(fc.constantFrom('+', '-', '*', '&', ','), { minLength: atoms.length - 1, maxLength: atoms.length - 1 }).map((ops) =>
    atoms.reduce((acc, a, i) => (i === 0 ? a : `${acc}${ops[i - 1] ?? '+'}${a}`), ''),
  ),
);

describe('formula shifting properties', () => {
  it('never throws on arbitrary text', () => {
    fc.assert(
      fc.property(fc.oneof(formulaLike, fc.string({ unit: 'binary', maxLength: 60 })), fc.integer({ min: -5, max: 5 }), (text, d) => {
        const out = shiftFormula(text, d, -d);
        expect(out === null || typeof out === 'string').toBe(true);
      }),
      RUNS,
    );
  });

  it('round-trips: shifting by (r, c) and back restores the formula', () => {
    fc.assert(
      fc.property(formula, fc.integer({ min: -3, max: 3 }), fc.integer({ min: -3, max: 3 }), (text, dr, dc) => {
        expect(shiftFormula(text, 0, 0)).toBe(text);
        const compiled = compileFormula(text);
        expect(compiled).not.toBeNull();
        const there = shiftFormula(text, dr, dc);
        if (there === null) {
          expect(compiled?.canShift(dr, dc)).toBe(false);
          return;
        }
        expect(shiftFormula(there, -dr, -dc)).toBe(text);
      }),
      RUNS,
    );
  });

  it('runs in linear time on long hostile input', () => {
    const inputs = ['['.repeat(200_000), '"'.repeat(200_001), 'A'.repeat(200_000), '$1:'.repeat(70_000), 'A1:'.repeat(70_000) + 'A1', "'".repeat(200_001)];
    const start = Date.now();
    for (const input of inputs) shiftFormula(input, 1, 1);
    expect(Date.now() - start).toBeLessThan(2000);
  });
});
