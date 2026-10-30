// Property tests for adjustFormula (docs/facts/ref-adjust.md): hostile text never throws, every
// adjusted formula can be read again, and inserting rows or columns and then deleting the same
// band restores the formula.
import fc from 'fast-check';
import { describe, expect, it } from 'vitest';
import { columnLetters } from '../../packages/xlvane/src/sml/cell-ref';
import { compileFormula } from '../../packages/xlvane/src/sml/formula-shift';
import { adjustFormula, type ReferenceChange } from '../../packages/xlvane/src/sml/ref-adjust';
import { NUM_RUNS } from './runs';

const RUNS = { numRuns: NUM_RUNS };

const formulaChars = fc.constantFrom(
  'A', 'B', 'x', 'Z', 'R', 'C', '1', '0', '9', '$', ':', '!', "'", '"', '[', ']', '#', '(', ')', ',', ';',
  ' ', '+', '-', '*', '/', '^', '&', '=', '<', '>', '%', '@', '{', '}', '.', 'E', '_', '\\', '?', 'ż', '~',
);
const formulaLike = fc.string({ unit: formulaChars, maxLength: 60 });

// References kept away from the end of the grid, where an inserted band can push a range end
// against the last row or column (R-09, R-10) and the round trip legitimately differs.
const cellRef = fc
  .record({ col: fc.integer({ min: 1, max: 300 }), row: fc.integer({ min: 1, max: 300 }), colAbs: fc.boolean(), rowAbs: fc.boolean() })
  .map(({ col, row, colAbs, rowAbs }) => `${colAbs ? '$' : ''}${columnLetters(col)}${rowAbs ? '$' : ''}${String(row)}`);
const ordered = fc
  .tuple(fc.integer({ min: 1, max: 300 }), fc.integer({ min: 1, max: 300 }), fc.integer({ min: 1, max: 300 }), fc.integer({ min: 1, max: 300 }))
  .map(([c1, c2, r1, r2]) => `${columnLetters(Math.min(c1, c2))}${String(Math.min(r1, r2))}:$${columnLetters(Math.max(c1, c2))}${String(Math.max(r1, r2))}`);
const atom = fc.oneof(
  cellRef,
  ordered,
  fc.tuple(cellRef, fc.constantFrom("'My Sheet'!", 'Data!', 'Other!', '[1]Data!')).map(([r, s]) => s + r),
  fc.tuple(ordered, fc.constantFrom('Data!', "'Other Sheet'!")).map(([r, s]) => s + r),
  fc.constantFrom('TAX_RATE', '"A1 text"', '#N/A', '1.5E-3', 'TRUE', 'SUM(1:3)', 'COUNT(A:C)', 'SUM($4:$9)', 'Table1[Col]'),
  fc.integer({ min: 0, max: 1000 }).map(String),
);
const formula = fc.array(atom, { minLength: 1, maxLength: 6 }).chain((atoms) =>
  fc.array(fc.constantFrom('+', '-', '*', '&', ','), { minLength: atoms.length - 1, maxLength: atoms.length - 1 }).map((ops) =>
    atoms.reduce((acc, a, i) => (i === 0 ? a : `${acc}${ops[i - 1] ?? '+'}${a}`), ''),
  ),
);
const change = fc.record({
  sheet: fc.constantFrom('Data', 'data', 'Other', 'My Sheet'),
  formulaSheet: fc.constantFrom('Data', 'Other'),
  axis: fc.constantFrom('row' as const, 'col' as const),
  at: fc.integer({ min: 1, max: 320 }),
  count: fc.integer({ min: 1, max: 5 }),
});

describe('reference adjustment properties', () => {
  it('never throws on arbitrary text, and its output can be read again', () => {
    fc.assert(
      fc.property(fc.oneof(formulaLike, formula, fc.string({ unit: 'binary', maxLength: 60 })), change, fc.boolean(), (text, c, remove) => {
        const out = adjustFormula(text, { ...c, count: remove ? -c.count : c.count });
        if (out !== null) expect(compileFormula(out)).not.toBeNull();
      }),
      RUNS,
    );
  });

  it('inserting a band and deleting it again restores the formula', () => {
    fc.assert(
      fc.property(formula, change, (text, c) => {
        const inserted = adjustFormula(text, c);
        expect(inserted).not.toBeNull();
        const back: ReferenceChange = { ...c, count: -c.count };
        expect(adjustFormula(inserted ?? '', back)).toBe(text);
      }),
      RUNS,
    );
  });
});
