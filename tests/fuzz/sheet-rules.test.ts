// Property tests for data validations and conditional formatting: what the writer accepts reads
// back the same, and hostile attribute values in the file never throw anything but XlvError
// (docs/decisions.md D-C01, D-C02).
import fc from 'fast-check';
import { describe, expect, it } from 'vitest';
import { isXlvError } from '../../packages/xlvane/src/errors';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { escapeXml, buildXlsx } from '../helpers/xlsx-builder';
import { NUM_RUNS } from './runs';

const RUNS = { numRuns: Math.max(1, Math.floor(NUM_RUNS / 10)) };
const value = fc.oneof(fc.string({ maxLength: 10 }), fc.constantFrom('1', '-1', '0', 'true', '1e400', 'A1', 'A1:B2 C3', '', ' ', '"', '_x0041_', '99999999999'));
const formula = fc.oneof(fc.integer({ min: -1000, max: 1000 }), fc.constantFrom('A1>0', 'SUM($A$1:$A$9)', '"x"', 'TRUE'));

describe('sheet rule properties', () => {
  it('reads hostile cfRule and dataValidation attributes without throwing anything but XlvError', () => {
    fc.assert(
      fc.property(fc.array(fc.tuple(fc.constantFrom('type', 'priority', 'dxfId', 'operator', 'rank', 'stdDev', 'sqref', 'iconSet', 'minLength', 'gte', 'val', 'errorStyle', 'showDropDown'), value), { maxLength: 6 }), (pairs) => {
        const attrs = pairs.map(([k, v]) => ` ${k}="${escapeXml(v)}"`);
        const unique = (list: string[]) => [...new Map(list.map((a) => [a.split('=')[0], a])).values()].join('');
        const after =
          `<conditionalFormatting${unique([' sqref="A1"', ...attrs])}><cfRule${unique([' type="cellIs"', ' priority="1"', ...attrs])}><formula>1</formula><colorScale><cfvo${unique(attrs)}/><cfvo type="max"/><color rgb="FF000000"/><color${unique(attrs)}/></colorScale></cfRule></conditionalFormatting>` +
          `<dataValidations><dataValidation${unique([' sqref="A1"', ...attrs])}><formula1>1</formula1></dataValidation></dataValidations>`;
        try {
          const sheet = readWorkbook(buildXlsx({ sheets: [{ name: 'S', after }] })).getSheet('S');
          expect(Array.isArray(sheet.conditionalFormats) && Array.isArray(sheet.dataValidations)).toBe(true);
        } catch (error) {
          if (!isXlvError(error)) throw error;
        }
      }),
      RUNS,
    );
  });

  it('reads back the validations and rules the writer accepted', () => {
    fc.assert(
      fc.property(
        fc.constantFrom('whole', 'decimal', 'textLength', 'date', 'time'),
        fc.constantFrom('between', 'notBetween', 'equal', 'lessThan', 'greaterThanOrEqual'),
        formula,
        formula,
        fc.string({ maxLength: 20 }).filter((s) => !/[\uD800-\uDFFF]/.test(s)),
        (type, operator, f1, f2, prompt) => {
          const wb = createWorkbook();
          const ranged = operator === 'between' || operator === 'notBetween';
          const rule = { type, operator, formula1: f1, ...(ranged ? { formula2: f2 } : {}), prompt, showInputMessage: true } as const;
          wb.addSheet('S').validate('B2:B9', rule).conditionalFormat('C1:C9', [{ type: 'cellIs', operator, formulas: ranged ? [f1, f2] : [f1], style: { font: { bold: true } } }]);
          const sheet = readWorkbook(writeWorkbook(wb)).getSheet('S');
          expect(sheet.dataValidations[0]).toMatchObject({ type, operator, formula1: String(f1), formula2: ranged ? String(f2) : null, prompt });
          expect(sheet.conditionalFormats[0]?.rules[0]).toMatchObject({ type: 'cellIs', operator, formulas: (ranged ? [f1, f2] : [f1]).map(String), style: { font: { bold: true } } });
        },
      ),
      RUNS,
    );
  });
});
