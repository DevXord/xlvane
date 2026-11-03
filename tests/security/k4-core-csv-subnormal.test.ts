// fromCSV after the K3 review (C11): the README says every number toCSV writes reads back as the
// same number, except whole numbers with more than 15 digits, but subnormal numbers (below
// 2.2250738585072014e-308, such as 5e-324 and 1e-310) stayed text. Cells can hold them (the writer
// and reader keep them, tests/unit/roundtrip.test.ts), and toCSV writes JavaScript's shortest form,
// so a subnormal is now a number when its text is exactly that form: no digit is lost. Other
// subnormal texts, whose digits a double cannot keep, stay text as before (D-R02).
import { describe, expect, it } from 'vitest';
import { fromCSV, toCSV } from '../../packages/xlvane/src/csv';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';

// The first line is a header, whose fields are always text.
const values = (csv: string): unknown[] => [...readWorkbook(fromCSV(csv)).sheets[0]!.cells()].slice(1).map((c) => c.value);

describe('C11: subnormal numbers in CSV', () => {
  it('reads back the subnormal numbers toCSV writes', () => {
    const numbers = [5e-324, -5e-324, 1e-310, 2.225073858507201e-308, 1.5e-323, 4.9406564584124654e-320];
    const wb = createWorkbook();
    const s = wb.addSheet('N').set('A1', 'n');
    numbers.forEach((n, i) => s.set(`A${String(i + 2)}`, n));
    const csv = toCSV(writeWorkbook(wb));
    expect(values(csv)).toEqual(numbers);
  });

  it('reads subnormals written in their shortest form as numbers', () => {
    expect(values('n\r\n5e-324\r\n1e-310\r\n-1e-310\r\n')).toEqual([5e-324, 1e-310, -1e-310]);
  });

  it('keeps subnormal texts whose digits a double cannot hold as text', () => {
    expect(values('n\r\n1.2345678901234e-315\r\n1.6e-323\r\n5.0e-324\r\n1e-400\r\n')).toEqual(['1.2345678901234e-315', '1.6e-323', '5.0e-324', '1e-400']);
  });
});
