// Formula-injection escaping also covers text that starts with a C0 control character (NUL from
// _x0000_ included, which @xlvane/compat removes when writing, so "\0=cmd" became "=cmd") and the
// small forms of = + - @ (U+FE66, U+FE62, U+FE63, U+FE6B, compatibility decompositions of the
// ASCII signs). Whether spreadsheet applications run such fields as formulas is not checked; the
// escape is a precaution (1.1.1 security review, D-H01; docs/facts/csv.md CSV-12).
import { describe, expect, it } from 'vitest';
import ExcelJS from '../../packages/compat/src/index';
import { toCSV } from '../../packages/xlvane/src/csv';
import { buildXlsx } from '../helpers/xlsx-builder';

const starts = ['_x0000_=cmd', '_x0001_=cmd', '_x001F_+1', '\u{FE66}cmd', '\u{FE62}1', '\u{FE63}1', '\u{FE6B}SUM(A1)'];

describe('more formula-like starts are escaped', () => {
  it('toCSV escapes text starting with a C0 control or a small form sign', () => {
    const data = `<row r="1">${starts.map((s, i) => `<c r="${String.fromCharCode(65 + i)}1" t="inlineStr"><is><t>${s}</t></is></c>`).join('')}</row>`;
    const csv = toCSV(buildXlsx({ sheets: [{ name: 'S', data }] }));
    const fields = csv.trimEnd().split(',');
    expect(fields).toHaveLength(starts.length);
    for (const f of fields) expect(f.startsWith(`"'`)).toBe(true);
    expect(toCSV(buildXlsx({ sheets: [{ name: 'S', data }] }), { formulaInjection: 'off' })).not.toContain("'");
  });

  it('@xlvane/compat escapes them after removing NUL', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.addRow(['\u0000=cmd', '\u0001=cmd', '\u{FE66}cmd', '\u{FE62}1', '\u{FE63}1', '\u{FE6B}x', 'plain']);
    const text = new TextDecoder().decode(await wb.csv.writeBuffer());
    const fields = text.trimEnd().split(',');
    expect(fields.slice(0, 6).every((f) => f.startsWith("'") || f.startsWith(`"'`))).toBe(true);
    expect(fields[6]).toBe('plain');
    expect(text).not.toContain('\u0000');
  });
});
