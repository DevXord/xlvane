// Values from the file are shortened when messages and error details show them (1.2.2 security
// review, item K10): a declared encoding or an attribute of megabytes must not be copied into
// every message, log line and report.
import { describe, expect, it } from 'vitest';
import { XlvError } from '../../packages/xlvane/src/errors';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { decodeXml } from '../../packages/xlvane/src/xml/parser';
import { expectXlvError } from '../helpers/expect-error';
import { buildXlsx } from '../helpers/xlsx-builder';

/** Letters that compress less than 200:1, so that the reader accepts the file. */
function noise(length: number): string {
  let seed = 11;
  let out = '';
  while (out.length < length) {
    seed = (seed * 1_103_515_245 + 12_345) % 2_147_483_648;
    out += 'abcdefghijklmnopqrstuvwxyz'[seed % 26]!.repeat(1 + (seed % 7));
  }
  return out.slice(0, length);
}

describe('values from the file in messages (K10)', () => {
  it('shortens a long declared encoding in the message and details', () => {
    const declared = 'x'.repeat(5_000_000);
    const error = expectXlvError(() => decodeXml(new TextEncoder().encode(`<?xml version="1.0" encoding="${declared}"?><r/>`), '/x.xml'), 'XLV_E_XML_ENCODING');
    expect(error.message.length).toBeLessThan(400);
    expect(String(error.details['encoding']).length).toBeLessThanOrEqual(201);
    expect(String(error.details['encoding']).startsWith('xxxx')).toBe(true);
    // A short one is shown as is.
    const short = expectXlvError(() => decodeXml(new TextEncoder().encode('<?xml version="1.0" encoding="latin1"?><r/>'), '/x.xml'), 'XLV_E_XML_ENCODING');
    expect(short.message).toContain('"latin1"');
    expect(short.details['encoding']).toBe('latin1');
  });

  it('shortens a long declared encoding of a workbook part', () => {
    const bytes = buildXlsx({ sheets: [{ name: 'S', data: `<?xml version="1.0" encoding="${noise(1_000_000)}"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>` }] });
    let thrown: unknown;
    try {
      readWorkbook(bytes);
    } catch (error) {
      thrown = error;
    }
    expect(thrown).toBeInstanceOf(XlvError);
    expect((thrown as XlvError).code).toBe('XLV_E_XML_ENCODING');
    expect((thrown as XlvError).message.length).toBeLessThan(400);
    expect(JSON.stringify((thrown as XlvError).details).length).toBeLessThan(400);
  });

  it('shortens long date1904 and sheet state values in warnings', () => {
    const long = noise(1_000_000);
    const bytes = buildXlsx({ workbookPr: `date1904="${long}"`, sheets: [{ name: 'S', data: '', state: long }] });
    const book = readWorkbook(bytes);
    expect(book.dateSystem).toBe('1900');
    expect(book.sheets[0]?.state).toBe('visible');
    const shown = book.warnings.filter((w) => w.code === 'XLV_W_INVALID_VALUE');
    expect(shown.map((w) => w.message.slice(0, 20))).toEqual(['workbookPr@date1904 ', 'Sheet "S" has the un']);
    for (const w of shown) {
      expect(w.message.length).toBeLessThan(400);
      expect(w.message).toContain(long.slice(0, 200));
    }
  });
});
