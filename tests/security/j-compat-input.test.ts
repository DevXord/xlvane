// xlsx.read(stream) and stream.xlsx.WorkbookReader with an async iterable stop reading as soon as
// the input passes maxInputBytes (256 MiB), instead of collecting the whole stream first
// (security review J, C8).
import { describe, expect, it } from 'vitest';
import { DEFAULT_READ_LIMITS } from '../../packages/xlvane/src/limits';
import { buildXlsx } from '../helpers/xlsx-builder';
import { ExcelJS } from '../exceljs/helpers';

const MiB = 1024 * 1024;

/** 1 MiB chunks (the same array each time), at most `max` of them, counting those pulled. */
function chunks(max: number): { source: AsyncIterable<Uint8Array>; pulled: () => number } {
  const chunk = new Uint8Array(MiB);
  let pulled = 0;
  return {
    source: {
      async *[Symbol.asyncIterator]() {
        while (pulled < max) {
          pulled++;
          yield chunk;
        }
      },
    },
    pulled: () => pulled,
  };
}

async function errorOf(run: () => Promise<unknown>): Promise<any> {
  try {
    await run();
  } catch (e) {
    return e;
  }
  return null;
}

describe('input read from a stream', () => {
  const limitChunks = DEFAULT_READ_LIMITS.maxInputBytes / MiB;

  it('xlsx.read stops after maxInputBytes and one chunk', async () => {
    const { source, pulled } = chunks(600);
    const error = await errorOf(() => new ExcelJS.Workbook().xlsx.read(source));
    expect(error?.code).toBe('XLV_E_LIMIT_EXCEEDED');
    expect(error.details).toEqual({ limit: 'maxInputBytes', value: DEFAULT_READ_LIMITS.maxInputBytes });
    expect(pulled()).toBe(limitChunks + 1);
  });

  it('stream.xlsx.WorkbookReader stops after maxInputBytes and one chunk', async () => {
    const { source, pulled } = chunks(600);
    const error = await errorOf(async () => {
      for await (const ws of new ExcelJS.stream.xlsx.WorkbookReader(source, {})) void ws;
    });
    expect(error?.code).toBe('XLV_E_LIMIT_EXCEEDED');
    expect(error.details).toEqual({ limit: 'maxInputBytes', value: DEFAULT_READ_LIMITS.maxInputBytes });
    expect(pulled()).toBe(limitChunks + 1);
  });

  it('still reads a workbook from a stream', async () => {
    const bytes = buildXlsx({ sheets: [{ name: 'S', data: '<row r="1"><c r="A1"><v>7</v></c></row>' }] });
    const parts = async function* (): AsyncGenerator<Uint8Array> {
      for (let i = 0; i < bytes.length; i += 100) yield bytes.subarray(i, i + 100);
    };
    const wb = await new ExcelJS.Workbook().xlsx.read(parts());
    expect(wb.getWorksheet('S').getCell('A1').value).toBe(7);
    let value: unknown;
    for await (const ws of new ExcelJS.stream.xlsx.WorkbookReader(parts(), {})) for await (const row of ws) value = row.getCell(1).value;
    expect(value).toBe(7);
  });
});
