// Shared helpers for the ported exceljs specs.
// Style fixtures ported from exceljs v4.4.0 spec/utils/data/styles.json
// (MIT, Copyright (c) 2014-2019 Guyon Roche), changes: typed as TypeScript; only the entries the
// ported specs use.
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { expect } from 'vitest';
import ExcelJSModule from '../../packages/compat/src/index';

/** The package under test, typed loosely: the specs exercise the exceljs API shape. */
export const ExcelJS: any = ExcelJSModule;

export const styles = {
  numFmts: {
    numFmt1: '# ?/?',
    numFmt2: '[Green]#,##0 ;[Red](#,##0)',
  },
  fonts: {
    arialBlackUI14: { name: 'Arial Black', family: 2, size: 14, underline: true, italic: true },
    comicSansUdB16: { name: 'Comic Sans MS', family: 4, size: 16, underline: 'double', bold: true },
    broadwayRedOutline20: { name: 'Broadway', family: 5, size: 20, outline: true, color: { argb: 'FFFF0000' } },
  },
  alignments: [
    { text: 'Top Left', alignment: { horizontal: 'left', vertical: 'top' } },
    { text: 'Middle Centre', alignment: { horizontal: 'center', vertical: 'middle' } },
    { text: 'Bottom Right', alignment: { horizontal: 'right', vertical: 'bottom' } },
  ],
  namedAlignments: {
    topLeft: { horizontal: 'left', vertical: 'top' },
    middleCentre: { horizontal: 'center', vertical: 'middle' },
    bottomRight: { horizontal: 'right', vertical: 'bottom' },
  },
  borders: {
    thin: { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } },
    doubleRed: {
      top: { style: 'double', color: { argb: 'FFFF0000' } },
      left: { style: 'double', color: { argb: 'FFFF0000' } },
      bottom: { style: 'double', color: { argb: 'FFFF0000' } },
      right: { style: 'double', color: { argb: 'FFFF0000' } },
    },
  },
  fills: {
    redDarkVertical: { type: 'pattern', pattern: 'darkVertical', fgColor: { argb: 'FFFF0000' } },
    redGreenDarkTrellis: { type: 'pattern', pattern: 'darkTrellis', fgColor: { argb: 'FFFF0000' }, bgColor: { argb: 'FF00FF00' } },
    blueWhiteHGrad: {
      type: 'gradient',
      gradient: 'angle',
      degree: 0,
      stops: [
        { position: 0, color: { argb: 'FF0000FF' } },
        { position: 1, color: { argb: 'FFFFFFFF' } },
      ],
    },
    rgbPathGrad: {
      type: 'gradient',
      gradient: 'path',
      center: { left: 0.5, top: 0.5 },
      stops: [
        { position: 0, color: { argb: 'FFFF0000' } },
        { position: 0.5, color: { argb: 'FF00FF00' } },
        { position: 1, color: { argb: 'FF0000FF' } },
      ],
    },
  },
} as const;

/** A temporary directory for files written by a test; removed afterwards. */
export async function withTempDir<T>(run: (dir: string) => Promise<T>): Promise<T> {
  const dir = await mkdtemp(join(tmpdir(), 'xlvane-exceljs-'));
  try {
    return await run(dir);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

/** Writes `wb` to a temporary .xlsx file and reads it back into a new Workbook. */
export async function roundTrip(wb: any): Promise<any> {
  return withTempDir(async (dir) => {
    const file = join(dir, 'wb.test.xlsx');
    await wb.xlsx.writeFile(file);
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(file);
    return wb2;
  });
}

/** Expects `fn` to throw (or reject with) an XlvError with code XLV_E_NOT_SUPPORTED. */
export async function expectNotSupported(fn: () => unknown): Promise<void> {
  let thrown: unknown;
  try {
    await fn();
  } catch (error) {
    thrown = error;
  }
  expect(thrown, 'expected XLV_E_NOT_SUPPORTED').toBeDefined();
  expect((thrown as { code?: unknown }).code).toBe('XLV_E_NOT_SUPPORTED');
}
