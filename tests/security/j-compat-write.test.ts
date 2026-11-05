// Writing with @xlvane/exceljs takes near-linear time in the number of shared formula groups,
// array formulas and images (security review J, C5 and C7): the overlap check of shared formula
// ranges no longer compares every group with every array formula, and identical images are
// found by length and hash instead of comparing every pair of images byte by byte.
import { unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { sharedFormulaRanges } from '../../packages/exceljs/src/convert';
import { columnLetter } from '../../packages/exceljs/src/model';
import { ExcelJS } from '../exceljs/helpers';

const enc = new TextDecoder();

function sheetXml(bytes: Uint8Array): string {
  return enc.decode(unzipSync(bytes)['xl/worksheets/sheet1.xml']!);
}

type Box = { top: number; left: number; bottom: number; right: number };

/** The result of sharedFormulaRanges computed pair by pair, as the code before the fix did. */
function naiveRanges(groups: [string, Box][], arrays: Box[]): Map<string, string> {
  const overlaps = (a: Box, b: Box): boolean => a.top <= b.bottom && b.top <= a.bottom && a.left <= b.right && b.left <= a.right;
  const out = new Map<string, string>();
  for (const [m, b] of groups) {
    const bad = groups.some(([n, c]) => n !== m && overlaps(b, c)) || arrays.some((a) => overlaps(a, b));
    if (!bad) out.set(m, `${columnLetter(b.left)}${String(b.top)}:${columnLetter(b.right)}${String(b.bottom)}`);
  }
  return out;
}

describe('shared formula ranges', () => {
  it('60 000 rows with a shared and an array formula each are written in a few seconds', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    const n = 60_000;
    for (let r = 1; r <= n; r++) {
      ws.getCell(r, 1).value = { formula: `B${String(r)}+1`, shareType: 'shared', ref: `A${String(r)}` };
      ws.getCell(r, 3).value = { formula: `B${String(r)}*2`, shareType: 'array', ref: `C${String(r)}` };
    }
    const start = performance.now();
    const ranges = sharedFormulaRanges(ws);
    const bytes = await wb.xlsx.writeBuffer();
    expect(performance.now() - start).toBeLessThan(10_000);
    expect(ranges.size).toBe(n);
    const xml = sheetXml(bytes);
    expect(xml).toContain(`<f t="shared" ref="A${String(n)}" si="${String(n - 1)}">`);
    expect(xml).toContain(`<f t="array" ref="C${String(n)}">`);
  });

  it('60 000 nested tall groups are checked in a few seconds', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    const n = 60_000;
    // Group r covers A{r}:A{n}, so every group overlaps every other one.
    for (let r = 1; r <= n; r++) ws.getCell(r, 1).value = { formula: `B${String(r)}`, shareType: 'shared', ref: `A${String(r)}:A${String(n)}` };
    ws.getCell(1, 3).value = { formula: 'D1', shareType: 'shared', ref: 'C1:C5' };
    ws.getCell(1, 5).value = { formula: 'D1', shareType: 'shared', ref: 'E1:E5' };
    ws.getCell(5, 5).value = { formula: 'C1', shareType: 'array', ref: 'E5:F5' };
    const start = performance.now();
    const ranges = sharedFormulaRanges(ws);
    expect(performance.now() - start).toBeLessThan(3000);
    // Only the group in column C overlaps nothing.
    expect([...ranges]).toEqual([['C1', 'C1:C5']]);
  });

  it('gives the same groups as the pairwise check', () => {
    let seed = 7;
    const rnd = (k: number): number => {
      seed = (seed * 1_103_515_245 + 12_345) % 2 ** 31;
      return 1 + (seed % k);
    };
    for (let round = 0; round < 30; round++) {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('S');
      const groups: [string, Box][] = [];
      const arrays: Box[] = [];
      const used = new Set<string>();
      for (let i = 0; i < 40; i++) {
        const top = rnd(30);
        const left = rnd(12);
        const address = `${columnLetter(left)}${String(top)}`;
        if (used.has(address)) continue;
        used.add(address);
        const box = { top, left, bottom: top + rnd(4) - 1, right: left + rnd(3) - 1 };
        const ref = `${address}:${columnLetter(box.right)}${String(box.bottom)}`;
        if (rnd(4) === 1) {
          ws.getCell(top, left).value = { formula: '1+1', shareType: 'array', ref };
          arrays.push(box);
        } else {
          ws.getCell(top, left).value = { formula: '1+1', shareType: 'shared', ref };
          groups.push([address, box]);
        }
      }
      expect(sharedFormulaRanges(ws)).toEqual(naiveRanges(groups, arrays));
    }
  });
});

describe('identical images', () => {
  /** A PNG signature followed by `size - 8` bytes that differ from image to image only at the end. */
  function png(size: number, n: number): Uint8Array {
    const bytes = new Uint8Array(size);
    bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    let x = 12_345;
    for (let i = 8; i < size; i++) {
      x = (x * 1_103_515_245 + 12_345) >>> 0;
      bytes[i] = x >>> 24;
    }
    bytes[size - 4] = n & 0xff;
    bytes[size - 3] = (n >> 8) & 0xff;
    return bytes;
  }

  it('400 different images of the same length are written in a few seconds', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    for (let i = 0; i < 400; i++) ws.addImage(wb.addImage({ buffer: png(64 * 1024, i), extension: 'png' }), `A${String(2 * i + 1)}:B${String(2 * i + 2)}`);
    const start = performance.now();
    const bytes = await wb.xlsx.writeBuffer();
    expect(performance.now() - start).toBeLessThan(15_000);
    expect(Object.keys(unzipSync(bytes)).filter((p) => p.startsWith('xl/media/'))).toHaveLength(400);
  });

  it('still writes identical images once', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    const ids = [png(1000, 1), png(1000, 2), png(1000, 1), png(999, 1), png(1000, 2)].map((b) => wb.addImage({ buffer: b, extension: 'png' }));
    ids.forEach((id: number, i: number) => ws.addImage(id, `A${String(2 * i + 1)}:B${String(2 * i + 2)}`));
    const files = unzipSync(await wb.xlsx.writeBuffer());
    expect(Object.keys(files).filter((p) => p.startsWith('xl/media/'))).toHaveLength(3);
  });
});
