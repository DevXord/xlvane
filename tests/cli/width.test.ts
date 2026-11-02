// Display width of text in the terminal: tables, cut text and wrapped paragraphs count columns as
// Node.js does (src/node_i18n.cc GetColumnWidth, v22.23.4; F-UNICODE-01): East Asian Wide and
// Fullwidth characters and emoji with Emoji_Presentation take two columns; control, format,
// nonspacing and enclosing marks and emoji modifiers that are not Wide take none (soft hyphen takes one); every other
// character, including East Asian Ambiguous ones, takes one.
import { writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { EAST_ASIAN_WIDE } from '../../packages/xlvane/src/cli/east-asian-width';
import { displayWidth, Output, wrap } from '../../packages/xlvane/src/cli/output';
import { cli, withFixtures } from './helpers';

describe('displayWidth', () => {
  it('counts columns, not UTF-16 code units', () => {
    expect(displayWidth('abc')).toBe(3);
    expect(displayWidth('')).toBe(0);
    expect(displayWidth('中文')).toBe(4); // U+4E2D, U+6587: W
    expect(displayWidth('한국어')).toBe(6); // Hangul syllables: W
    expect(displayWidth('ｶﾀｶﾅ')).toBe(4); // halfwidth katakana U+FF76…: H
    expect(displayWidth('ＡＢ')).toBe(4); // fullwidth Latin U+FF21, U+FF22: F
    expect(displayWidth('𠀀')).toBe(2); // U+20000, CJK Extension B (plane 2): W
    expect(displayWidth('…')).toBe(1); // U+2026: A, narrow by default as in Node.js
    expect(displayWidth('é')).toBe(1); // e + U+0301 combining acute (Mn)
    expect(displayWidth('a​b')).toBe(2); // zero width space (Cf)
    expect(displayWidth('soft­hyphen')).toBe(11); // U+00AD is Cf but takes one column
    expect(displayWidth('😀')).toBe(2); // U+1F600: W and Emoji_Presentation
    // The emoji modifier U+1F3FD is East Asian Wide (EastAsianWidth-17.0.0.txt: 1F3FB..1F3FF ; W), and
    // Node.js checks W before Emoji_Modifier, so it counts two columns too.
    expect(displayWidth('👍🏽')).toBe(4);
    expect(displayWidth('❤')).toBe(1); // U+2764 without a variation selector: N, no Emoji_Presentation
    expect(displayWidth('\\x1b')).toBe(4); // escaped control characters are plain ASCII
  });

  it('uses a sorted table of East Asian Wide and Fullwidth ranges', () => {
    expect(EAST_ASIAN_WIDE.length % 2).toBe(0);
    for (let i = 1; i < EAST_ASIAN_WIDE.length; i++) expect(EAST_ASIAN_WIDE[i]!).toBeGreaterThan(EAST_ASIAN_WIDE[i - 1]! - (i % 2 === 1 ? 1 : 0));
    expect(EAST_ASIAN_WIDE[EAST_ASIAN_WIDE.length - 1]).toBeLessThanOrEqual(0x10ffff);
  });
});

describe('tables with wide characters', () => {
  /** The column at which each cell after the first starts, in display columns. */
  const starts = (line: string, cells: readonly string[]): number[] => {
    const result: number[] = [];
    let from = 0;
    for (const cell of cells.slice(1)) {
      const at = line.indexOf(cell, from);
      result.push(displayWidth(line.slice(0, at)));
      from = at + cell.length;
    }
    return result;
  };

  it('align columns by display width', () => {
    const out = new Output(false, 80);
    const rows = [
      ['北京', '1200', 'x'],
      ['Warszawa', '980', 'y'],
      ['東京都', '15', 'z'],
    ];
    const lines = out.table(['City', 'Revenue', 'Flag'], rows, '', ['left', 'right', 'left']).split('\n');
    const expected = starts(lines[0]!, ['City', 'Revenue', 'Flag']);
    rows.forEach((r, i) => {
      const line = lines[i + 1]!;
      // Right-aligned numbers end where the header ends; the third column starts with the header.
      expect(displayWidth(line.slice(0, line.indexOf(r[1]!) + r[1]!.length))).toBe(expected[0]! + 'Revenue'.length);
      expect(starts(line, r)[1]).toBe(expected[1]);
    });
  });

  it('cut cells to the column width without splitting characters', () => {
    const out = new Output(false, 14);
    const text = out.table(['Name', 'Note'], [['名前', '中文中文中文中文中文'], ['😀😀😀😀😀😀', 'ok']]);
    for (const line of text.split('\n')) {
      expect(displayWidth(line)).toBeLessThanOrEqual(14);
      expect(line).not.toMatch(/[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/);
    }
    expect(text).toContain('…');
  });
});

describe('headers and styled text', () => {
  it('a header with a line break or tab is measured as it is printed (found by review)', () => {
    const out = new Output(false, 80);
    const text = out.table(['Total\nsales', 'Q'], [['12', 'x']]);
    expect(text.split('\n')[0]).toBe('Total⏎sales  Q');
  });

  it('colour codes take no columns', () => {
    expect(displayWidth('\u001b[1mFix:\u001b[22m')).toBe(4);
    const wrapped = wrap('aaaa bbbb cccc dddd', 19, '      ', '\u001b[1mFix:\u001b[22m  ');
    expect(wrapped.split('\n')[0]).toBe('\u001b[1mFix:\u001b[22m  aaaa bbbb');
  });
});

describe('wrap', () => {
  it('measures words by display width', () => {
    const text = wrap('漢字漢字 漢字漢字 漢字漢字', 18);
    for (const line of text.split('\n')) expect(displayWidth(line)).toBeLessThanOrEqual(18);
    expect(text.split('\n')).toHaveLength(2);
  });
});

describe('read command', () => {
  it('prints a sheet with Chinese text in aligned columns', async () => {
    await withFixtures(async (dir) => {
      await writeFile(join(dir, 'cjk.csv'), 'City,Revenue\r\n北京,1200\r\nWarszawa,980\r\n');
      expect((await cli(['convert', 'cjk.csv', 'cjk.xlsx'], { cwd: dir })).code).toBe(0);
      const r = await cli(['read', 'cjk.xlsx'], { cwd: dir });
      expect(r.code).toBe(0);
      const lines = r.stdout.split('\n').filter((l) => /1200|980|Revenue/.test(l));
      const ends = lines.map((l) => displayWidth(l.trimEnd()));
      expect(new Set(ends).size).toBe(1);
    });
  });
});
