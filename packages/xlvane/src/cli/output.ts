// Text output of the CLI: colour, tables and quoting (docs/decisions.md D-602).
import { isoDate } from '../util/iso-date';
import { EAST_ASIAN_WIDE } from './east-asian-width';
import { nodeUtil } from './node-util';
import type { CliIO } from './types';

/**
 * Whether to colour a stream. `util.styleText` respects NO_COLOR, NODE_DISABLE_COLORS,
 * FORCE_COLOR and isTTY only from Node.js 22.8.0 (F-NODE-10), so the CLI applies the same rules
 * itself, as documented for Node.js (doc/api/cli.md, v22.x): FORCE_COLOR "1", "true", "" or
 * "2"/"3" turns colour on and NO_COLOR / NODE_DISABLE_COLORS are then ignored; any other
 * FORCE_COLOR value turns it off; NO_COLOR or NODE_DISABLE_COLORS with any value, and
 * TERM=dumb (a terminal without ANSI colour, terminfo), turn it off; otherwise colour is used on
 * a terminal only. TERM=dumb follows `getColorDepth` in Node.js lib/internal/tty.js (v22.x and
 * v24.x, read 2026-10-31).
 */
export function useColor(env: CliIO['env'], isTTY: boolean): boolean {
  const force = env['FORCE_COLOR'];
  if (force !== undefined) return ['1', '2', '3', 'true', ''].includes(force);
  if (env['NO_COLOR'] !== undefined || env['NODE_DISABLE_COLORS'] !== undefined || env['TERM'] === 'dumb') return false;
  return isTTY;
}

/**
 * Width for tables and wrapped text: the terminal's width when it reports one above 0 (some
 * pseudo-terminals report 0), otherwise COLUMNS, otherwise 100.
 */
export function termWidth(env: CliIO['env'], columns: number | undefined, isTTY: boolean): number {
  if (isTTY && columns !== undefined && columns > 0) return columns;
  const fromEnv = Number(env['COLUMNS']);
  if (isTTY && Number.isInteger(fromEnv) && fromEnv > 0) return fromEnv;
  return 100;
}

/** Styled text for one stream. Colour only adds to the words; it never carries information alone. */
export class Output {
  constructor(
    readonly color: boolean,
    /** Width available for tables. */
    readonly width: number,
  ) {}

  private style(format: string | readonly string[], text: string): string {
    return this.color && text !== '' ? nodeUtil.styleText(format, text, { validateStream: false }) : text;
  }

  bold(text: string): string {
    return this.style('bold', text);
  }
  dim(text: string): string {
    return this.style('dim', text);
  }
  red(text: string): string {
    return this.style('red', text);
  }
  green(text: string): string {
    return this.style('green', text);
  }
  yellow(text: string): string {
    return this.style('yellow', text);
  }
  cyan(text: string): string {
    return this.style('cyan', text);
  }

  /**
   * An aligned table. Columns are cut with "…" to fit `width`; the widest columns shrink first.
   * Widths are display columns ({@link displayWidth}), so CJK text and emoji stay aligned.
   */
  table(headers: readonly string[], rows: readonly (readonly string[])[], indent = '', align: readonly ('left' | 'right')[] = []): string {
    const count = headers.length;
    const widths = headers.map((h, i) => Math.max(displayWidth(oneLine(h)), ...rows.map((r) => displayWidth(oneLine(r[i] ?? '')))));
    const gap = 2;
    const available = Math.max(count * 4, this.width - indent.length - gap * (count - 1));
    let total = widths.reduce((a, b) => a + b, 0);
    while (total > available) {
      let widest = 0;
      for (let i = 1; i < count; i++) if ((widths[i] ?? 0) > (widths[widest] ?? 0)) widest = i;
      if ((widths[widest] ?? 0) <= 4) break;
      widths[widest] = (widths[widest] ?? 0) - 1;
      total--;
    }
    const line = (cells: readonly string[], head: boolean): string =>
      indent +
      cells
        .map((c, i) => {
          const w = widths[i] ?? 0;
          const text = fit(oneLine(c), w);
          const pad = ' '.repeat(Math.max(0, w - displayWidth(text)));
          const padded = align[i] === 'right' && !head ? pad + text : i === count - 1 ? text : text + pad;
          return head ? this.bold(padded) : padded;
        })
        .join(' '.repeat(gap))
        .trimEnd();
    return [line(headers, true), ...rows.map((r) => line(r, false))].join('\n');
  }
}

function oneLine(text: string): string {
  return escapeControls(text.replace(/\r\n|\r|\n/g, '⏎').replace(/\t/g, ' '));
}

/**
 * Control characters that a terminal would act on (C0 except tab and line feed, DEL, and C1,
 * which includes the 8-bit CSI U+009B) written as `\x1b`, and bidirectional formatting characters
 * written as `\u202e`, so text from a file, such as a cell
 * holding `_x001B_]0;title_x0007_`, cannot move the cursor, change the window title or hide
 * output (docs/decisions.md D-R02). Used for all human-readable output; `--json` is unchanged.
 */
export function escapeControls(text: string): string {
  return text.replace(/[\u0000-\u0008\u000b-\u001f\u007f-\u009f]/g, (c) => `\\x${c.charCodeAt(0).toString(16).padStart(2, '0')}`).replace(BIDI, (c) => `\\u${c.charCodeAt(0).toString(16).padStart(4, '0')}`);
}

/**
 * The characters with the Unicode property Bidi_Control (checked against `/\p{Bidi_Control}/u` in
 * Node.js 22, Unicode 16.0; tests/cli/release-fixes.test.ts). They reorder the text shown after
 * them, so a cell holding "abc\u202Egnp.exe" would be displayed as "abcexe.png". Written as
 * `\u202e` in human-readable output, like other control characters (D-R08).
 */
const BIDI = /[\u061c\u200e\u200f\u202a-\u202e\u2066-\u2069]/g;

/** The colour sequences `Output` writes (styleText: bold, dim, red, green, yellow, cyan and their resets). */
const OWN_STYLES = /\u001b\[(?:1|2|22|31|32|33|36|39)m/y;

/**
 * Final filter for human-readable output: escapes every control character (see
 * {@link escapeControls}) except line feeds, tabs and, when `keepStyles` is true, the exact
 * colour sequences that `Output` itself writes. Applied to everything the CLI prints for people.
 */
export function terminalSafe(text: string, keepStyles: boolean): string {
  let result = '';
  let start = 0;
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    const control =
      (code < 0x20 && code !== 0x0a && code !== 0x09) ||
      (code >= 0x7f && code <= 0x9f) ||
      code === 0x061c ||
      code === 0x200e ||
      code === 0x200f ||
      (code >= 0x202a && code <= 0x202e) ||
      (code >= 0x2066 && code <= 0x2069);
    if (!control) continue;
    if (keepStyles && code === 0x1b) {
      OWN_STYLES.lastIndex = i;
      const m = OWN_STYLES.exec(text);
      if (m !== null) {
        i += m[0].length - 1;
        continue;
      }
    }
    result += text.slice(start, i) + escapeControls(text[i]!);
    start = i + 1;
  }
  return start === 0 ? text : result + text.slice(start);
}

/** Cuts text to `width` display columns with "…", between grapheme clusters. */
function fit(text: string, width: number): string {
  if (displayWidth(text) <= width) return text;
  let result = '';
  let used = 0;
  for (const part of graphemes(text)) {
    const w = displayWidth(part);
    if (used + w > width - 1) break;
    result += part;
    used += w;
  }
  return `${result}…`;
}

/**
 * Grapheme clusters of a text (Intl.Segmenter), or its code points when the Node.js build has no
 * Intl support, so that cutting never splits a surrogate pair or a letter from its accent.
 */
function graphemes(text: string): Iterable<string> {
  if (segmenter === undefined) return Array.from(text);
  return Array.from(segmenter.segment(text), (s) => s.segment);
}

const segmenter = typeof Intl === 'object' && typeof Intl.Segmenter === 'function' ? new Intl.Segmenter(undefined, { granularity: 'grapheme' }) : undefined;

/**
 * The number of terminal columns a text takes, counted as Node.js counts them for its own output
 * (src/node_i18n.cc GetColumnWidth, v22.23.4; F-UNICODE-01): East Asian Wide and Fullwidth
 * characters (Unicode East_Asian_Width W and F) and emoji with Emoji_Presentation take two
 * columns; control and format characters (except the soft hyphen), nonspacing and enclosing marks
 * and emoji modifiers that are not Wide take none; every other character, including East Asian
 * Ambiguous ones, takes one. Each code point of an emoji sequence is counted, as Node.js does by
 * default, so a terminal that draws the sequence as one picture shows it narrower. Colour codes
 * (`ESC [ … m`, as `Output` writes them) take no columns, as Node.js strips them first.
 */
export function displayWidth(text: string): number {
  if (text.includes('\u001b')) text = text.replace(/\u001b\[[0-9;]*m/g, '');
  let width = 0;
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    if (code >= 0x20 && code < 0x7f) width++;
    else {
      for (const ch of text.slice(i)) width += codePointWidth(ch.codePointAt(0) ?? 0, ch);
      break;
    }
  }
  return width;
}

function codePointWidth(code: number, ch: string): number {
  if (isEastAsianWide(code) || EMOJI_PRESENTATION.test(ch)) return 2;
  if (code !== 0xad && ZERO_WIDTH.test(ch)) return 0;
  return 1;
}

const EMOJI_PRESENTATION = /^\p{Emoji_Presentation}$/u;
const ZERO_WIDTH = /^[\p{Cc}\p{Cf}\p{Me}\p{Mn}\p{Emoji_Modifier}]$/u;

/** Binary search in the sorted [first, last, first, last, …] ranges of EAST_ASIAN_WIDE. */
function isEastAsianWide(code: number): boolean {
  if (code < 0x1100) return false;
  let lo = 0;
  let hi = EAST_ASIAN_WIDE.length / 2 - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    if (code < EAST_ASIAN_WIDE[mid * 2]!) hi = mid - 1;
    else if (code > EAST_ASIAN_WIDE[mid * 2 + 1]!) lo = mid + 1;
    else return true;
  }
  return false;
}

/** Quotes an argument for a POSIX shell when needed, so "Try:" commands can be copied. */
export function quoteArg(arg: string): string {
  if (/^[\w@%+=:,./-]+$/.test(arg)) return arg;
  return `'${arg.replace(/'/g, `'\\''`)}'`;
}

/** A command line as text. */
export function commandLine(args: readonly string[]): string {
  return ['xlvane', ...args].map(quoteArg).join(' ');
}

/**
 * A cell value as text for a table: numbers to 15 significant digits (what spreadsheet
 * applications show, so 0.1 + 0.2 reads 0.3), dates as `2024-01-31` or `2024-01-31 13:45:00`,
 * times of day alone (serial dates before 1900-01-01) as `13:45:00`.
 */
export function showValue(value: unknown): string {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) {
    const text = isoDate(value).replace('T', ' ');
    return value.getUTCFullYear() < 1900 && text.includes(' ') ? text.slice(text.indexOf(' ') + 1) : text;
  }
  if (typeof value === 'number') return Number.isFinite(value) ? String(Number(value.toPrecision(15))) : String(value);
  if (typeof value === 'string') return value;
  return String(value);
}

/** "812 bytes", "2.8 kB", "3.4 MB" (powers of 1000). */
export function humanSize(bytes: number): string {
  if (bytes < 1000) return plural(bytes, 'byte');
  if (bytes < 1_000_000) return `${(bytes / 1000).toFixed(1)} kB`;
  return `${(bytes / 1_000_000).toFixed(1)} MB`;
}

/** Wraps a paragraph at `width` display columns; lines after the first start with `indent`. */
export function wrap(text: string, width: number, indent = '', firstIndent: string = indent): string {
  const lines: string[] = [];
  let line = '';
  let prefix = firstIndent;
  for (const word of text.split(/\s+/).filter((w) => w !== '')) {
    if (line !== '' && displayWidth(prefix) + displayWidth(line) + 1 + displayWidth(word) > width) {
      lines.push(prefix + line);
      prefix = indent;
      line = word;
    } else line = line === '' ? word : `${line} ${word}`;
  }
  if (line !== '') lines.push(prefix + line);
  return lines.join('\n');
}

/** "1 row", "2 rows". */
export function plural(n: number, word: string, many = `${word}s`): string {
  return `${String(n)} ${n === 1 ? word : many}`;
}
