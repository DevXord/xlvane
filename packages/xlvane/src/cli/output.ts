// Text output of the CLI: colour, tables and quoting (docs/decisions.md D-602).
import { styleText } from 'node:util';
import type { CliIO } from './types';

/**
 * Whether to colour a stream. `util.styleText` respects NO_COLOR, NODE_DISABLE_COLORS,
 * FORCE_COLOR and isTTY only from Node.js 22.8.0 (F-NODE-10), so the CLI applies the same rules
 * itself, as documented for Node.js (doc/api/cli.md, v22.x): FORCE_COLOR "1", "true", "" or
 * "2"/"3" turns colour on and NO_COLOR / NODE_DISABLE_COLORS are then ignored; any other
 * FORCE_COLOR value turns it off; NO_COLOR or NODE_DISABLE_COLORS with any value turn it off;
 * otherwise colour is used on a terminal only.
 */
export function useColor(env: CliIO['env'], isTTY: boolean): boolean {
  const force = env['FORCE_COLOR'];
  if (force !== undefined) return ['1', '2', '3', 'true', ''].includes(force);
  if (env['NO_COLOR'] !== undefined || env['NODE_DISABLE_COLORS'] !== undefined) return false;
  return isTTY;
}

/** Styled text for one stream. Colour only adds to the words; it never carries information alone. */
export class Output {
  constructor(
    readonly color: boolean,
    /** Width available for tables. */
    readonly width: number,
  ) {}

  private style(format: string | readonly string[], text: string): string {
    return this.color && text !== '' ? styleText(format, text, { validateStream: false }) : text;
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
   * Width is counted in UTF-16 code units, so wide characters may misalign (documented).
   */
  table(headers: readonly string[], rows: readonly (readonly string[])[], indent = ''): string {
    const count = headers.length;
    const widths = headers.map((h, i) => Math.max(h.length, ...rows.map((r) => oneLine(r[i] ?? '').length)));
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
          const padded = i === count - 1 ? text : text.padEnd(w);
          return head ? this.bold(padded) : padded;
        })
        .join(' '.repeat(gap))
        .trimEnd();
    return [line(headers, true), ...rows.map((r) => line(r, false))].join('\n');
  }
}

function oneLine(text: string): string {
  return text.replace(/\r\n|\r|\n/g, '⏎').replace(/\t/g, ' ');
}

function fit(text: string, width: number): string {
  return text.length <= width ? text : `${text.slice(0, Math.max(0, width - 1))}…`;
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

/** A cell value as text for a table. */
export function showValue(value: unknown): string {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return value.toISOString().slice(0, 23).replace(/T00:00:00\.000$/, '');
  if (typeof value === 'string') return value;
  return String(value);
}

/** "1 row", "2 rows". */
export function plural(n: number, word: string, many = `${word}s`): string {
  return `${String(n)} ${n === 1 ? word : many}`;
}
