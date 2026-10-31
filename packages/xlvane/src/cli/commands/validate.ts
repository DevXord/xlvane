import { inspect } from '../../inspect';
import { readTable } from '../../table';
import { suggest } from '../../util/suggest';
import { flag, headerOption, str, usageError } from '../args';
import { readBook } from '../files';
import { plural } from '../output';
import type { CommandSpec } from '../types';
import { warningsBlock } from '../warnings';

const TYPES = ['number', 'string', 'date', 'boolean', 'any'] as const;
type ColumnType = (typeof TYPES)[number];

/** One rule of --columns: "name:type", optional when the type ends with "?". */
interface ColumnRule {
  readonly name: string;
  readonly type: ColumnType;
  readonly optional: boolean;
}

/** A problem found by validate. */
export interface ValidationProblem {
  readonly kind: 'missing-column' | 'empty' | 'type';
  readonly column: string;
  readonly cell: string | null;
  readonly row: number | null;
  readonly message: string;
}

function parseRules(text: string): ColumnRule[] {
  return text.split(',').map((part) => {
    const m = /^\s*([^:]+?)\s*(?::\s*([A-Za-z]+)(\?)?\s*)?$/.exec(part);
    const type = (m?.[2] ?? 'any').toLowerCase() as ColumnType;
    if (!m || !(TYPES as readonly string[]).includes(type)) {
      const close = m?.[2] !== undefined ? suggest(m[2], TYPES, 1) : [];
      const didYouMean = close.length > 0 ? `Did you mean "${m![1]!}:${close[0]!}${m![3] ?? ''}"? ` : '';
      throw usageError(`Cannot read the column rule "${part.trim()}".`, `${didYouMean}Write rules as "name:type" separated by commas; types: ${TYPES.join(', ')}; add "?" to allow empty cells, e.g. "id:number,note:string?".`, {
        command: 'validate',
        option: '--columns',
      });
    }
    return { name: m[1]!, type, optional: m[3] === '?' };
  });
}

function typeOf(v: unknown): Exclude<ColumnType, 'any'> | 'empty' {
  if (v === null || v === undefined || v === '') return 'empty';
  if (typeof v === 'number') return 'number';
  if (typeof v === 'boolean') return 'boolean';
  if (v instanceof Date) return 'date';
  return 'string';
}

export const validateCommand: CommandSpec = {
  name: 'validate',
  summary: 'Check a workbook and the types of its columns',
  description:
    'Checks that a file can be read safely and reports warnings (macros, external links, damaged parts). With --columns it also checks a table: every listed column must exist in the header row, its cells must not be empty unless the type ends with "?", and they must have the given type (number, string, date, boolean or any). Each problem names the cell. The exit code is 1 when problems are found; warnings count only with --strict.',
  usage: '<file> [options]',
  positionals: { min: 1, max: 1, names: ['file'] },
  checks: 'problems were found',
  options: {
    sheet: { type: 'string', value: 'name', description: 'Sheet to check with --columns (default: the first visible sheet)' },
    columns: { type: 'string', value: 'name:type,...', description: 'Column rules, e.g. "Region:string,Revenue:number"' },
    header: { type: 'string', value: 'auto|row', description: 'Header row for --columns (default: detected)' },
    strict: { type: 'boolean', description: 'Treat warnings as problems' },
  },
  examples: [
    { command: 'xlvane validate report.xlsx', description: 'Check that the file reads cleanly' },
    { command: 'xlvane validate report.xlsx --sheet Sales --columns "Region:string,Revenue:number"', description: 'Check the columns of a table' },
    { command: 'xlvane validate report.xlsx --strict --json', description: 'Fail on warnings too; problems as JSON' },
  ],
  related: ['inspect', 'read'],
  async run(ctx) {
    const file = ctx.args.positionals[0]!;
    const rulesText = str(ctx.args, 'columns');
    const rules = rulesText !== undefined ? parseRules(rulesText) : [];
    const header = headerOption(ctx.args, 'validate');
    if (header === false) throw usageError('"xlvane validate" needs a header row to match --columns.', 'Use --header auto (default) or the row number.', { command: 'validate', option: '--header' });
    const { workbook } = await readBook(ctx, file);
    const warnings = inspect(workbook).warnings;
    const problems: ValidationProblem[] = [];
    let sheetName: string | null = null;
    let rowsChecked = 0;
    if (rules.length > 0) {
      const sheet = str(ctx.args, 'sheet');
      const table = readTable(workbook, { ...(sheet !== undefined ? { sheet } : {}), header });
      sheetName = table.sheet;
      rowsChecked = table.rows.length;
      for (const rule of rules) {
        const column = table.columns.find((c) => c.header === rule.name);
        if (!column) {
          const headers = table.columns.map((c) => c.header).filter((h): h is string => h !== null);
          const close = suggest(rule.name, headers, 1);
          const didYouMean = close.length > 0 ? ` Did you mean "${close[0]!}"?` : '';
          problems.push({ kind: 'missing-column', column: rule.name, cell: null, row: table.headerRow, message: `Column "${rule.name}" is missing.${didYouMean} Header row: ${headers.map((h) => `"${h}"`).join(', ') || '(none)'}.` });
          continue;
        }
        table.rows.forEach((row, i) => {
          const rowNumber = table.rowNumbers[i]!;
          const cell = `${column.letter}${String(rowNumber)}`;
          const t = typeOf(row[column.key]);
          if (t === 'empty') {
            if (!rule.optional) problems.push({ kind: 'empty', column: rule.name, cell, row: rowNumber, message: `${cell} is empty; "${rule.name}" needs a value.` });
          } else if (rule.type !== 'any' && t !== rule.type) {
            problems.push({ kind: 'type', column: rule.name, cell, row: rowNumber, message: `${cell} is a ${t}; "${rule.name}" needs a ${rule.type}.` });
          }
        });
      }
    }
    const strict = flag(ctx.args, 'strict');
    const ok = problems.length === 0 && (!strict || warnings.length === 0);
    const result = { file, ok, sheet: sheetName, rowsChecked, problems, warnings };
    return {
      result,
      exitCode: ok ? 0 : 1,
      render: (out) => {
        const lines: string[] = [];
        if (problems.length > 0) {
          const shown = problems.slice(0, 100);
          const cellWidth = Math.max(0, ...shown.map((p) => (p.cell ?? '').length));
          lines.push(out.bold(out.red(`Problems (${String(problems.length)}):`)));
          // The cell first, aligned, then the message without the cell it starts with.
          lines.push(
            ...shown.map((p) => {
              const text = p.cell !== null && p.message.startsWith(`${p.cell} `) ? p.message.slice(p.cell.length + 1) : p.message;
              return `  ${out.red('problem')}  ${p.cell === null ? ''.padEnd(cellWidth) : out.bold(p.cell.padEnd(cellWidth))}  ${text}`.replace(/ +$/, '');
            }),
          );
          if (problems.length > 100) lines.push(out.dim(`  … ${String(problems.length - 100)} more; use --json for all`));
          lines.push('');
        }
        if (warnings.length > 0) lines.push(...warningsBlock(out, warnings), '');
        const checked = sheetName !== null ? `, ${plural(rowsChecked, 'row')} of sheet "${sheetName}" checked against ${plural(rules.length, 'column rule')}` : '';
        if (ok && warnings.length > 0) lines.push(`${out.yellow(`OK with ${plural(warnings.length, 'warning')}`)}: ${file}${checked}. Use --strict to fail on warnings.`);
        else if (ok) lines.push(`${out.green('OK')}: ${file}${checked}.`);
        else lines.push(`${out.red('Failed')}: ${file}${checked}${strict && problems.length === 0 ? ' (warnings count with --strict)' : ''}.`);
        return lines.join('\n');
      },
    };
  },
};
