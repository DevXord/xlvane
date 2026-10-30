import { inspect } from '../../inspect';
import { readTable } from '../../table';
import { flag, headerOption, str, usageError } from '../args';
import { readBook } from '../files';
import { plural } from '../output';
import type { CommandSpec } from '../types';
import { warningLines } from './inspect';

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
    const m = /^\s*([^:]+?)\s*(?::\s*([a-z]+)(\?)?\s*)?$/.exec(part);
    const type = (m?.[2] ?? 'any') as ColumnType;
    if (!m || !(TYPES as readonly string[]).includes(type)) {
      throw usageError(`Cannot read the column rule "${part.trim()}".`, `Write rules as "name:type" separated by commas; types: ${TYPES.join(', ')}; add "?" to allow empty cells, e.g. "id:number,note:string?".`, {
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
          problems.push({ kind: 'missing-column', column: rule.name, cell: null, row: table.headerRow, message: `Column "${rule.name}" is missing. Header row: ${headers.map((h) => `"${h}"`).join(', ') || '(none)'}.` });
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
          lines.push(`${out.red('Problems')} (${String(problems.length)}):`);
          lines.push(...problems.slice(0, 100).map((p) => `  ${out.red('problem')} ${p.message}`));
          if (problems.length > 100) lines.push(out.dim(`  … ${String(problems.length - 100)} more; use --json for all`));
        }
        if (warnings.length > 0) lines.push(`Warnings (${String(warnings.length)}):`, ...warningLines(out, warnings));
        const checked = sheetName !== null ? `, ${plural(rowsChecked, 'row')} of sheet "${sheetName}" checked against ${plural(rules.length, 'column rule')}` : '';
        lines.push(ok ? `${out.green('OK')}: ${file}${checked}.` : `${out.red('Failed')}: ${file}${checked}${strict && problems.length === 0 ? ' (warnings count with --strict)' : ''}.`);
        return lines.join('\n');
      },
    };
  },
};
