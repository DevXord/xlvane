import { parseRange } from '../../sml/cell-ref';
import { readTable } from '../../table';
import { isoDate } from '../../util/iso-date';
import { headerOption, str, usageError } from '../args';
import { readBook } from '../files';
import { plural, showValue } from '../output';
import type { CommandContext, CommandSpec } from '../types';
import { warningNote } from '../warnings';

/** Rows printed (and in --json) unless --limit is given. */
const DEFAULT_LIMIT = 50;

/** --limit: a whole number of 1 or more, or "all". */
export function limitOption(ctx: CommandContext, command: string): number | undefined {
  const v = str(ctx.args, 'limit');
  if (v === undefined) return undefined;
  if (v === 'all') return Number.POSITIVE_INFINITY;
  if (!/^\d+$/.test(v) || Number(v) < 1 || !Number.isSafeInteger(Number(v))) {
    throw usageError(`The option "--limit" must be a whole number of 1 or more, or "all", not "${v}".`, 'For example --limit 20, or --limit all for every row.', { command, option: '--limit' });
  }
  return Number(v);
}

/** --range: checked before the file is read, so a typo is a usage error (exit code 2). */
export function rangeOption(ctx: CommandContext, command: string): string | undefined {
  const v = str(ctx.args, 'range');
  if (v !== undefined && parseRange(v) === null) {
    throw usageError(`The option "--range" must be a range such as A1:D20, not "${v}".`, 'Write the top-left and bottom-right cells separated by a colon, for example --range A1:D20.', { command, option: '--range' });
  }
  return v;
}

export const readCommand: CommandSpec = {
  name: 'read',
  summary: 'Print a sheet as a table',
  description:
    'Prints the rows of a sheet. The header row is detected unless --header gives its number or "none" (then columns are named by their letters). Long cells are cut to fit the terminal, numbers are shown to 15 significant digits and dates without a time zone. --json gives the values in full, with dates as ISO 8601 text without a time zone. The first 50 rows are shown unless --limit says otherwise; "--limit all" shows every row. Warnings found while reading are noted on stderr (in --json, under "warnings").',
  usage: '<file> [options]',
  positionals: { min: 1, max: 1, names: ['file'] },
  options: {
    sheet: { type: 'string', value: 'name', description: 'Sheet to read (default: the first visible sheet)' },
    range: { type: 'string', value: 'A1:D20', description: 'Read only this range' },
    header: { type: 'string', value: 'auto|none|row', description: 'Header row: detect it (default), none, or its number' },
    limit: { type: 'string', value: 'n|all', description: `Rows to show (default ${String(DEFAULT_LIMIT)}, also with --json; "all" for every row)` },
  },
  examples: [
    { command: 'xlvane read report.xlsx', description: 'Print the first visible sheet' },
    { command: 'xlvane read report.xlsx --sheet Sales --limit 10', description: 'The first 10 rows of the sheet "Sales"' },
    { command: 'xlvane read report.xlsx --range A1:B3 --header none', description: 'A range, with columns named A, B' },
    { command: 'xlvane read report.xlsx --json --limit all', description: 'All rows as JSON objects' },
  ],
  related: ['inspect', 'convert'],
  async run(ctx) {
    const file = ctx.args.positionals[0]!;
    const header = headerOption(ctx.args, 'read');
    const limit = limitOption(ctx, 'read') ?? DEFAULT_LIMIT;
    const range = rangeOption(ctx, 'read');
    const { workbook } = await readBook(ctx, file);
    const sheet = str(ctx.args, 'sheet');
    const table = readTable(workbook, { ...(sheet !== undefined ? { sheet } : {}), ...(range !== undefined ? { range } : {}), header });
    const shown = table.rows.slice(0, limit);
    const warnings = workbook.warnings;
    const result = {
      sheet: table.sheet,
      headerRow: table.headerRow,
      range: table.range,
      columns: table.columns.map((c) => ({ key: c.key, header: c.header, letter: c.letter })),
      // Dates as ISO 8601 text without a time zone, as in CSV (docs/cli-json.md).
      rows: shown.map((row) => Object.fromEntries(Object.entries(row).map(([k, v]) => [k, v instanceof Date ? isoDate(v) : v]))),
      rowNumbers: table.rowNumbers.slice(0, limit),
      totalRows: table.rows.length,
      truncated: shown.length < table.rows.length,
      warnings,
    };
    const note = warningNote(warnings, file);
    if (note !== null) ctx.note(note);
    const about = `Sheet "${table.sheet}"${table.range ? `, ${table.range}` : ''}${table.headerRow ? `, header in row ${String(table.headerRow)}` : ''}`;
    // The summary goes to stderr, so that "xlvane read file > out.txt" holds only the table.
    if (table.columns.length > 0) {
      if (result.truncated) ctx.note(`${about}: showing ${String(shown.length)} of ${plural(table.rows.length, 'row')}. Use --limit all to see every row.`);
      else ctx.note(`${about}: ${plural(table.rows.length, 'row')}.`);
    }
    return {
      result,
      exitCode: 0,
      render: (out) => {
        if (table.columns.length === 0) return `Sheet "${table.sheet}" has no data${table.range ? ` in ${table.range}` : ''}.`;
        const numeric = table.columns.map((c) => (shown.length > 0 && shown.every((row) => row[c.key] === null || typeof row[c.key] === 'number') ? 'right' : 'left'));
        return out.table(
          table.columns.map((c) => c.header ?? c.key),
          shown.map((row) => table.columns.map((c) => showValue(row[c.key]))),
          '',
          numeric,
        );
      },
    };
  },
};
