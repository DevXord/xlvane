import { readTable } from '../../table';
import { headerOption, positiveInt, str } from '../args';
import { readBook } from '../files';
import { plural, showValue } from '../output';
import type { CommandSpec } from '../types';

/** Rows shown in the terminal unless --limit is given. */
const DEFAULT_LIMIT = 50;

export const readCommand: CommandSpec = {
  name: 'read',
  summary: 'Print a sheet as a table',
  description:
    'Prints the rows of a sheet. The header row is detected unless --header gives its number or "none" (then columns are named by their letters). Long cells are cut to fit the terminal; --json prints every row in full, with dates as ISO 8601 strings.',
  usage: '<file> [options]',
  positionals: { min: 1, max: 1, names: ['file'] },
  options: {
    sheet: { type: 'string', value: 'name', description: 'Sheet to read (default: the first visible sheet)' },
    range: { type: 'string', value: 'A1:D20', description: 'Read only this range' },
    header: { type: 'string', value: 'auto|none|row', description: 'Header row: detect it (default), none, or its number' },
    limit: { type: 'string', value: 'n', description: `Rows to print (default ${String(DEFAULT_LIMIT)}; with --json: all)` },
  },
  examples: [
    { command: 'xlvane read report.xlsx', description: 'Print the first visible sheet' },
    { command: 'xlvane read report.xlsx --sheet Sales --limit 10', description: 'The first 10 rows of the sheet "Sales"' },
    { command: 'xlvane read report.xlsx --range A1:B3 --header none', description: 'A range, with columns named A, B' },
    { command: 'xlvane read report.xlsx --json', description: 'All rows as JSON objects' },
  ],
  related: ['inspect', 'convert'],
  async run(ctx) {
    const file = ctx.args.positionals[0]!;
    const header = headerOption(ctx.args, 'read');
    const limitOption = positiveInt(ctx.args, 'read', 'limit');
    const { workbook } = await readBook(ctx, file);
    const sheet = str(ctx.args, 'sheet');
    const range = str(ctx.args, 'range');
    const table = readTable(workbook, { ...(sheet !== undefined ? { sheet } : {}), ...(range !== undefined ? { range } : {}), header });
    const limit = limitOption ?? (ctx.json ? table.rows.length : DEFAULT_LIMIT);
    const shown = table.rows.slice(0, limit);
    const result = {
      sheet: table.sheet,
      headerRow: table.headerRow,
      range: table.range,
      columns: table.columns.map((c) => ({ key: c.key, header: c.header, letter: c.letter })),
      // Dates as ISO 8601 strings, so the JSON is lossless (docs/cli-json.md).
      rows: shown.map((row) => Object.fromEntries(Object.entries(row).map(([k, v]) => [k, v instanceof Date ? v.toISOString() : v]))),
      rowNumbers: table.rowNumbers.slice(0, limit),
      totalRows: table.rows.length,
      truncated: shown.length < table.rows.length,
    };
    return {
      result,
      exitCode: 0,
      render: (out) => {
        if (table.columns.length === 0) return `Sheet "${table.sheet}" has no data${table.range ? ` in ${table.range}` : ''}.`;
        const lines = [
          out.table(
            table.columns.map((c) => c.header ?? c.key),
            shown.map((row) => table.columns.map((c) => showValue(row[c.key]))),
          ),
        ];
        const about = `Sheet "${table.sheet}"${table.range ? `, ${table.range}` : ''}${table.headerRow ? `, header in row ${String(table.headerRow)}` : ''}`;
        if (result.truncated) lines.push(out.dim(`${about}: showing ${String(shown.length)} of ${plural(table.rows.length, 'row')}. Use --limit to see more, or --json for all rows.`));
        else if (!ctx.quiet) lines.push(out.dim(`${about}: ${plural(table.rows.length, 'row')}.`));
        return lines.join('\n');
      },
    };
  },
};
