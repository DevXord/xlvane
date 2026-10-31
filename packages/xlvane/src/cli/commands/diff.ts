import { diff, type DiffCell } from '../../diff';
import { flag, positiveInt } from '../args';
import { readBook, withPath } from '../files';
import { plural, type Output } from '../output';
import type { CommandSpec } from '../types';

/** Changes listed per sheet unless --limit is given (all of them are in --json). */
const DEFAULT_LIMIT = 50;

function showCell(c: DiffCell | null): string {
  if (c === null) return '(empty)';
  const value = c.value === null ? '' : typeof c.value === 'string' ? JSON.stringify(c.value) : String(c.value);
  return c.formula !== null ? `=${c.formula}${value !== '' ? ` (${value})` : ''}` : value === '' ? '(no value)' : value;
}

function kindLabel(out: Output, kind: 'added' | 'removed' | 'changed'): string {
  const text = kind.padEnd(7);
  return kind === 'added' ? out.green(text) : kind === 'removed' ? out.red(text) : out.yellow(text);
}

export const diffCommand: CommandSpec = {
  name: 'diff',
  summary: 'Show what changed between two workbooks',
  description:
    'Compares two workbooks cell by cell: values, formulas and, with --styles, formatting. Sheets are matched by name. The exit code is 1 when the workbooks differ, as with the diff tool, so scripts can check the result.',
  usage: '<old.xlsx> <new.xlsx> [options]',
  positionals: { min: 2, max: 2, names: ['old', 'new'] },
  checks: 'the workbooks differ',
  options: {
    styles: { type: 'boolean', description: 'Also compare cell formatting' },
    limit: { type: 'string', value: 'n', description: `Changes listed per sheet (default ${String(DEFAULT_LIMIT)}; --json lists all)` },
  },
  examples: [
    { command: 'xlvane diff old.xlsx new.xlsx', description: 'List changed, added and removed cells' },
    { command: 'xlvane diff old.xlsx new.xlsx --styles', description: 'Include formatting changes' },
    { command: 'xlvane diff old.xlsx new.xlsx --json', description: 'The changes as JSON' },
  ],
  related: ['edit', 'inspect'],
  async run(ctx) {
    const [a, b] = ctx.args.positionals as [string, string];
    const limit = positiveInt(ctx.args, 'diff', 'limit') ?? DEFAULT_LIMIT;
    const before = (await readBook(ctx, a)).workbook;
    const after = (await readBook(ctx, b)).workbook;
    let result;
    try {
      result = diff(before, after, { styles: flag(ctx.args, 'styles') });
    } catch (error) {
      throw withPath(error, `${a}, ${b}`);
    }
    const r = result;
    return {
      result: r,
      exitCode: r.equal ? 0 : 1,
      render: (out) => {
        if (r.equal) return `${out.green('No differences')} between ${a} and ${b}.`;
        const s = r.summary;
        const lines = [
          `${out.bold(`${a} → ${b}`)}: ${plural(s.cellsChanged, 'cell')} changed, ${String(s.cellsAdded)} added, ${String(s.cellsRemoved)} removed` +
            (s.sheetsAdded + s.sheetsRemoved > 0 ? `; sheets added: ${String(s.sheetsAdded)}, removed: ${String(s.sheetsRemoved)}` : ''),
        ];
        for (const name of r.sheetsAdded) lines.push(`${out.green('added sheet')}   "${name}"`);
        for (const name of r.sheetsRemoved) lines.push(`${out.red('removed sheet')} "${name}"`);
        for (const sheet of r.sheets) {
          if (sheet.cells.length === 0 && !sheet.state) continue;
          lines.push('', out.bold(`Sheet "${sheet.name}"`));
          if (sheet.state) lines.push(`  ${out.yellow('changed')} visibility: ${sheet.state.before} → ${sheet.state.after}`);
          for (const c of sheet.cells.slice(0, limit)) {
            const style = c.kind === 'changed' && showCell(c.before) === showCell(c.after) ? ' (formatting)' : '';
            const detail = c.kind === 'added' ? showCell(c.after) : c.kind === 'removed' ? showCell(c.before) : `${showCell(c.before)} → ${showCell(c.after)}${style}`;
            lines.push(`  ${kindLabel(out, c.kind)} ${c.ref.padEnd(6)} ${detail}`);
          }
          if (sheet.cells.length > limit) lines.push(out.dim(`  … ${String(sheet.cells.length - limit)} more; use --limit or --json`));
          if (sheet.truncated) lines.push(out.dim('  The comparison stopped at the cell limit; not every change is listed.'));
        }
        return lines.join('\n');
      },
    };
  },
};
