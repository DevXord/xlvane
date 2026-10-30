import { inferValue } from '../../csv';
import { appendRows, editCells, type CellEdit } from '../../edit';
import { XlvError, type XlvWarning } from '../../errors';
import { readWorkbook } from '../../read';
import { pickSheet } from '../../table';
import type { CellInput } from '../../write';
import { flag, list, str, usageError } from '../args';
import { readBytes, readText, withPath, writeOutput, wroteLine } from '../files';
import { plural } from '../output';
import type { CommandSpec } from '../types';
import { warningLines } from './inspect';

/** "Sheet!A1=value", "'My sheet'!A1=value" or "A1=value". */
const SET_PATTERN = /^(?:(?:'((?:[^']|'')+)'|([^!='][^!=]*))!)?([A-Za-z]{1,3}[1-9]\d{0,6})=([\s\S]*)$/;

/**
 * The value of a --set: empty clears the cell, "=..." is a formula, "'..." is text as typed,
 * otherwise numbers, true/false and ISO 8601 dates are recognised as by `fromCSV` (D-606).
 */
export function parseSetValue(text: string): CellInput {
  if (text.startsWith('=') && text.length > 1) return { formula: text.slice(1) };
  if (text.startsWith("'")) return text.slice(1);
  return inferValue(text, Date.UTC(1900, 0, 1));
}

function parseSet(text: string, defaultSheet: string | undefined): { sheet: string | undefined; cell: string; value: CellInput } {
  const m = SET_PATTERN.exec(text);
  if (!m) {
    throw usageError(`Cannot read --set "${text}".`, 'Write it as "Sheet!A1=value" (or "A1=value" with --sheet). Quote sheet names with spaces: "\'My sheet\'!A1=5".', {
      command: 'edit',
      option: '--set',
    });
  }
  const sheet = m[1] !== undefined ? m[1].replace(/''/g, "'") : (m[2] ?? defaultSheet);
  return { sheet, cell: m[3]!.toUpperCase(), value: parseSetValue(m[4]!) };
}

async function readRows(ctx: Parameters<CommandSpec['run']>[0], path: string): Promise<unknown[]> {
  if (!path.toLowerCase().endsWith('.json')) {
    throw usageError(`"--append" expects a .json file, not "${path}".`, 'Pass a JSON array of objects (matched to the header row) or arrays (by position).', {
      command: 'edit',
      option: '--append',
    });
  }
  const text = await readText(ctx, path);
  let rows: unknown;
  try {
    rows = JSON.parse(text);
  } catch (error) {
    throw new XlvError('XLV_E_JSON_INVALID', `"${path}" is not valid JSON.`, { hint: 'Check the file with a JSON validator.', details: { path }, cause: error });
  }
  if (!Array.isArray(rows)) {
    throw new XlvError('XLV_E_JSON_INVALID', `"${path}" must hold a JSON array of rows.`, { hint: 'For example [{"name": "Tea", "total": 3}].', details: { path } });
  }
  return rows;
}

export const editCommand: CommandSpec = {
  name: 'edit',
  summary: 'Change cells or append rows, keeping everything else',
  description:
    'Changes cells and appends rows in an existing file while keeping the rest of it (styles, formulas, charts, other sheets) as it is. Values in --set are typed like CSV fields: numbers, true/false and ISO 8601 dates are recognised, "=..." is a formula, a leading apostrophe keeps the text as typed, and an empty value clears the cell. xlvane does not calculate formulas; the edited file asks spreadsheet programs to recalculate them when it is opened. The result goes to --output, or replaces the input with --force.',
  usage: '<file> (--set <Sheet!A1=value>... | --append <rows.json>) [options]',
  positionals: { min: 1, max: 1, names: ['file'] },
  writes: true,
  options: {
    set: { type: 'string', multiple: true, value: 'Sheet!A1=value', description: 'Set a cell' },
    append: { type: 'string', value: 'rows.json', description: 'Append the rows of a JSON array below the last row' },
    sheet: { type: 'string', value: 'name', description: 'Sheet for --set without a sheet name, and for --append' },
    output: { type: 'string', short: 'o', value: 'path', description: 'Write the result to a new file' },
  },
  examples: [
    { command: 'xlvane edit report.xlsx --set "Sales!B2=1250" -o updated.xlsx', description: 'Set one cell and write a new file' },
    { command: 'xlvane edit report.xlsx --sheet Sales --set B3=990 --set "C3==B3*2" --force', description: 'Set a value and a formula in place' },
    { command: 'xlvane edit report.xlsx --append rows.json --sheet Sales -o more.xlsx', description: 'Append rows; object keys are matched to the header row' },
    { command: 'xlvane edit report.xlsx --set "Sales!A2=North-East" --dry-run', description: 'Check the edit without writing' },
  ],
  related: ['create', 'diff'],
  async run(ctx) {
    const file = ctx.args.positionals[0]!;
    const sets = list(ctx.args, 'set');
    const append = str(ctx.args, 'append');
    const sheetOption = str(ctx.args, 'sheet');
    if (sets.length === 0 && append === undefined) {
      throw usageError('"xlvane edit" needs --set or --append.', 'For example: xlvane edit report.xlsx --set "Sales!B2=1250" -o updated.xlsx', { command: 'edit' });
    }
    const parsed = sets.map((s) => parseSet(s, sheetOption));
    const missing = parsed.find((p) => p.sheet === undefined);
    if (missing) {
      throw usageError(`--set "${missing.cell}=…" has no sheet name.`, 'Write "Sheet!A1=value", or give the sheet with --sheet.', { command: 'edit', option: '--set' });
    }
    const output = str(ctx.args, 'output');
    const force = flag(ctx.args, 'force');
    const dryRun = flag(ctx.args, 'dry-run');
    if (output === undefined && !force && !dryRun) {
      throw new XlvError('XLV_E_FILE_EXISTS', `"${file}" would be changed in place.`, {
        hint: 'Add --force to change the file in place, or write the result to a new file with --output <path>.',
        details: { path: file },
      });
    }
    const rows = append !== undefined ? await readRows(ctx, append) : [];
    let bytes = await readBytes(ctx, file);
    const warnings: XlvWarning[] = [];
    try {
      if (parsed.length > 0) {
        const edits: CellEdit[] = parsed.map((p) => ({ sheet: p.sheet!, cell: p.cell, value: p.value }));
        const r = editCells(bytes, edits);
        bytes = r.bytes;
        warnings.push(...r.warnings);
      }
      let appendSheet: string | null = null;
      if (append !== undefined) {
        appendSheet = sheetOption ?? pickSheet(readWorkbook(bytes), undefined).name;
        const r = appendRows(bytes, appendSheet, rows as never);
        bytes = r.bytes;
        warnings.push(...r.warnings);
      }
      // A dry run of an in-place edit reports the replacement without needing --force.
      const w = await writeOutput(ctx, output ?? file, bytes, { force: force || (dryRun && output === undefined), dryRun });
      const result = { input: file, output: w.path, cellsSet: parsed.length, rowsAppended: rows.length, appendSheet, warnings, bytes: w.bytes, written: w.written, replaced: w.replaced };
      return {
        result,
        exitCode: 0,
        render: (out) => {
          if (ctx.quiet && w.written) return '';
          const done = [parsed.length > 0 ? `set ${plural(parsed.length, 'cell')}` : '', rows.length > 0 ? `appended ${plural(rows.length, 'row')} to "${appendSheet ?? ''}"` : '']
            .filter((p) => p !== '')
            .join(', ');
          return [`${done.charAt(0).toUpperCase()}${done.slice(1)}.`, ...warningLines(out, warnings), wroteLine(w)].join('\n');
        },
      };
    } catch (error) {
      throw withPath(error, file);
    }
  },
};
