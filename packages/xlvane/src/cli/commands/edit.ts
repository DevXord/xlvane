import { inferValue } from '../../csv';
import { appendRows, editCells, type CellEdit } from '../../edit';
import { deleteRows, insertRows } from '../../rows';
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

/** A row number of --insert-rows / --delete-rows: 1 to 1048576. */
function rowNumber(text: string, option: string): number {
  const n = /^[1-9]\d{0,6}$/.test(text.trim()) ? Number(text.trim()) : Number.NaN;
  if (!Number.isInteger(n) || n > 1_048_576) {
    throw usageError(`${option} expects a row number from 1 to 1048576, not "${text}".`, `For example ${option} 5.`, { command: 'edit', option });
  }
  return n;
}

/** The rows of --delete-rows: "5" or "5:7". */
function rowSpan(text: string): { at: number; count: number } {
  const [a, b] = text.split(':');
  const at = rowNumber(a ?? '', '--delete-rows');
  const end = b === undefined ? at : rowNumber(b, '--delete-rows');
  if (end < at) throw usageError(`--delete-rows "${text}" ends before it starts.`, 'Write the first row first, for example --delete-rows 5:7.', { command: 'edit', option: '--delete-rows' });
  return { at, count: end - at + 1 };
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
  summary: 'Change cells, insert, delete or append rows, keeping everything else',
  description:
    'Changes cells, and inserts, deletes or appends rows, in an existing file while keeping the rest of it (styles, formulas, charts, other sheets) as it is. Values in --set are typed like CSV fields: numbers, true/false and ISO 8601 dates are recognised, "=..." is a formula, a leading apostrophe keeps the text as typed, and an empty value clears the cell. xlvane does not calculate formulas; the edited file asks spreadsheet programs to recalculate them when it is opened. Rows inserted with --insert-rows or deleted with --delete-rows move the rows below, and formulas on every sheet, merged cells, hyperlinks, notes, conditional formats, data validation and defined names follow them; sheets with tables, pictures or charts are refused. Row changes are applied first, then --set, then --append, so cell addresses in --set refer to the rows after the change. The result goes to --output, or replaces the input with --force.',
  usage: '<file> (--set <Sheet!A1=value>... | --append <rows.json> | --insert-rows <row> | --delete-rows <row[:row]>) [options]',
  positionals: { min: 1, max: 1, names: ['file'] },
  writes: true,
  options: {
    set: { type: 'string', multiple: true, value: 'Sheet!A1=value', description: 'Set a cell' },
    append: { type: 'string', value: 'rows.json', description: 'Append the rows of a JSON array below the last row' },
    'insert-rows': { type: 'string', value: 'row', description: 'Insert empty rows before this row (needs --sheet)' },
    count: { type: 'string', value: 'n', description: 'How many rows --insert-rows inserts (default 1)' },
    'insert-style': { type: 'string', value: 'above|below|none', description: 'Format new rows like the row above or below (default none)' },
    'delete-rows': { type: 'string', value: 'row[:row]', description: 'Delete these rows (needs --sheet)' },
    sheet: { type: 'string', value: 'name', description: 'Sheet for --set without a sheet name, --append and the row options' },
    output: { type: 'string', short: 'o', value: 'path', description: 'Write the result to a new file' },
  },
  examples: [
    { command: 'xlvane edit report.xlsx --set "Sales!B2=1250" -o updated.xlsx', description: 'Set one cell and write a new file' },
    { command: 'xlvane edit report.xlsx --append rows.json --sheet Sales -o more.xlsx', description: 'Append rows; object keys are matched to the header row' },
    { command: 'xlvane edit report.xlsx --sheet Sales --insert-rows 5 --count 2 --insert-style above -o more.xlsx', description: 'Insert two rows before row 5, formatted like row 4' },
    { command: 'xlvane edit report.xlsx --sheet Sales --delete-rows 5:6 --force', description: 'Delete rows 5 and 6 in place' },
  ],
  related: ['create', 'diff'],
  async run(ctx) {
    const file = ctx.args.positionals[0]!;
    const sets = list(ctx.args, 'set');
    const append = str(ctx.args, 'append');
    const sheetOption = str(ctx.args, 'sheet');
    const insertAt = str(ctx.args, 'insert-rows');
    const deleteSpan = str(ctx.args, 'delete-rows');
    if (sets.length === 0 && append === undefined && insertAt === undefined && deleteSpan === undefined) {
      throw usageError('"xlvane edit" needs --set, --append, --insert-rows or --delete-rows.', 'For example: xlvane edit report.xlsx --set "Sales!B2=1250" -o updated.xlsx', { command: 'edit' });
    }
    const countText = str(ctx.args, 'count');
    const styleText = str(ctx.args, 'insert-style');
    if ((countText !== undefined || styleText !== undefined) && insertAt === undefined) {
      throw usageError(`${countText !== undefined ? '--count' : '--insert-style'} works only with --insert-rows.`, 'For example: --insert-rows 5 --count 2. To delete several rows, write --delete-rows 5:6.', { command: 'edit' });
    }
    const insert = insertAt === undefined ? null : { at: rowNumber(insertAt, '--insert-rows'), count: countText === undefined ? 1 : rowNumber(countText, '--count') };
    const remove = deleteSpan === undefined ? null : rowSpan(deleteSpan);
    const style = styleText ?? 'none';
    if (style !== 'above' && style !== 'below' && style !== 'none') {
      throw usageError(`--insert-style must be above, below or none, not "${style}".`, 'For example --insert-style above.', { command: 'edit', option: '--insert-style' });
    }
    if ((insert !== null || remove !== null) && sheetOption === undefined) {
      throw usageError(`${insert !== null ? '--insert-rows' : '--delete-rows'} needs --sheet.`, 'Name the sheet whose rows change, for example --sheet Sales.', { command: 'edit', option: '--sheet' });
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
      if (insert !== null) {
        const r = insertRows(bytes, { sheet: sheetOption!, at: insert.at, count: insert.count, style });
        bytes = r.bytes;
        warnings.push(...r.warnings);
      }
      if (remove !== null) {
        const r = deleteRows(bytes, { sheet: sheetOption!, at: remove.at, count: remove.count });
        bytes = r.bytes;
        warnings.push(...r.warnings);
      }
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
      const unique = [...new Map(warnings.map((x) => [`${x.code}|${x.message}`, x])).values()];
      const result = {
        input: file,
        output: w.path,
        rowsInserted: insert?.count ?? 0,
        rowsDeleted: remove?.count ?? 0,
        rowSheet: insert !== null || remove !== null ? (sheetOption ?? null) : null,
        cellsSet: parsed.length,
        rowsAppended: rows.length,
        appendSheet,
        warnings: unique,
        bytes: w.bytes,
        written: w.written,
        replaced: w.replaced,
      };
      return {
        result,
        exitCode: 0,
        render: (out) => {
          if (ctx.quiet && w.written) return '';
          const done = [
            insert !== null ? `inserted ${plural(insert.count, 'row')} before row ${String(insert.at)} of "${sheetOption ?? ''}"` : '',
            remove !== null ? `deleted ${remove.count === 1 ? `row ${String(remove.at)}` : `rows ${String(remove.at)}–${String(remove.at + remove.count - 1)}`} of "${sheetOption ?? ''}"` : '',
            parsed.length > 0 ? `set ${plural(parsed.length, 'cell')}` : '',
            rows.length > 0 ? `appended ${plural(rows.length, 'row')} to "${appendSheet ?? ''}"` : '',
          ]
            .filter((p) => p !== '')
            .join(', ');
          return [`${done.charAt(0).toUpperCase()}${done.slice(1)}.`, ...warningLines(out, unique), wroteLine(w)].join('\n');
        },
      };
    } catch (error) {
      throw withPath(error, file);
    }
  },
};
