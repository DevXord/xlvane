import { extname } from 'node:path';
import { fromCSV, toCSV } from '../../csv';
import { XlvError } from '../../errors';
import { fromJSON, toJSON } from '../../table';
import { flag, headerOption, str, usageError } from '../args';
import { readBook, readText, writeOutput, wroteLine, type WriteOutcome } from '../files';
import type { CommandContext, CommandOutcome, CommandSpec } from '../types';

type Format = 'xlsx' | 'csv' | 'tsv' | 'json';

function formatOf(command: string, path: string, role: 'input' | 'output'): Format {
  const ext = extname(path).toLowerCase();
  if (ext === '.xlsx' || (role === 'input' && (ext === '.xlsm' || ext === '.xltx' || ext === '.xltm'))) return 'xlsx';
  if (ext === '.csv') return 'csv';
  if (ext === '.tsv') return 'tsv';
  if (ext === '.json') return 'json';
  const allowed = role === 'input' ? '.xlsx, .xlsm, .csv, .tsv or .json' : '.xlsx, .csv, .tsv or .json';
  throw usageError(`Cannot tell the ${role} format of "${path}" from its extension.`, `Use a file name ending in ${allowed}.`, { command, [role]: path });
}

/** The delimiter option: one character, or "tab". */
function delimiterOf(ctx: CommandContext, command: string, format: Format): string {
  const d = str(ctx.args, 'delimiter');
  if (d === undefined) return format === 'tsv' ? '\t' : ',';
  if (d === 'tab' || d === '\\t') return '\t';
  if (d.length !== 1) throw usageError(`The option "--delimiter" must be one character or "tab", not "${d}".`, 'For example --delimiter ";".', { command, option: '--delimiter' });
  return d;
}

/** Result of convert and create. */
export interface ConvertResult {
  readonly input: string;
  readonly output: string;
  readonly from: Format;
  readonly to: Format;
  readonly bytes: number;
  readonly written: boolean;
  readonly replaced: boolean;
}

/** Builds an xlsx file from CSV or JSON text (shared by convert and create). */
export async function tableToXlsx(ctx: CommandContext, command: string, input: string, from: Format, preset: 'plain' | 'report'): Promise<Uint8Array> {
  const text = await readText(ctx, input);
  const sheet = str(ctx.args, 'sheet');
  const columns = str(ctx.args, 'columns');
  const base = { preset, ...(sheet !== undefined ? { sheet } : {}) };
  if (from === 'json') {
    try {
      return fromJSON(text, { ...base, ...(columns !== undefined ? { columns: columns.split(',').map((c) => c.trim()).filter((c) => c !== '') } : {}) });
    } catch (error) {
      if (error instanceof XlvError && error.code === 'XLV_E_JSON_INVALID') {
        throw new XlvError(error.code, error.message, { hint: 'Fix the JSON file; the message gives the problem. The file must hold an array of objects or arrays.', details: { ...error.details, path: input }, cause: error });
      }
      throw error;
    }
  }
  if (columns !== undefined) throw usageError('The option "--columns" works with JSON input only.', 'For CSV input the first row names the columns.', { command, option: '--columns' });
  return fromCSV(text, { ...base, delimiter: delimiterOf(ctx, command, from) });
}

export function convertOutcome(ctx: CommandContext, r: ConvertResult, w: WriteOutcome): CommandOutcome {
  return { result: r, exitCode: 0, render: () => (ctx.quiet && w.written ? '' : wroteLine(w)) };
}

export const convertCommand: CommandSpec = {
  name: 'convert',
  summary: 'Convert between XLSX, CSV and JSON',
  description:
    'Converts a sheet to CSV, TSV or JSON (an array of row objects), or CSV, TSV and JSON to a plain XLSX file. The formats come from the file extensions. CSV output escapes text cells that start with =, +, -, @ and similar characters, so that spreadsheet programs do not run them as formulas (OWASP CSV injection); dates are written as ISO 8601. Use "xlvane create" for a formatted report.',
  usage: '<input> <output> [options]',
  positionals: { min: 2, max: 2, names: ['input', 'output'] },
  writes: true,
  options: {
    sheet: { type: 'string', value: 'name', description: 'Sheet to convert (default: the first visible sheet); for XLSX output, its name' },
    range: { type: 'string', value: 'A1:D20', description: 'Convert only this range (XLSX input)' },
    header: { type: 'string', value: 'auto|none|row', description: 'Header row for JSON output (default: detected)' },
    delimiter: { type: 'string', value: 'char', description: 'CSV delimiter (default "," and tab for .tsv)' },
    bom: { type: 'boolean', description: 'Start CSV output with a byte order mark (for Excel)' },
  },
  examples: [
    { command: 'xlvane convert report.xlsx out.csv', description: 'The first visible sheet as CSV' },
    { command: 'xlvane convert report.xlsx costs.json --sheet Costs', description: 'A sheet as an array of JSON objects' },
    { command: 'xlvane convert sales.csv sales.xlsx', description: 'CSV to XLSX (numbers, true/false and ISO 8601 dates are recognised)' },
    { command: 'xlvane convert report.xlsx out.csv --delimiter ";" --bom', description: 'CSV for spreadsheet programs that expect semicolons' },
  ],
  related: ['create', 'read'],
  async run(ctx) {
    const [input, output] = ctx.args.positionals as [string, string];
    const from = formatOf('convert', input, 'input');
    const to = formatOf('convert', output, 'output');
    if ((from === 'xlsx') === (to === 'xlsx')) {
      throw usageError(`Cannot convert ${from.toUpperCase()} to ${to.toUpperCase()}.`, 'One side must be an .xlsx file: XLSX to CSV/TSV/JSON, or CSV/TSV/JSON to XLSX.', { command: 'convert' });
    }
    let data: Uint8Array | string;
    if (from === 'xlsx') {
      const { workbook } = await readBook(ctx, input);
      const sheet = str(ctx.args, 'sheet');
      const range = str(ctx.args, 'range');
      const where = { ...(sheet !== undefined ? { sheet } : {}), ...(range !== undefined ? { range } : {}) };
      if (to === 'json') data = `${JSON.stringify(toJSON(workbook, { ...where, header: headerOption(ctx.args, 'convert') }).rows, null, 2)}\n`;
      else data = toCSV(workbook, { ...where, delimiter: delimiterOf(ctx, 'convert', to), bom: flag(ctx.args, 'bom') });
    } else {
      data = await tableToXlsx(ctx, 'convert', input, from, 'plain');
    }
    const w = await writeOutput(ctx, output, data, { force: flag(ctx.args, 'force'), dryRun: flag(ctx.args, 'dry-run') });
    return convertOutcome(ctx, { input, output, from, to, bytes: w.bytes, written: w.written, replaced: w.replaced }, w);
  },
};

export const createCommand: CommandSpec = {
  name: 'create',
  summary: 'Build a formatted XLSX report from JSON or CSV',
  description:
    'Builds a report from an array of JSON objects (or arrays) or a CSV file: bold header row, frozen first row, autofilter, column widths that fit the content, and date and number formats. --columns chooses and orders the JSON properties.',
  usage: '<data.json|data.csv> <output.xlsx> [options]',
  positionals: { min: 2, max: 2, names: ['data', 'output'] },
  writes: true,
  options: {
    sheet: { type: 'string', value: 'name', description: 'Sheet name (default "Sheet1")' },
    columns: { type: 'string', value: 'a,b,c', description: 'JSON properties to include, in this order' },
    delimiter: { type: 'string', value: 'char', description: 'CSV delimiter (default "," and tab for .tsv)' },
  },
  examples: [
    { command: 'xlvane create data.json new-report.xlsx', description: 'A report with one column per property' },
    { command: 'xlvane create data.json people.xlsx --sheet People --columns name,city', description: 'Choose the sheet name and the columns' },
    { command: 'xlvane create sales.csv sales-report.xlsx --dry-run', description: 'Check the input without writing a file' },
  ],
  related: ['convert', 'edit'],
  async run(ctx) {
    const [input, output] = ctx.args.positionals as [string, string];
    const from = formatOf('create', input, 'input');
    if (from === 'xlsx') throw usageError('"xlvane create" builds a report from JSON or CSV data.', 'Pass a .json, .csv or .tsv file as the input.', { command: 'create', input });
    const to = formatOf('create', output, 'output');
    if (to !== 'xlsx') throw usageError('"xlvane create" writes an .xlsx file.', 'Give an output name ending in .xlsx.', { command: 'create', output });
    const data = await tableToXlsx(ctx, 'create', input, from, 'report');
    const w = await writeOutput(ctx, output, data, { force: flag(ctx.args, 'force'), dryRun: flag(ctx.args, 'dry-run') });
    return convertOutcome(ctx, { input, output, from, to, bytes: w.bytes, written: w.written, replaced: w.replaced }, w);
  },
};
