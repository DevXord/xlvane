// The tools of the MCP server (docs/PLAN.md section 4; docs/decisions.md D-702). Each tool has a
// JSON Schema for its input (checked by the SDK through fromJsonSchema), returns its result as
// structuredContent and as JSON text, and reports XlvErrors as tool errors with code and hint.
import { fromJsonSchema, type CallToolResult, type McpServer, type ToolAnnotations } from '@modelcontextprotocol/server';
import {
  appendRows,
  diff,
  editCells,
  ERROR_VALUES,
  fromCSV,
  fromJSON,
  inspect,
  isXlvError,
  readTable,
  readWorkbook,
  toCSV,
  toJSON,
  writeTable,
  XlvError,
  type CellEdit,
  type CellInput,
  type TableColumnSpec,
  type Workbook,
  type XlvWarning,
} from 'xlvane';
import { readAllowed, writeAllowed, type Roots } from './paths';

/** Limits of the tools; set when the server starts. */
export interface ToolLimits {
  /** Largest number of rows `read_table` returns in one call. */
  readonly maxRows: number;
  /** Largest result, in characters of its JSON text (D-I02). */
  readonly maxResponseChars: number;
  /** The `maxCells` read limit of every workbook read (D-J01). */
  readonly maxCells: number;
}

/** Problems `validate` lists; `problemsTotal` counts all of them (D-I02). */
const MAX_PROBLEMS = 1000;

type Json = Record<string, unknown>;
type Args = Record<string, unknown>;

const READ: ToolAnnotations = { readOnlyHint: true, openWorldHint: false };
// Write tools never delete; they replace a file only when the call passes "overwrite": true.
const WRITE: ToolAnnotations = { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false };

const PATH = { type: 'string', minLength: 1, description: 'File path, relative to the first allowed directory or absolute inside an allowed directory.' };
const SHEET = { type: 'string', description: 'Sheet name. Default: the first visible worksheet.' };
const OVERWRITE = { type: 'boolean', description: 'Replace the output file if it exists. Default false: an existing file is never replaced.' };
const HEADER = {
  description: 'Header row: "auto" (default) detects it, a 1-based row number, or false for no header (columns are named by letter).',
  anyOf: [{ const: 'auto' }, { type: 'integer', minimum: 1 }, { const: false }],
};
const CELL_VALUE = {
  description: 'A number, string, boolean or null (clears the cell); {"formula": "SUM(A1:A3)", "result": 6} for a formula (without "="; xlvane does not calculate it); {"error": "#N/A"}; {"date": "2024-01-31"} for a date (ISO 8601, no time zone).',
  anyOf: [
    { type: ['number', 'string', 'boolean', 'null'] },
    { type: 'object', properties: { formula: { type: 'string', minLength: 1 }, result: { type: ['number', 'string', 'boolean', 'null'] } }, required: ['formula'], additionalProperties: false },
    { type: 'object', properties: { error: { enum: [...ERROR_VALUES] } }, required: ['error'], additionalProperties: false },
    { type: 'object', properties: { date: { type: 'string', pattern: '^\\d{4}-\\d{2}-\\d{2}([T ]\\d{2}:\\d{2}(:\\d{2}(\\.\\d{1,3})?)?)?$' } }, required: ['date'], additionalProperties: false },
  ],
};

function object(properties: Json, required: readonly string[] = []): ReturnType<typeof fromJsonSchema> {
  return fromJsonSchema({ type: 'object', properties, required: [...required], additionalProperties: false } as never);
}

function ok(result: Json, maxChars: number): CallToolResult {
  const text = JSON.stringify(result);
  // Text from a file goes into the client's context, so a result has a size limit (D-I02).
  if (text.length > maxChars) {
    return fail(
      new XlvError('XLV_E_LIMIT_EXCEEDED', `The result is ${String(text.length)} characters of JSON, more than the server's limit of ${String(maxChars)}.`, {
        hint: 'Ask for less: one sheet, a smaller range, fewer changes (maxChanges) or fewer rows; the limit is set when the server starts (--max-response-chars).',
        details: { limit: 'maxResponseChars', value: maxChars },
      }),
    );
  }
  return { content: [{ type: 'text', text }], structuredContent: result };
}

function fail(error: unknown, maxChars: number = Number.POSITIVE_INFINITY): CallToolResult {
  const e = isXlvError(error)
    ? error
    : new XlvError('XLV_E_INTERNAL', `Unexpected error: ${error instanceof Error ? error.message : String(error)}`, {
        hint: 'This is a bug in xlvane; the call may succeed with other input.',
        details: { name: error instanceof Error ? error.name : typeof error },
      });
  let result: Json = { error: e.toJSON() };
  let text = JSON.stringify(result);
  // An error can echo long arguments (a path); it is limited like any result (D-I04).
  if (text.length > maxChars) {
    const json = e.toJSON();
    result = { error: { ...json, message: `${json.message.slice(0, Math.max(0, Math.floor(maxChars / 4)))}…`, hint: json.hint.slice(0, Math.max(0, Math.floor(maxChars / 4))), details: { cut: true } } };
    text = JSON.stringify(result);
    if (text.length > maxChars) {
      result = { error: { name: json.name, code: json.code, message: 'The error message is longer than the result limit.', location: {}, hint: '', docsUrl: json.docsUrl, details: { cut: true } } };
      text = JSON.stringify(result);
    }
  }
  return { content: [{ type: 'text', text }], structuredContent: result, isError: true };
}

function handler(maxChars: number, run: (args: Args) => Promise<Json>): (args: unknown) => Promise<CallToolResult> {
  return async (args) => {
    try {
      return ok(await run((args ?? {}) as Args), maxChars);
    } catch (error) {
      return fail(error, maxChars);
    }
  };
}

async function book(roots: Roots, limits: ToolLimits, path: string): Promise<Workbook> {
  const bytes = await readAllowed(roots, path);
  try {
    return readWorkbook(bytes, { limits: { maxCells: limits.maxCells } });
  } catch (error) {
    throw withPath(error, path);
  }
}

/**
 * `JSON.stringify(rows, null, 2)` plus a line feed, built row by row: every row repeats the header
 * names, so a small file can make a text far larger than the file limit, or than a string can be.
 * The text stops as soon as it is longer than `maxChars` characters, before encoding (D-J01).
 */
function jsonText(rows: readonly unknown[], maxChars: number, path: string): string {
  if (rows.length === 0) return '[]\n';
  const parts = ['['];
  let length = 1;
  for (let i = 0; i < rows.length; i++) {
    const part = `${i === 0 ? '' : ','}\n  ${JSON.stringify(rows[i], null, 2).replace(/\n/g, '\n  ')}`;
    length += part.length;
    if (length > maxChars) {
      throw new XlvError('XLV_E_LIMIT_EXCEEDED', `The JSON text would be longer than the server's limit of ${String(maxChars)} bytes.`, {
        hint: 'Convert fewer rows or columns (sheet with a smaller table), or raise --max-file-size.',
        details: { path, limit: 'maxFileBytes', value: maxChars },
      });
    }
    parts.push(part);
  }
  parts.push('\n]\n');
  return parts.join('');
}

function withPath(error: unknown, path: string): unknown {
  if (error instanceof XlvError && error.details['path'] === undefined) {
    return new XlvError(error.code, error.message, { location: error.location, hint: error.hint, details: { ...error.details, path }, cause: error });
  }
  return error;
}

function isoValues(row: Record<string, unknown>): Record<string, unknown> {
  return Object.fromEntries(Object.entries(row).map(([k, v]) => [k, v instanceof Date ? v.toISOString() : v]));
}

/** Converts a JSON cell value of the tools to a CellInput ({ date } becomes a Date, read as UTC). */
export function cellInput(value: unknown): CellInput {
  if (value !== null && typeof value === 'object' && 'date' in value) {
    const text = String((value as { date: unknown }).date).replace(' ', 'T');
    const invalid = (): XlvError => new XlvError('XLV_E_INVALID_VALUE', `"${text.slice(0, 40)}" is not a valid date.`, { hint: 'Use ISO 8601, for example {"date": "2024-01-31"} or {"date": "2024-01-31T13:45:00"}.' });
    const m = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?)?$/.exec(text);
    if (m === null) throw invalid();
    const d = new Date(/T/.test(text) ? `${text}Z` : `${text}T00:00:00Z`);
    // Date() moves impossible values (February 30, hour 25) to another day; they are refused (D-I04).
    const parts = [d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate(), d.getUTCHours(), d.getUTCMinutes(), d.getUTCSeconds()];
    const wanted = [m[1], m[2], m[3], m[4] ?? '0', m[5] ?? '0', m[6] ?? '0'].map(Number);
    if (Number.isNaN(d.getTime()) || parts.some((v, i) => v !== wanted[i])) throw invalid();
    return d;
  }
  return value as CellInput;
}

function extension(path: string): string {
  const m = /\.([A-Za-z0-9]+)$/.exec(path);
  return m ? m[1]!.toLowerCase() : '';
}

const TYPES = ['number', 'string', 'date', 'boolean', 'any'] as const;

function typeOf(v: unknown): string {
  if (v === null || v === undefined || v === '') return 'empty';
  if (typeof v === 'number') return 'number';
  if (typeof v === 'boolean') return 'boolean';
  if (v instanceof Date) return 'date';
  return 'string';
}

/** Registers the read tools, and the write tools when writing is allowed, in a fixed order. */
export function registerTools(server: McpServer, roots: Roots, allowWrite: boolean, limits: ToolLimits): void {
  server.registerTool(
    'inspect_workbook',
    {
      title: 'Inspect a workbook',
      description:
        'Describes an XLSX file without changing it: sheets (including hidden ones), used ranges, detected header rows, column types with sample values, merged cells, example formulas, and warnings (macros, external links, damaged parts). Macros are never run and links are never followed. Use it before read_table to find the sheet and header row.',
      inputSchema: object({ path: PATH, sheet: { ...SHEET, description: 'Describe only this sheet.' } }, ['path']),
      annotations: READ,
    },
    handler(limits.maxResponseChars, async (a) => {
      const workbook = await book(roots, limits, a['path'] as string);
      const report = inspect(workbook);
      if (typeof a['sheet'] !== 'string') return report as unknown as Json;
      const sheet = workbook.getSheet(a['sheet']);
      return { ...report, sheets: report.sheets.filter((s) => s.name === sheet.name), warnings: report.warnings.filter((w) => w.location.sheet === undefined || w.location.sheet === sheet.name) };
    }),
  );

  server.registerTool(
    'read_table',
    {
      title: 'Read a sheet as rows',
      description: `Reads a sheet as an array of row objects keyed by the header row (or by column letters with header: false). Dates are returned as ISO 8601 strings in UTC, error cells as their code (e.g. "#N/A"), formulas as their cached result. Returns at most ${String(limits.maxRows)} rows per call; use offset to page through larger sheets (totalRows and truncated say whether more rows exist).`,
      inputSchema: object(
        {
          path: PATH,
          sheet: SHEET,
          range: { type: 'string', description: 'Only this A1 range, e.g. "A1:D100".' },
          header: HEADER,
          offset: { type: 'integer', minimum: 0, description: 'Rows to skip. Default 0.' },
          limit: { type: 'integer', minimum: 1, maximum: limits.maxRows, description: `Rows to return. Default ${String(Math.min(200, limits.maxRows))}.` },
        },
        ['path'],
      ),
      annotations: READ,
    },
    handler(limits.maxResponseChars, async (a) => {
      const workbook = await book(roots, limits, a['path'] as string);
      const table = readTable(workbook, {
        ...(typeof a['sheet'] === 'string' ? { sheet: a['sheet'] } : {}),
        ...(typeof a['range'] === 'string' ? { range: a['range'] } : {}),
        header: (a['header'] ?? 'auto') as 'auto' | number | false,
      });
      const offset = (a['offset'] as number | undefined) ?? 0;
      const limit = (a['limit'] as number | undefined) ?? Math.min(200, limits.maxRows);
      const head = {
        sheet: table.sheet,
        headerRow: table.headerRow,
        range: table.range,
        columns: table.columns.map((c) => ({ key: c.key, header: c.header, letter: c.letter })),
      };
      // The rows that fit in the result size limit (D-I02); the rest are read with nextOffset.
      // The JSON text of the result is the parts below plus each row and its number.
      let used = JSON.stringify({ ...head, rows: [], rowNumbers: [], offset, totalRows: table.rows.length, truncated: true, nextOffset: table.rows.length }).length;
      const rows: Record<string, unknown>[] = [];
      for (let i = offset; i < Math.min(table.rows.length, offset + limit); i++) {
        const row = isoValues(table.rows[i]!);
        const cost = JSON.stringify(row).length + String(table.rowNumbers[i]!).length + 2;
        if (used + cost > limits.maxResponseChars) break;
        used += cost;
        rows.push(row);
      }
      if (rows.length === 0 && offset < table.rows.length && limit > 0) {
        throw new XlvError('XLV_E_LIMIT_EXCEEDED', `Row ${String(table.rowNumbers[offset]!)} alone is larger than the server's result limit of ${String(limits.maxResponseChars)} characters.`, {
          hint: 'Read fewer columns with range (for example "A1:D100"); the limit is set when the server starts (--max-response-chars).',
          details: { limit: 'maxResponseChars', value: limits.maxResponseChars },
        });
      }
      const end = offset + rows.length;
      return {
        ...head,
        rows,
        rowNumbers: table.rowNumbers.slice(offset, end),
        offset,
        totalRows: table.rows.length,
        truncated: end < table.rows.length,
        ...(end < table.rows.length ? { nextOffset: end } : {}),
      };
    }),
  );

  server.registerTool(
    'diff',
    {
      title: 'Compare two workbooks',
      description: 'Compares two XLSX files cell by cell (values and formulas; formatting with styles: true). Sheets are matched by name. Use it to check the result of an edit.',
      inputSchema: object(
        {
          before: PATH,
          after: PATH,
          styles: { type: 'boolean', description: 'Also compare cell formatting. Default false.' },
          maxChanges: { type: 'integer', minimum: 1, maximum: 10000, description: 'Changes listed per sheet. Default 200; the summary always counts all of them.' },
        },
        ['before', 'after'],
      ),
      annotations: READ,
    },
    handler(limits.maxResponseChars, async (a) => {
      const before = await book(roots, limits, a['before'] as string);
      const after = await book(roots, limits, a['after'] as string);
      const r = diff(before, after, { styles: a['styles'] === true });
      const max = (a['maxChanges'] as number | undefined) ?? 200;
      return { ...r, sheets: r.sheets.map((s) => ({ ...s, cells: s.cells.slice(0, max), listed: Math.min(max, s.cells.length), total: s.cells.length })) };
    }),
  );

  server.registerTool(
    'validate',
    {
      title: 'Check a workbook and its columns',
      description:
        'Reports read warnings of an XLSX file and, with columns, checks a table: each listed column must exist in the header row, its cells must not be empty unless optional is true, and must have the given type. Each problem names the cell. ok is false when problems were found (warnings count only with strict: true).',
      inputSchema: object(
        {
          path: PATH,
          sheet: SHEET,
          columns: {
            type: 'array',
            maxItems: 256,
            items: { type: 'object', properties: { name: { type: 'string', minLength: 1 }, type: { enum: [...TYPES] }, optional: { type: 'boolean' } }, required: ['name'], additionalProperties: false },
            description: 'Column rules, matched to the header row by name. type defaults to "any".',
          },
          strict: { type: 'boolean', description: 'Treat warnings as problems. Default false.' },
        },
        ['path'],
      ),
      annotations: READ,
    },
    handler(limits.maxResponseChars, async (a) => {
      const workbook = await book(roots, limits, a['path'] as string);
      const warnings = inspect(workbook).warnings;
      const rules = (a['columns'] as { name: string; type?: string; optional?: boolean }[] | undefined) ?? [];
      const names = new Set<string>();
      for (const rule of rules) {
        if (names.has(rule.name)) {
          throw new XlvError('XLV_E_INVALID_INPUT', `Column "${rule.name.slice(0, 100)}" has more than one rule.`, { hint: 'Give each column one rule.', details: { column: rule.name.slice(0, 100) } });
        }
        names.add(rule.name);
      }
      // Problems beyond MAX_PROBLEMS are counted, not kept (D-I04).
      const problems: Json[] = [];
      let problemsTotal = 0;
      const problem = (p: () => Json): void => {
        problemsTotal++;
        if (problems.length < MAX_PROBLEMS) problems.push(p());
      };
      let sheetName: string | null = null;
      let rowsChecked = 0;
      if (rules.length > 0) {
        const table = readTable(workbook, typeof a['sheet'] === 'string' ? { sheet: a['sheet'] } : {});
        sheetName = table.sheet;
        rowsChecked = table.rows.length;
        for (const rule of rules) {
          const column = table.columns.find((c) => c.header === rule.name);
          if (!column) {
            problem(() => ({ kind: 'missing-column', column: rule.name, cell: null, row: table.headerRow, message: `Column "${rule.name}" is missing from the header row.` }));
            continue;
          }
          table.rows.forEach((row, i) => {
            const cell = `${column.letter}${String(table.rowNumbers[i]!)}`;
            const t = typeOf(row[column.key]);
            const want = rule.type ?? 'any';
            if (t === 'empty') {
              if (rule.optional !== true) problem(() => ({ kind: 'empty', column: rule.name, cell, row: table.rowNumbers[i]!, message: `${cell} is empty; "${rule.name}" needs a value.` }));
            } else if (want !== 'any' && t !== want) {
              problem(() => ({ kind: 'type', column: rule.name, cell, row: table.rowNumbers[i]!, message: `${cell} is a ${t}; "${rule.name}" needs a ${want}.` }));
            }
          });
        }
      }
      const ok = problemsTotal === 0 && (a['strict'] !== true || warnings.length === 0);
      return { ok, sheet: sheetName, rowsChecked, problems, problemsTotal, warnings };
    }),
  );

  if (!allowWrite) return;

  server.registerTool(
    'write_table',
    {
      title: 'Write rows to a new XLSX file',
      description:
        'Writes an array of row objects (or arrays) to an XLSX file. The "report" preset (default) adds a bold header row, frozen first row, autofilter, fitted column widths and date and number formats; "plain" writes values only. Strings that look like ISO dates stay strings; pass {"date": "..."} for dates.',
      inputSchema: object(
        {
          path: PATH,
          rows: { type: 'array', items: { type: ['object', 'array'] }, description: 'Row objects (keys become columns) or arrays.' },
          columns: { type: 'array', items: { type: 'string' }, description: 'Keys to include, in this order. Default: every key, in order of first appearance.' },
          sheet: { type: 'string', description: 'Sheet name. Default "Sheet1".' },
          preset: { enum: ['report', 'plain'] },
          overwrite: OVERWRITE,
        },
        ['path', 'rows'],
      ),
      annotations: WRITE,
    },
    handler(limits.maxResponseChars, async (a) => {
      const rows = (a['rows'] as unknown[]).map((r) =>
        Array.isArray(r) ? r.map(cellInput) : Object.fromEntries(Object.entries(r as Json).map(([k, v]) => [k, cellInput(v)])),
      );
      const bytes = writeTable(rows as never, {
        ...(Array.isArray(a['columns']) ? { columns: a['columns'] as TableColumnSpec[] } : {}),
        ...(typeof a['sheet'] === 'string' ? { sheet: a['sheet'] } : {}),
        ...(typeof a['preset'] === 'string' ? { preset: a['preset'] as 'report' | 'plain' } : {}),
      });
      const w = await writeAllowed(roots, a['path'] as string, bytes, a['overwrite'] === true);
      return { ...w, rows: rows.length };
    }),
  );

  server.registerTool(
    'edit_cells',
    {
      title: 'Change cells or append rows',
      description:
        'Changes cells and appends rows in an existing XLSX file while keeping everything else in it (styles, formulas, other sheets) unchanged. The result is written to output; to change the file in place, omit output and pass "overwrite": true. xlvane does not calculate formulas; the edited file asks spreadsheet programs to recalculate them when it is opened.',
      inputSchema: object(
        {
          path: PATH,
          edits: {
            type: 'array',
            items: { type: 'object', properties: { sheet: { type: 'string', minLength: 1 }, cell: { type: 'string', pattern: '^[A-Za-z]{1,3}[1-9][0-9]{0,6}$' }, value: CELL_VALUE }, required: ['sheet', 'cell', 'value'], additionalProperties: false },
            description: 'Cells to set, e.g. [{"sheet": "Sales", "cell": "B2", "value": 1250}].',
          },
          append: {
            type: 'object',
            properties: {
              sheet: { type: 'string', minLength: 1 },
              rows: { type: 'array', items: { type: ['object', 'array'] } },
              header: { type: 'integer', minimum: 1, description: 'Row that holds the column names, for object rows. Default: detected.' },
            },
            required: ['sheet', 'rows'],
            additionalProperties: false,
            description: 'Rows to add below the last row; object keys are matched to the header row.',
          },
          output: { ...PATH, description: 'Where to write the result. Default: the input file, which requires "overwrite": true.' },
          overwrite: OVERWRITE,
        },
        ['path'],
      ),
      annotations: WRITE,
    },
    handler(limits.maxResponseChars, async (a) => {
      const path = a['path'] as string;
      const edits = (a['edits'] as { sheet: string; cell: string; value: unknown }[] | undefined) ?? [];
      const append = a['append'] as { sheet: string; rows: unknown[]; header?: number } | undefined;
      if (edits.length === 0 && append === undefined) {
        throw new XlvError('XLV_E_INVALID_INPUT', 'Nothing to do: pass edits, append or both.', { hint: 'For example {"path": "report.xlsx", "edits": [{"sheet": "Sales", "cell": "B2", "value": 1250}], "output": "updated.xlsx"}.' });
      }
      let bytes = await readAllowed(roots, path);
      const warnings: XlvWarning[] = [];
      try {
        if (edits.length > 0) {
          const r = editCells(
            bytes,
            edits.map((e): CellEdit => ({ sheet: e.sheet, cell: e.cell.toUpperCase(), value: cellInput(e.value) })),
            { limits: { maxCells: limits.maxCells } },
          );
          bytes = r.bytes;
          warnings.push(...r.warnings);
        }
        if (append) {
          const rows = append.rows.map((r) => (Array.isArray(r) ? r.map(cellInput) : Object.fromEntries(Object.entries(r as Json).map(([k, v]) => [k, cellInput(v)]))));
          const r = appendRows(bytes, append.sheet, rows as never, { limits: { maxCells: limits.maxCells }, ...(append.header !== undefined ? { header: append.header } : {}) });
          bytes = r.bytes;
          warnings.push(...r.warnings);
        }
      } catch (error) {
        throw withPath(error, path);
      }
      const output = typeof a['output'] === 'string' ? a['output'] : path;
      const w = await writeAllowed(roots, output, bytes, a['overwrite'] === true);
      return { ...w, cellsSet: edits.length, rowsAppended: append?.rows.length ?? 0, warnings };
    }),
  );

  server.registerTool(
    'convert',
    {
      title: 'Convert between XLSX, CSV and JSON',
      description:
        'Converts a sheet of an XLSX file to CSV, TSV or JSON (an array of row objects), or CSV, TSV or JSON to a plain XLSX file. Formats come from the file extensions. CSV output escapes text that starts with =, +, -, @ and similar characters, so spreadsheet programs do not run it as a formula; dates are written as ISO 8601.',
      inputSchema: object(
        {
          input: PATH,
          output: PATH,
          sheet: { type: 'string', description: 'Sheet to convert (XLSX input), or the sheet name (XLSX output).' },
          delimiter: { type: 'string', minLength: 1, maxLength: 1, description: 'CSV delimiter. Default "," (tab for .tsv).' },
          overwrite: OVERWRITE,
        },
        ['input', 'output'],
      ),
      annotations: WRITE,
    },
    handler(limits.maxResponseChars, async (a) => {
      const input = a['input'] as string;
      const output = a['output'] as string;
      const from = extension(input);
      const to = extension(output);
      const tables = ['csv', 'tsv', 'json'];
      const books = ['xlsx', 'xlsm'];
      const sheet = typeof a['sheet'] === 'string' ? { sheet: a['sheet'] } : {};
      let data: Uint8Array | string;
      if (books.includes(from) && tables.includes(to)) {
        const workbook = await book(roots, limits, input);
        data = to === 'json' ? jsonText(toJSON(workbook, sheet).rows, roots.maxFileBytes, output) : toCSV(workbook, { ...sheet, delimiter: (a['delimiter'] as string | undefined) ?? (to === 'tsv' ? '\t' : ',') });
      } else if (tables.includes(from) && to === 'xlsx') {
        let text: string;
        try {
          text = new TextDecoder('utf-8', { fatal: true }).decode(await readAllowed(roots, input));
        } catch (error) {
          if (error instanceof XlvError) throw error;
          throw new XlvError('XLV_E_INVALID_INPUT', `"${input}" is not valid UTF-8 text.`, { hint: 'Save the CSV or JSON file as UTF-8 and convert it again.', details: { path: input } });
        }
        try {
          data =
            from === 'json'
              ? fromJSON(text, { ...sheet, preset: 'plain', maxCells: limits.maxCells })
              : fromCSV(text, { ...sheet, delimiter: (a['delimiter'] as string | undefined) ?? (from === 'tsv' ? '\t' : ','), maxCells: limits.maxCells });
        } catch (error) {
          throw withPath(error, input);
        }
      } else {
        throw new XlvError('XLV_E_INVALID_INPUT', `Cannot convert .${from || '?'} to .${to || '?'}.`, {
          hint: 'Supported: .xlsx/.xlsm to .csv/.tsv/.json, and .csv/.tsv/.json to .xlsx.',
          details: { input, output },
        });
      }
      const w = await writeAllowed(roots, output, data, a['overwrite'] === true);
      return { ...w, from, to };
    }),
  );
}
