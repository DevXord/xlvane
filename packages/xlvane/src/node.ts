// Node.js-only entry point (`xlvane/node`): file-system helpers around the core API. Uses
// fs.promises.stat and fs.promises.readFile (Node.js fs documentation; F-NODE-03).
import { readFile, stat } from 'node:fs/promises';
import { XlvError } from './errors';
import { inspect, type InspectOptions, type InspectReport } from './inspect';
import { resolveLimits } from './limits';
import type { Workbook } from './model';
import { readWorkbook, type ReadOptions } from './read';

function describePath(path: string | URL): string {
  return typeof path === 'string' ? path : path.href;
}

async function readInput(path: string | URL, limitsOption: unknown): Promise<Uint8Array> {
  const limits = resolveLimits(limitsOption);
  const shown = describePath(path);
  let size: number;
  try {
    const info = await stat(path);
    if (!info.isFile()) {
      throw new XlvError('XLV_E_FILE_READ', `"${shown}" is not a file.`, {
        location: {},
        hint: 'Pass the path of an .xlsx file, not a directory.',
        details: { path: shown },
      });
    }
    size = info.size;
  } catch (error) {
    if (error instanceof XlvError) throw error;
    throw fileError(shown, error);
  }
  if (size > limits.maxInputBytes) {
    throw new XlvError('XLV_E_LIMIT_EXCEEDED', `"${shown}" is ${size} bytes, more than the limit.`, {
      hint: `If you trust this file, raise the limit, for example { limits: { maxInputBytes: ${String(limits.maxInputBytes * 4)} } }.`,
      details: { limit: 'maxInputBytes', value: limits.maxInputBytes },
    });
  }
  try {
    return await readFile(path);
  } catch (error) {
    throw fileError(shown, error);
  }
}

function fileError(shown: string, cause: unknown): XlvError {
  const code = typeof cause === 'object' && cause !== null && 'code' in cause ? String((cause as { code: unknown }).code) : 'UNKNOWN';
  const hint =
    code === 'ENOENT'
      ? 'Check the path: the file does not exist. Relative paths are resolved against the current working directory.'
      : code === 'EACCES' || code === 'EPERM'
        ? 'The process is not allowed to read the file. Check its permissions.'
        : 'Check that the path points to a readable file.';
  return new XlvError('XLV_E_FILE_READ', `Cannot read "${shown}" (${code}).`, { hint, details: { path: shown, code }, cause });
}

/**
 * Reads an XLSX file from disk. Node.js only; in other runtimes pass the bytes to
 * `readWorkbook` from `xlvane`. The file size is checked against `limits.maxInputBytes`
 * before the file is read.
 *
 * @example
 * ```ts
 * import { readWorkbookFile } from 'xlvane/node';
 *
 * const workbook = await readWorkbookFile('report.xlsx');
 * console.log(workbook.sheets.map((sheet) => sheet.name));
 * ```
 */
export async function readWorkbookFile(path: string | URL, options?: ReadOptions): Promise<Workbook> {
  const bytes = await readInput(path, options?.limits);
  return readWorkbook(bytes, options);
}

/**
 * Inspects an XLSX file from disk and returns the same JSON report as `inspect`.
 *
 * @example
 * ```ts
 * import { inspectFile } from 'xlvane/node';
 *
 * const report = await inspectFile('report.xlsx');
 * console.log(JSON.stringify(report, null, 2));
 * ```
 */
export async function inspectFile(path: string | URL, options?: InspectOptions): Promise<InspectReport> {
  const bytes = await readInput(path, options?.limits);
  return inspect(bytes, options);
}
