// File helpers shared by xlvane/node and the CLI (Node.js only). Uses fs.promises.stat and
// readFile (Node.js fs documentation; F-NODE-03).
import { readFile, stat } from 'node:fs/promises';
import { XlvError } from '../errors';
import { resolveLimits } from '../limits';

/** The path as shown in messages. */
export function describePath(path: string | URL): string {
  return typeof path === 'string' ? path : path.href;
}

export async function checkFile(path: string | URL, limits: { readonly maxInputBytes: number }): Promise<void> {
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
}

export async function readInput(path: string | URL, limitsOption: unknown): Promise<Uint8Array> {
  await checkFile(path, resolveLimits(limitsOption));
  try {
    return await readFile(path);
  } catch (error) {
    throw fileError(describePath(path), error);
  }
}

export function fileError(shown: string, cause: unknown): XlvError {
  const code = typeof cause === 'object' && cause !== null && 'code' in cause ? String((cause as { code: unknown }).code) : 'UNKNOWN';
  const hint =
    code === 'ENOENT'
      ? 'Check the path: the file does not exist. Relative paths are resolved against the current working directory.'
      : code === 'EACCES' || code === 'EPERM'
        ? 'The process is not allowed to read the file. Check its permissions.'
        : 'Check that the path points to a readable file.';
  return new XlvError('XLV_E_FILE_READ', `Cannot read "${shown}" (${code}).`, { hint, details: { path: shown, code }, cause });
}
