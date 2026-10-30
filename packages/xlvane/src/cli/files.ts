// File access of the CLI: reading with the library limits, a progress line for large files,
// and writing that never replaces a file without --force (docs/decisions.md D-604).
import { readdir, rename, rm, stat, writeFile } from 'node:fs/promises';
import { basename, dirname, isAbsolute, join, resolve } from 'node:path';
import { XlvError } from '../errors';
import type { Workbook } from '../model';
import { readWorkbook } from '../read';
import { readInput } from '../util/node-files';
import { suggest } from '../util/suggest';
import type { CommandContext } from './types';

/** Files above this size get a progress line on stderr (when it is a terminal). */
const PROGRESS_BYTES = 10 * 1024 * 1024;

/** A path argument resolved against the working directory. */
export function resolvePath(ctx: CommandContext, path: string): string {
  return isAbsolute(path) ? path : resolve(ctx.io.cwd, path);
}

async function progress<T>(ctx: CommandContext, path: string, work: () => Promise<T>): Promise<T> {
  let shown = false;
  if (!ctx.json && !ctx.quiet && ctx.io.stderrIsTTY) {
    try {
      const { size } = await stat(resolvePath(ctx, path));
      if (size > PROGRESS_BYTES) {
        ctx.io.stderr(`Reading ${path} (${(size / 1024 / 1024).toFixed(1)} MB)…`);
        shown = true;
      }
    } catch {
      // The read below reports the problem.
    }
  }
  try {
    return await work();
  } finally {
    // Carriage return and "erase in line" remove the progress line.
    if (shown) ctx.io.stderr('\r\u001b[K');
  }
}

/** Reads a file as bytes (size checked against the default limits). */
export async function readBytes(ctx: CommandContext, path: string): Promise<Uint8Array> {
  ctx.state.lastFile = path;
  return progress(ctx, path, async () => {
    const resolved = resolvePath(ctx, path);
    try {
      return await readInput(resolved, undefined);
    } catch (error) {
      throw await withFileHint(ctx, path, asTyped(error, resolved, path));
    }
  });
}

/** Shows the path as the user typed it, not resolved against the working directory. */
function asTyped(error: unknown, resolved: string, typed: string): unknown {
  if (!(error instanceof XlvError) || resolved === typed) return error;
  return new XlvError(error.code, error.message.split(resolved).join(typed), {
    location: error.location,
    hint: error.hint,
    details: { ...error.details, ...(error.details['path'] === resolved ? { path: typed } : {}) },
    cause: error,
  });
}

/** Reads a workbook. */
export async function readBook(ctx: CommandContext, path: string): Promise<{ bytes: Uint8Array; workbook: Workbook }> {
  const bytes = await readBytes(ctx, path);
  try {
    return { bytes, workbook: readWorkbook(bytes) };
  } catch (error) {
    throw withPath(error, path);
  }
}

/** Reads a text file (UTF-8, BOM removed). */
export async function readText(ctx: CommandContext, path: string): Promise<string> {
  const bytes = await readBytes(ctx, path);
  try {
    return new TextDecoder('utf-8', { fatal: true, ignoreBOM: false }).decode(bytes);
  } catch {
    throw new XlvError('XLV_E_INVALID_INPUT', `"${path}" is not valid UTF-8 text.`, {
      hint: 'Save the file as UTF-8 and try again.',
      details: { path },
    });
  }
}

/** Adds the file path to an XlvError that has none, so the message says where. */
export function withPath(error: unknown, path: string): unknown {
  if (error instanceof XlvError && error.details['path'] === undefined) {
    return new XlvError(error.code, error.message, { location: error.location, hint: error.hint, details: { ...error.details, path }, cause: error });
  }
  return error;
}

/** For a missing file, suggests files with a similar name in the same directory. */
async function withFileHint(ctx: CommandContext, path: string, error: unknown): Promise<unknown> {
  if (!(error instanceof XlvError) || error.details['code'] !== 'ENOENT') return error;
  let names: string[];
  try {
    names = await readdir(dirname(resolvePath(ctx, path)));
  } catch {
    return error;
  }
  const close = suggest(basename(path), names, 1);
  if (close.length === 0) return error;
  const candidate = join(dirname(path), close[0]!).replace(/^\.\//, '');
  return new XlvError(error.code, error.message, {
    location: error.location,
    hint: `Did you mean "${candidate}"? ${error.hint}`,
    details: { ...error.details, path, suggestions: [candidate] },
    cause: error,
  });
}

/** Result of writing (or not writing, with --dry-run) an output file. */
export interface WriteOutcome {
  readonly path: string;
  readonly bytes: number;
  readonly written: boolean;
  readonly replaced: boolean;
}

/**
 * Writes an output file. Without `force` an existing file is never replaced: the write is an
 * exclusive create ('wx'). With `force` the data goes to a temporary file in the same directory
 * that is then renamed over the target, so a failed write leaves the old file intact.
 */
export async function writeOutput(ctx: CommandContext, path: string, data: Uint8Array | string, options: { force: boolean; dryRun: boolean }): Promise<WriteOutcome> {
  const target = resolvePath(ctx, path);
  const bytes = typeof data === 'string' ? new TextEncoder().encode(data) : data;
  let exists = false;
  try {
    const info = await stat(target);
    exists = true;
    if (!info.isFile()) {
      throw new XlvError('XLV_E_FILE_WRITE', `"${path}" is a directory.`, { hint: 'Give a file name for the output.', details: { path, code: 'EISDIR' } });
    }
  } catch (error) {
    if (error instanceof XlvError) throw error;
  }
  if (exists && !options.force) {
    throw new XlvError('XLV_E_FILE_EXISTS', `"${path}" already exists.`, {
      hint: 'Add --force to replace it, or choose another output path.',
      details: { path },
    });
  }
  if (options.dryRun) return { path, bytes: bytes.length, written: false, replaced: exists };
  try {
    if (exists) {
      const temp = join(dirname(target), `.${basename(target)}.xlvane-${String(Date.now())}.tmp`);
      try {
        await writeFile(temp, bytes, { flag: 'wx' });
        await rename(temp, target);
      } catch (error) {
        await rm(temp, { force: true });
        throw error;
      }
    } else {
      await writeFile(target, bytes, { flag: 'wx' });
    }
  } catch (error) {
    const code = typeof error === 'object' && error !== null && 'code' in error ? String((error as { code: unknown }).code) : 'UNKNOWN';
    if (code === 'EEXIST') {
      throw new XlvError('XLV_E_FILE_EXISTS', `"${path}" already exists.`, { hint: 'Add --force to replace it, or choose another output path.', details: { path }, cause: error });
    }
    throw new XlvError('XLV_E_FILE_WRITE', `Cannot write "${path}" (${code}).`, {
      hint: code === 'ENOENT' ? 'The directory does not exist. Create it first.' : 'Check that the directory exists and is writable.',
      details: { path, code },
      cause: error,
    });
  }
  return { path, bytes: bytes.length, written: true, replaced: exists };
}

/** The line every writing command ends with. */
export function wroteLine(w: WriteOutcome): string {
  if (!w.written) return `Dry run: would ${w.replaced ? 'replace' : 'write'} ${w.path} (${String(w.bytes)} bytes). Nothing was written.`;
  return `${w.replaced ? 'Replaced' : 'Wrote'} ${w.path} (${String(w.bytes)} bytes).`;
}
