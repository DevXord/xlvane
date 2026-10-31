// File access of the CLI: reading with the library limits, a progress line for large files,
// and writing that never replaces a file without --force (docs/decisions.md D-604).
import { chmod, readdir, rename, rm, stat, writeFile } from 'node:fs/promises';
import { basename, dirname, extname, isAbsolute, join, resolve } from 'node:path';
import { XlvError } from '../errors';
import type { Workbook } from '../model';
import { readWorkbook } from '../read';
import { readInput } from '../util/node-files';
import { suggest } from '../util/suggest';
import { usageError } from './args';
import { escapeControls, humanSize, type Output } from './output';
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
        ctx.io.stderr(`Reading ${escapeControls(path)} (${humanSize(size)})…`);
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

/** File names a workbook can be read from. */
const isWorkbookName = (name: string): boolean => /\.(xlsx|xlsm|xltx|xltm)$/i.test(name);

/** File names with the same extension as `path` (any name when `path` has none). */
const sameExtension = (path: string) => (name: string): boolean => extname(name).toLowerCase() === extname(path).toLowerCase();

/** Reads a file as bytes (size checked against the default limits), without a progress line. */
async function readRaw(ctx: CommandContext, path: string, accept: (name: string) => boolean): Promise<Uint8Array> {
  ctx.state.lastFile = path;
  if (path === '-') {
    throw usageError('Reading from standard input ("-") is not supported.', 'Save the data to a file first and pass its path, for example: some-command > data.csv && xlvane convert data.csv data.xlsx.', { option: '-' });
  }
  const resolved = resolvePath(ctx, path);
  try {
    return await readInput(resolved, undefined);
  } catch (error) {
    throw await withFileHint(ctx, path, notFound(asTyped(error, resolved, path), path), accept);
  }
}

/**
 * Reads a file as bytes (size checked against the default limits). For a missing file, only
 * names that `accept` allows are suggested (default: workbook files).
 */
export async function readBytes(ctx: CommandContext, path: string, accept: (name: string) => boolean = isWorkbookName): Promise<Uint8Array> {
  return progress(ctx, path, () => readRaw(ctx, path, accept));
}

/** "File not found" instead of the system error code, which people do not need to decode. */
function notFound(error: unknown, path: string): unknown {
  if (!(error instanceof XlvError) || error.details['code'] !== 'ENOENT') return error;
  return new XlvError(error.code, `File not found: "${path}".`, { location: error.location, hint: 'Check the name and the folder; paths are relative to the current directory.', details: { ...error.details, path }, cause: error });
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
  // The progress line stays up while the file is parsed, which is where large files take time.
  return progress(ctx, path, async () => {
    const bytes = await readRaw(ctx, path, isWorkbookName);
    try {
      return { bytes, workbook: readWorkbook(bytes) };
    } catch (error) {
      throw withPath(error, path);
    }
  });
}

/** Reads a text file (UTF-8, BOM removed). */
export async function readText(ctx: CommandContext, path: string): Promise<string> {
  const bytes = await readBytes(ctx, path, sameExtension(path));
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

/**
 * For a missing file, suggests a file with a similar name in the same directory: only a name the
 * command can read (`accept`), and never a file that another argument of the command already
 * names, such as the other side of a diff (D-R08).
 */
async function withFileHint(ctx: CommandContext, path: string, error: unknown, accept: (name: string) => boolean): Promise<unknown> {
  if (!(error instanceof XlvError) || error.details['code'] !== 'ENOENT') return error;
  let names: string[];
  try {
    names = await readdir(dirname(resolvePath(ctx, path)));
  } catch {
    return error;
  }
  const others = new Set(ctx.argv.filter((a) => a !== path && !a.startsWith('-')).map((a) => resolvePath(ctx, a)));
  const candidates = names.filter((n) => accept(n) && !others.has(resolvePath(ctx, join(dirname(path), n))));
  const close = suggest(basename(path), candidates, 1);
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
 * that is then renamed over the target, so a failed write leaves the old file intact; the new
 * file gets the permission bits of the old one (owner and group are not changed).
 */
export async function writeOutput(ctx: CommandContext, path: string, data: Uint8Array | string, options: { force: boolean; dryRun: boolean }): Promise<WriteOutcome> {
  const target = resolvePath(ctx, path);
  const bytes = typeof data === 'string' ? new TextEncoder().encode(data) : data;
  let exists = false;
  let mode: number | null = null;
  try {
    const info = await stat(target);
    exists = true;
    // Permission bits of the file being replaced (fs.Stats.mode; the type bits are dropped).
    mode = info.mode & 0o7777;
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
        // The replacement keeps the permissions of the file it replaces, as editing in place would.
        if (mode !== null) await chmod(temp, mode);
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

/** The line every writing command ends with; green when something was written. */
export function wroteLine(w: WriteOutcome, out?: Output): string {
  if (!w.written) {
    const text = `Dry run: would ${w.replaced ? 'replace' : 'write'} ${w.path} (${humanSize(w.bytes)}). Nothing was written.`;
    return out ? out.dim(text) : text;
  }
  const verb = w.replaced ? 'Replaced' : 'Wrote';
  return `${out ? out.green(verb) : verb} ${w.path} (${humanSize(w.bytes)}).`;
}
