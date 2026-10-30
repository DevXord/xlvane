// Path policy of the MCP server: every file read or written must lie inside one of the
// directories given at start (--root), after symbolic links are resolved (docs/decisions.md D-703).
import { lstat, readFile, realpath, rename, rm, stat, writeFile } from 'node:fs/promises';
import { basename, dirname, isAbsolute, join, resolve, sep } from 'node:path';
import { XlvError } from 'xlvane';

/** The allowed directories, resolved to real paths. */
export interface Roots {
  /** Real paths of the --root directories, in the order given. */
  readonly dirs: readonly string[];
  /** Largest file read or written, in bytes. */
  readonly maxFileBytes: number;
}

function errorCode(error: unknown): string {
  return typeof error === 'object' && error !== null && 'code' in error ? String((error as { code: unknown }).code) : 'UNKNOWN';
}

/** Resolves the --root directories; each must exist and be a directory. */
export async function resolveRoots(dirs: readonly string[], cwd: string, maxFileBytes: number): Promise<Roots> {
  if (dirs.length === 0) {
    throw new XlvError('XLV_E_INVALID_OPTION', 'No directory was given.', {
      hint: 'Start the server with at least one directory it may use, for example: xlvane-mcp --root ./data',
      details: { option: 'roots' },
    });
  }
  const out: string[] = [];
  for (const dir of dirs) {
    let real: string;
    try {
      real = await realpath(isAbsolute(dir) ? dir : resolve(cwd, dir));
    } catch (error) {
      throw new XlvError('XLV_E_FILE_READ', `The directory "${dir}" does not exist (${errorCode(error)}).`, { hint: 'Pass an existing directory to --root.', details: { path: dir } });
    }
    if (!(await stat(real)).isDirectory()) {
      throw new XlvError('XLV_E_FILE_READ', `"${dir}" is not a directory.`, { hint: 'Pass a directory to --root, not a file.', details: { path: dir } });
    }
    out.push(real);
  }
  return { dirs: out, maxFileBytes };
}

function inside(roots: Roots, real: string): boolean {
  return roots.dirs.some((root) => real === root || real.startsWith(root.endsWith(sep) ? root : root + sep));
}

function notAllowed(roots: Roots, path: string, why: string): XlvError {
  return new XlvError('XLV_E_PATH_NOT_ALLOWED', `"${path}" is outside the directories this server may use${why}.`, {
    hint: `Use a path inside ${roots.dirs.map((d) => `"${d}"`).join(' or ')}. Relative paths are resolved against the first one.`,
    details: { path, roots: [...roots.dirs] },
  });
}

function absolute(roots: Roots, path: string): string {
  if (typeof path !== 'string' || path === '' || path.includes('\0')) {
    throw new XlvError('XLV_E_INVALID_INPUT', 'The path must be a non-empty string.', { hint: 'Pass a file name such as "report.xlsx".', details: { path: String(path) } });
  }
  return isAbsolute(path) ? path : resolve(roots.dirs[0]!, path);
}

/** Reads a file inside the roots, after checking its real path and size. */
export async function readAllowed(roots: Roots, path: string): Promise<Uint8Array> {
  let real: string;
  try {
    real = await realpath(absolute(roots, path));
  } catch (error) {
    if (error instanceof XlvError) throw error;
    throw new XlvError('XLV_E_FILE_READ', `Cannot read "${path}" (${errorCode(error)}).`, {
      hint: errorCode(error) === 'ENOENT' ? 'The file does not exist. Use inspect_workbook on a file that exists, or check the name.' : 'Check that the path points to a readable file.',
      details: { path, code: errorCode(error) },
    });
  }
  if (!inside(roots, real)) throw notAllowed(roots, path, real === absolute(roots, path) ? '' : ' (after following symbolic links)');
  const info = await stat(real);
  if (!info.isFile()) throw new XlvError('XLV_E_FILE_READ', `"${path}" is not a file.`, { hint: 'Pass the path of a file.', details: { path } });
  if (info.size > roots.maxFileBytes) {
    throw new XlvError('XLV_E_LIMIT_EXCEEDED', `"${path}" is ${String(info.size)} bytes, more than the server's limit of ${String(roots.maxFileBytes)}.`, {
      hint: 'The limit is set when the server starts (--max-file-size).',
      details: { path, limit: 'maxFileBytes', value: roots.maxFileBytes },
    });
  }
  return readFile(real);
}

/** Result of a write. */
export interface Written {
  readonly path: string;
  readonly bytes: number;
  readonly replaced: boolean;
}

/**
 * Writes a file inside the roots. The real path of the directory must be inside a root and the
 * file itself must not be a symbolic link. A new file is created exclusively ('wx'); an existing
 * one is replaced only with `overwrite`, through a temporary file and a rename.
 */
export async function writeAllowed(roots: Roots, path: string, data: Uint8Array | string, overwrite: boolean): Promise<Written> {
  const target = absolute(roots, path);
  const name = basename(target);
  if (name === '' || name === '.' || name === '..') throw new XlvError('XLV_E_INVALID_INPUT', `"${path}" is not a file name.`, { hint: 'Pass a path that ends in a file name.', details: { path } });
  let dir: string;
  try {
    dir = await realpath(dirname(target));
  } catch (error) {
    throw new XlvError('XLV_E_FILE_WRITE', `The directory of "${path}" does not exist (${errorCode(error)}).`, { hint: 'Write into an existing directory inside the allowed roots.', details: { path, code: errorCode(error) } });
  }
  if (!inside(roots, dir)) throw notAllowed(roots, path, dir === dirname(target) ? '' : ' (after following symbolic links)');
  const file = join(dir, name);
  const bytes = typeof data === 'string' ? new TextEncoder().encode(data) : data;
  if (bytes.length > roots.maxFileBytes) {
    throw new XlvError('XLV_E_LIMIT_EXCEEDED', `The output would be ${String(bytes.length)} bytes, more than the server's limit of ${String(roots.maxFileBytes)}.`, {
      hint: 'Write fewer rows per file.',
      details: { path, limit: 'maxFileBytes', value: roots.maxFileBytes },
    });
  }
  let exists = false;
  try {
    const info = await lstat(file);
    exists = true;
    if (info.isSymbolicLink()) throw notAllowed(roots, path, ' (it is a symbolic link)');
    if (!info.isFile()) throw new XlvError('XLV_E_FILE_WRITE', `"${path}" is not a file.`, { hint: 'Choose a file name that is not a directory.', details: { path, code: 'EISDIR' } });
  } catch (error) {
    if (error instanceof XlvError) throw error;
  }
  if (exists && !overwrite) {
    throw new XlvError('XLV_E_FILE_EXISTS', `"${path}" already exists.`, { hint: 'Choose another path, or pass "overwrite": true to replace the file.', details: { path } });
  }
  try {
    if (exists) {
      const temp = join(dir, `.${name}.xlvane-${String(Date.now())}.tmp`);
      try {
        await writeFile(temp, bytes, { flag: 'wx' });
        await rename(temp, file);
      } catch (error) {
        await rm(temp, { force: true });
        throw error;
      }
    } else {
      // 'wx' fails if anything, including a symbolic link, appeared at the path meanwhile.
      await writeFile(file, bytes, { flag: 'wx' });
    }
  } catch (error) {
    if (errorCode(error) === 'EEXIST') {
      throw new XlvError('XLV_E_FILE_EXISTS', `"${path}" already exists.`, { hint: 'Choose another path, or pass "overwrite": true to replace the file.', details: { path } });
    }
    throw new XlvError('XLV_E_FILE_WRITE', `Cannot write "${path}" (${errorCode(error)}).`, { hint: 'Check that the directory is writable.', details: { path, code: errorCode(error) } });
  }
  return { path, bytes: bytes.length, replaced: exists };
}
