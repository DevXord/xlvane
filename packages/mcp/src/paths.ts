// Path policy of the MCP server: every file read or written must lie inside one of the
// directories given at start (--root), after symbolic links are resolved (docs/decisions.md D-703,
// D-I02, D-I04).
import { constants } from 'node:fs';
import { lstat, open, readlink, realpath, rename, rm, stat, writeFile } from 'node:fs/promises';
import { basename, dirname, isAbsolute, join, parse, resolve, sep } from 'node:path';
import { XlvError } from 'xlvane';

/** The allowed directories, resolved to real paths. */
export interface Roots {
  /** Real paths of the --root directories, in the order given. */
  readonly dirs: readonly string[];
  /** The --root directories as given, made absolute and normalized (for the check before any file system access). */
  readonly given: readonly string[];
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
  const given: string[] = [];
  for (const dir of dirs) {
    const full = resolve(cwd, dir);
    let real: string;
    try {
      real = await realpath(full);
    } catch (error) {
      throw new XlvError('XLV_E_FILE_READ', `The directory "${dir}" does not exist (${errorCode(error)}).`, { hint: 'Pass an existing directory to --root.', details: { path: dir } });
    }
    if (!(await stat(real)).isDirectory()) {
      throw new XlvError('XLV_E_FILE_READ', `"${dir}" is not a directory.`, { hint: 'Pass a directory to --root, not a file.', details: { path: dir } });
    }
    out.push(real);
    given.push(full);
  }
  return { dirs: out, given, maxFileBytes };
}

function within(dirs: readonly string[], path: string): boolean {
  return dirs.some((root) => path === root || path.startsWith(root.endsWith(sep) ? root : root + sep));
}

function inside(roots: Roots, real: string): boolean {
  return within(roots.dirs, real);
}

/**
 * The real path of `target`. When it cannot be resolved (it does not exist, or a file stands where
 * a directory should), the nearest ancestor that exists is resolved instead: if that leads outside
 * the roots, the path is refused as outside, so the answer does not depend on what exists behind a
 * symbolic link inside a root (D-J01). Otherwise `onMissing` gives the error for the missing path.
 */
async function realInside(roots: Roots, path: string, target: string, onMissing: (error: unknown) => XlvError): Promise<string> {
  let real: string;
  try {
    real = await realpath(target);
  } catch (error) {
    if (await leadsOutside(roots, target)) throw notAllowed(roots, path, '');
    throw onMissing(error);
  }
  if (!inside(roots, real)) throw notAllowed(roots, path, '');
  return real;
}

/**
 * Whether a path that does not resolve would lead outside the roots. The path is walked one
 * component at a time from the file system root, as the system resolves it: a symbolic link (also
 * one whose target is missing) is replaced by its text, so ".." in that text goes to the parent of
 * the directory the link leads to, not of the link (D-J01, D-K03). From the first component that
 * does not exist the rest of the path is taken as written.
 */
async function leadsOutside(roots: Roots, path: string): Promise<boolean> {
  const split = (p: string): string[] => p.slice(parse(p).root.length).split(/[\\/]/).filter((part) => part !== '');
  const pending = split(path);
  let current = parse(path).root;
  let links = 0;
  while (pending.length > 0) {
    const part = pending.shift()!;
    if (part === '.') continue;
    if (part === '..') {
      current = dirname(current);
      continue;
    }
    const next = join(current, part);
    let info;
    try {
      info = await lstat(next);
    } catch {
      return !inside(roots, resolve(next, ...pending));
    }
    if (!info.isSymbolicLink()) {
      current = next;
      continue;
    }
    // A chain of links longer than the system follows is refused like a path outside.
    if (++links > 40) return true;
    let text: string;
    try {
      text = await readlink(next);
    } catch {
      return true;
    }
    if (isAbsolute(text)) {
      const target = resolve(text);
      current = parse(target).root;
      pending.unshift(...split(target));
    } else pending.unshift(...text.split(/[\\/]/).filter((part) => part !== ''));
  }
  return !inside(roots, current);
}

function notAllowed(roots: Roots, path: string, why: string): XlvError {
  return new XlvError('XLV_E_PATH_NOT_ALLOWED', `"${path}" is outside the directories this server may use${why}.`, {
    hint: `Use a path inside ${roots.dirs.map((d) => `"${d}"`).join(' or ')}. Relative paths are resolved against the first one.`,
    details: { path, roots: [...roots.dirs] },
  });
}

/**
 * The absolute, normalized path of a tool argument. A path that is not inside a root even before
 * symbolic links are resolved is refused here, before the file system is touched, so the answer
 * for a path outside the roots does not depend on what exists there (D-I04).
 */
function absolute(roots: Roots, path: string): string {
  if (typeof path !== 'string' || path === '' || path.includes('\0')) {
    throw new XlvError('XLV_E_INVALID_INPUT', 'The path must be a non-empty string.', { hint: 'Pass a file name such as "report.xlsx".', details: { path: String(path).slice(0, 200) } });
  }
  const full = isAbsolute(path) ? resolve(path) : resolve(roots.dirs[0]!, path);
  if (!within(roots.dirs, full) && !within(roots.given, full)) throw notAllowed(roots, path, '');
  // On Windows "name:stream" opens an NTFS alternate data stream of "name".
  if (sep === '\\' && full.slice(2).includes(':')) {
    throw new XlvError('XLV_E_INVALID_INPUT', `"${path}" contains ":" after the drive.`, { hint: 'Use a file name without ":".', details: { path } });
  }
  return full;
}

/** Points in {@link readAllowedWith} where tests change the file system (races). */
export interface ReadHooks {
  readonly afterCheck?: () => Promise<void>;
  readonly afterOpen?: () => Promise<void>;
  readonly afterVerify?: () => Promise<void>;
}

/** Reads a file inside the roots, after checking its real path and size. */
export function readAllowed(roots: Roots, path: string): Promise<Uint8Array> {
  return readAllowedWith(roots, path, {});
}

/**
 * {@link readAllowed} with hooks for tests. The file is opened first and checked afterwards: the
 * path must still resolve inside a root to the same file (device and inode) as the one opened,
 * so a file or directory swapped for a symbolic link between the check and the read is refused
 * (D-I02). At most the size limit plus one byte is read, whatever the size was when checked.
 */
export async function readAllowedWith(roots: Roots, path: string, hooks: ReadHooks): Promise<Uint8Array> {
  const target = absolute(roots, path);
  const resolveInside = (): Promise<string> =>
    realInside(
      roots,
      path,
      target,
      (error) =>
        new XlvError('XLV_E_FILE_READ', `Cannot read "${path}" (${errorCode(error)}).`, {
          hint: errorCode(error) === 'ENOENT' ? 'The file does not exist. Use inspect_workbook on a file that exists, or check the name.' : 'Check that the path points to a readable file.',
          details: { path, code: errorCode(error) },
        }),
    );
  const notAFile = (): XlvError => new XlvError('XLV_E_FILE_READ', `"${path}" is not a file.`, { hint: 'Pass the path of a file.', details: { path } });
  // Only regular files are opened: opening a FIFO or a device could block (D-I04).
  // A file removed between the real-path check and stat is a read error, not an internal one.
  const statOf = async (): Promise<Awaited<ReturnType<typeof stat>>> => {
    const real = await resolveInside();
    try {
      return await stat(real);
    } catch (error) {
      throw new XlvError('XLV_E_FILE_READ', `Cannot read "${path}" (${errorCode(error)}).`, { hint: 'Check that the file still exists.', details: { path, code: errorCode(error) } });
    }
  };
  if (!(await statOf()).isFile()) throw notAFile();
  await hooks.afterCheck?.();
  let handle;
  try {
    // O_NONBLOCK: a file swapped for a FIFO meanwhile does not block the open (POSIX; 0 where absent).
    handle = await open(target, constants.O_RDONLY | (constants.O_NONBLOCK ?? 0));
  } catch (error) {
    throw new XlvError('XLV_E_FILE_READ', `Cannot read "${path}" (${errorCode(error)}).`, { hint: 'Check that the path points to a readable file.', details: { path, code: errorCode(error) } });
  }
  try {
    await hooks.afterOpen?.();
    const opened = await handle.stat();
    // First: the path must still lead, inside a root, to the file that was opened. Only then
    // are its type and size reported, so nothing is said about a file outside the roots (D-I04).
    const now = await statOf();
    if (now.dev !== opened.dev || now.ino !== opened.ino) {
      throw new XlvError('XLV_E_FILE_READ', `"${path}" changed while it was being read.`, { hint: 'Read the file again when nothing else is writing to it.', details: { path } });
    }
    if (!opened.isFile()) throw notAFile();
    const tooLarge = (size: number): XlvError =>
      new XlvError('XLV_E_LIMIT_EXCEEDED', `"${path}" is ${String(size)} bytes or more, more than the server's limit of ${String(roots.maxFileBytes)}.`, {
        hint: 'The limit is set when the server starts (--max-file-size).',
        details: { path, limit: 'maxFileBytes', value: roots.maxFileBytes },
      });
    if (opened.size > roots.maxFileBytes) throw tooLarge(opened.size);
    await hooks.afterVerify?.();
    const chunks: Uint8Array[] = [];
    let total = 0;
    for (;;) {
      const buffer = new Uint8Array(Math.min(1 << 20, roots.maxFileBytes + 1 - total));
      if (buffer.length === 0) throw tooLarge(total);
      const { bytesRead } = await handle.read(buffer, 0, buffer.length, null);
      if (bytesRead === 0) break;
      chunks.push(buffer.subarray(0, bytesRead));
      total += bytesRead;
    }
    const out = new Uint8Array(total);
    let at = 0;
    for (const c of chunks) {
      out.set(c, at);
      at += c.length;
    }
    return out;
  } finally {
    await handle.close();
  }
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
  const dir = await realInside(
    roots,
    path,
    dirname(target),
    (error) =>
      new XlvError('XLV_E_FILE_WRITE', `The directory of "${path}" does not exist (${errorCode(error)}).`, { hint: 'Write into an existing directory inside the allowed roots.', details: { path, code: errorCode(error) } }),
  );
  const file = join(dir, name);
  const bytes = typeof data === 'string' ? new TextEncoder().encode(data) : data;
  if (bytes.length > roots.maxFileBytes) {
    throw new XlvError('XLV_E_LIMIT_EXCEEDED', `The output would be ${String(bytes.length)} bytes, more than the server's limit of ${String(roots.maxFileBytes)}.`, {
      hint: 'Write fewer rows per file.',
      details: { path, limit: 'maxFileBytes', value: roots.maxFileBytes },
    });
  }
  let existing: { mode: number } | null = null;
  try {
    const info = await lstat(file);
    if (info.isSymbolicLink()) throw notAllowed(roots, path, ' (it is a symbolic link)');
    if (!info.isFile()) throw new XlvError('XLV_E_FILE_WRITE', `"${path}" is not a file.`, { hint: 'Choose a file name that is not a directory.', details: { path, code: 'EISDIR' } });
    existing = info;
  } catch (error) {
    if (error instanceof XlvError) throw error;
  }
  const exists = existing !== null;
  if (exists && !overwrite) {
    throw new XlvError('XLV_E_FILE_EXISTS', `"${path}" already exists.`, { hint: 'Choose another path, or pass "overwrite": true to replace the file.', details: { path } });
  }
  // A file without write permission for its owner is not replaced (D-I04).
  if (existing !== null && (existing.mode & 0o200) === 0) {
    throw new XlvError('XLV_E_FILE_WRITE', `"${path}" is read-only.`, { hint: 'Write to another path, or make the file writable first.', details: { path, code: 'EACCES' } });
  }
  if (existing === null) {
    try {
      // 'wx' fails if anything, including a symbolic link, appeared at the path meanwhile.
      await writeFile(file, bytes, { flag: 'wx' });
    } catch (error) {
      if (errorCode(error) === 'EEXIST') {
        throw new XlvError('XLV_E_FILE_EXISTS', `"${path}" already exists.`, { hint: 'Choose another path, or pass "overwrite": true to replace the file.', details: { path } });
      }
      throw new XlvError('XLV_E_FILE_WRITE', `Cannot write "${path}" (${errorCode(error)}).`, { hint: 'Check that the directory is writable.', details: { path, code: errorCode(error) } });
    }
    return { path, bytes: bytes.length, replaced: false };
  }
  // Replace: a new temporary file with a random name ('wx'), given the permissions of the file it
  // replaces, then renamed over it; a temporary file is removed only if this call created it (D-I04).
  const random = Array.from(globalThis.crypto.getRandomValues(new Uint8Array(8)), (b) => b.toString(16).padStart(2, '0')).join('');
  // At most 100 characters of the name, so a long name still leaves room for the suffix (D-K03).
  const temp = join(dir, `.${name.slice(0, 100)}.xlvane-${random}.tmp`);
  let created = false;
  try {
    // Readable by the owner only until it has the permissions of the file it replaces; setuid,
    // setgid and sticky bits are not carried over (D-J01).
    const handle = await open(temp, 'wx', 0o600);
    created = true;
    try {
      await handle.writeFile(bytes);
      // Through the open handle, so a file swapped in for the temporary one is not changed (D-J01).
      await handle.chmod(existing.mode & 0o777);
    } finally {
      await handle.close();
    }
    await rename(temp, file);
  } catch (error) {
    if (created) await rm(temp, { force: true });
    throw new XlvError('XLV_E_FILE_WRITE', `Cannot write "${path}" (${errorCode(error)}).`, { hint: 'Check that the directory is writable.', details: { path, code: errorCode(error) } });
  }
  return { path, bytes: bytes.length, replaced: true };
}
