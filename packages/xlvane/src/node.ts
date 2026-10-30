// Node.js-only entry point (`xlvane/node`): file-system helpers around the core API. Uses
// fs.promises.stat, readFile, writeFile and open (Node.js fs documentation; F-NODE-03, F-NODE-04,
// F-NODE-11 … F-NODE-13).
import { openAsBlob } from 'node:fs';
import { open, unlink, writeFile, type FileHandle } from 'node:fs/promises';
import { XlvError } from './errors';
import { inspect, type InspectOptions, type InspectReport } from './inspect';
import { resolveLimits } from './limits';
import { checkFile, describePath, fileError, readInput } from './util/node-files';
import type { Workbook } from './model';
import { readWorkbook, type ReadOptions } from './read';
import { writeWorkbook, type NewWorkbook } from './write';
import { openWorkbookStream, STREAM_LIMIT_DEFAULTS, type OpenWorkbookStreamOptions, type WorkbookStream } from './stream-read';
import { writeWorkbookStream, type WriteWorkbookStreamOptions } from './stream-write';

/** File handles of writes in progress; see writeWorkbookStreamFile. */
const OPEN_HANDLES = new Set<FileHandle>();

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

/**
 * Options for {@link writeWorkbookFile}.
 *
 * @example
 * ```ts
 * import { createWorkbook } from 'xlvane';
 * import { writeWorkbookFile, type WriteFileOptions } from 'xlvane/node';
 *
 * const workbook = createWorkbook();
 * workbook.addSheet('Data').set('A1', 1);
 * const options: WriteFileOptions = { overwrite: true };
 * await writeWorkbookFile(workbook, 'out.xlsx', options);
 * ```
 */
export interface WriteFileOptions {
  /** Replace an existing file. Default `false`: an existing file is never overwritten. */
  readonly overwrite?: boolean;
}

/**
 * Writes a workbook created with `createWorkbook` to an XLSX file. Node.js only; in other
 * runtimes use `writeWorkbook` from `xlvane` and save the bytes yourself. An existing file is
 * not replaced unless `overwrite` is `true` (`XLV_E_FILE_EXISTS`); the check and the write are a
 * single exclusive-create operation, so no other process can slip a file in between.
 *
 * @example
 * ```ts
 * import { createWorkbook } from 'xlvane';
 * import { writeWorkbookFile } from 'xlvane/node';
 *
 * const workbook = createWorkbook();
 * workbook.addSheet('Data').set('A1', 'Hello');
 * await writeWorkbookFile(workbook, 'hello.xlsx');
 * ```
 */
export async function writeWorkbookFile(workbook: NewWorkbook, path: string | URL, options?: WriteFileOptions): Promise<void> {
  const overwrite = options?.overwrite ?? false;
  if (typeof overwrite !== 'boolean') {
    throw new XlvError('XLV_E_INVALID_OPTION', 'The option "overwrite" must be true or false.', {
      hint: 'Pass { overwrite: true } to replace an existing file.',
      details: { option: 'overwrite' },
    });
  }
  const bytes = writeWorkbook(workbook);
  const shown = describePath(path);
  try {
    // 'wx' fails if the path exists; 'w' creates or truncates (Node.js fs "File system flags", F-NODE-04).
    await writeFile(path, bytes, { flag: overwrite ? 'w' : 'wx' });
  } catch (error) {
    const code = typeof error === 'object' && error !== null && 'code' in error ? String((error as { code: unknown }).code) : 'UNKNOWN';
    if (code === 'EEXIST') {
      throw new XlvError('XLV_E_FILE_EXISTS', `"${shown}" already exists.`, {
        hint: 'Choose another path, or pass { overwrite: true } to replace the file.',
        details: { path: shown },
        cause: error,
      });
    }
    throw new XlvError('XLV_E_FILE_WRITE', `Cannot write "${shown}" (${code}).`, {
      hint:
        code === 'ENOENT'
          ? 'The directory does not exist. Create it first.'
          : code === 'EACCES' || code === 'EPERM'
            ? 'The process is not allowed to write there. Check the permissions.'
            : 'Check that the directory exists and is writable.',
      details: { path: shown, code },
      cause: error,
    });
  }
}

/**
 * Opens an XLSX file for streaming. The file is read through `fs.openAsBlob`, so neither the
 * file nor its cells are held in memory; see `openWorkbookStream`. The file must not change
 * while it is being read.
 *
 * @example
 * ```ts
 * import { openWorkbookStreamFile } from 'xlvane/node';
 *
 * const book = await openWorkbookStreamFile('large.xlsx');
 * let rows = 0;
 * for await (const _row of book.rows(book.sheets[0]!.name)) rows++;
 * console.log(rows);
 * ```
 */
export async function openWorkbookStreamFile(path: string | URL, options?: OpenWorkbookStreamOptions): Promise<WorkbookStream> {
  await checkFile(path, resolveLimits({ ...STREAM_LIMIT_DEFAULTS, ...options?.limits }));
  let blob: Blob;
  try {
    blob = await openAsBlob(path);
  } catch (error) {
    throw fileError(describePath(path), error);
  }
  return openWorkbookStream(blob, options);
}

/**
 * Writes a streamed workbook (see `writeWorkbookStream`) to a file with constant memory. An
 * existing file is not replaced unless `overwrite` is `true` (`XLV_E_FILE_EXISTS`). If writing
 * fails part-way, for example because a row holds an invalid value, the incomplete file is
 * removed and the error is thrown. The file is opened when the first bytes are ready, so an
 * existing file is reported then, and a write whose rows never arrive creates no file.
 *
 * @example
 * ```ts
 * import { writeWorkbookStreamFile } from 'xlvane/node';
 *
 * function* rows() {
 *   for (let i = 1; i <= 1_000_000; i++) yield [i, `item ${i}`];
 * }
 * await writeWorkbookStreamFile('items.xlsx', { sheets: [{ name: 'Items', rows: rows() }] });
 * ```
 */
export async function writeWorkbookStreamFile(path: string | URL, spec: WriteWorkbookStreamOptions, options?: WriteFileOptions): Promise<void> {
  const overwrite = options?.overwrite ?? false;
  if (typeof overwrite !== 'boolean') {
    throw new XlvError('XLV_E_INVALID_OPTION', 'The option "overwrite" must be true or false.', {
      hint: 'Pass { overwrite: true } to replace an existing file.',
      details: { option: 'overwrite' },
    });
  }
  const stream = writeWorkbookStream(spec);
  const shown = describePath(path);
  const reader = stream.getReader();
  let handle: FileHandle | null = null;
  let created = false;
  try {
    // The file is opened only when the first bytes exist, so a writer that is never started
    // (e.g. sheets that never arrive) holds no file. The handle is closed explicitly in every
    // case and is kept in OPEN_HANDLES meanwhile: Node.js 25 throws, uncatchably, when a
    // FileHandle is closed by garbage collection (DEP0137, F-NODE-11).
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      if (handle === null) {
        handle = await open(path, overwrite ? 'w' : 'wx');
        OPEN_HANDLES.add(handle);
        created = true;
      }
      // write() may write fewer bytes than given; it reports how many (F-NODE-12).
      for (let at = 0; at < value.byteLength; ) at += (await handle.write(value, at)).bytesWritten;
    }
    if (handle === null) {
      // An empty stream still gives a (valid, if unusual) file, as before.
      handle = await open(path, overwrite ? 'w' : 'wx');
      OPEN_HANDLES.add(handle);
      created = true;
    }
  } catch (error) {
    await reader.cancel(error).catch(() => undefined);
    if (handle !== null) {
      await handle.close().catch(() => undefined);
      OPEN_HANDLES.delete(handle);
      handle = null;
    }
    if (created) await unlink(path).catch(() => undefined);
    if (error instanceof XlvError) throw error;
    const code = typeof error === 'object' && error !== null && 'code' in error ? String((error as { code: unknown }).code) : 'UNKNOWN';
    if (code === 'EEXIST') {
      throw new XlvError('XLV_E_FILE_EXISTS', `"${shown}" already exists.`, {
        hint: 'Choose another path, or pass { overwrite: true } to replace the file.',
        details: { path: shown },
        cause: error,
      });
    }
    throw new XlvError('XLV_E_FILE_WRITE', `Cannot write "${shown}" (${code}).`, {
      hint: code === 'ENOENT' ? 'The directory does not exist. Create it first.' : 'Check that the directory exists and is writable.',
      details: { path: shown, code },
      cause: error,
    });
  }
  try {
    await handle.close();
  } finally {
    OPEN_HANDLES.delete(handle);
  }
}
