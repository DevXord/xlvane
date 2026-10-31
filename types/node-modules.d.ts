// Minimal declarations for the few Node.js built-in modules used by `xlvane/node`, the tests
// and the scripts. The project does not depend on @types/node (docs/decisions.md, D-006).
// Signatures follow https://nodejs.org/api/fs.html, path.html, url.html and os.html and are
// intentionally narrower than the real APIs: add a member here only when code needs it.

declare module 'node:fs/promises' {
  export interface Stats {
    readonly size: number;
    /** File type and permission bits (https://nodejs.org/api/fs.html#statsmode). */
    readonly mode: number;
    isFile(): boolean;
    isDirectory(): boolean;
    isSymbolicLink(): boolean;
  }
  export function readFile(path: string | URL): Promise<Uint8Array>;
  export function readFile(path: string | URL, encoding: 'utf8'): Promise<string>;
  export function writeFile(
    path: string | URL,
    data: Uint8Array | string | AsyncIterable<Uint8Array>,
    options?: { flag?: string },
  ): Promise<void>;
  export function unlink(path: string | URL): Promise<void>;
  /** https://nodejs.org/api/fs.html#class-filehandle (F-NODE-11). */
  export interface FileHandle {
    write(data: Uint8Array, offset?: number): Promise<{ bytesWritten: number }>;
    close(): Promise<void>;
  }
  export function open(path: string | URL, flags?: string): Promise<FileHandle>;
  export function stat(path: string | URL): Promise<Stats>;
  export function readdir(path: string | URL): Promise<string[]>;
  export function mkdir(path: string | URL, options?: { recursive?: boolean }): Promise<string | undefined>;
  export function mkdtemp(prefix: string): Promise<string>;
  export function rm(path: string | URL, options?: { recursive?: boolean; force?: boolean }): Promise<void>;
  /** https://nodejs.org/api/fs.html#fspromiseschmodpath-mode */
  export function chmod(path: string | URL, mode: number): Promise<void>;
  export function rename(oldPath: string | URL, newPath: string | URL): Promise<void>;
  export function realpath(path: string | URL): Promise<string>;
  export function lstat(path: string | URL): Promise<Stats>;
  export function symlink(target: string | URL, path: string | URL): Promise<void>;
  export function copyFile(src: string | URL, dest: string | URL): Promise<void>;
}

declare module 'node:fs' {
  /** Added in v19.8.0, stable in v22.17.0 (docs/decisions.md F-NODE-06). */
  export function openAsBlob(path: string | URL, options?: { type?: string }): Promise<Blob>;
  export function existsSync(path: string | URL): boolean;
  export function readdirSync(path: string | URL): string[];
  export function readFileSync(path: string | URL): Uint8Array;
  export function readFileSync(path: string | URL, encoding: 'utf8'): string;
}

declare module 'node:path' {
  export function join(...paths: string[]): string;
  export function resolve(...paths: string[]): string;
  export function isAbsolute(path: string): boolean;
  export function relative(from: string, to: string): string;
  export function dirname(path: string): string;
  export function basename(path: string, suffix?: string): string;
  export function extname(path: string): string;
  export const sep: string;
}

declare module 'node:url' {
  export function fileURLToPath(url: string | URL): string;
  export function pathToFileURL(path: string): URL;
}

declare module 'node:os' {
  export function tmpdir(): string;
}

declare module 'node:zlib' {
  /** Added in v22.2.0 (docs/decisions.md, F-NODE-02). */
  export function crc32(data: Uint8Array | string, value?: number): number;
  export function deflateRawSync(data: Uint8Array, options?: { level?: number }): Uint8Array;
  export function inflateRawSync(data: Uint8Array): Uint8Array;
}

declare module 'node:child_process' {
  /** https://nodejs.org/api/child_process.html#child_processspawnsynccommand-args-options (tests only). */
  export function spawnSync(
    command: string,
    args: readonly string[],
    options?: { timeout?: number; encoding?: 'utf8'; cwd?: string; env?: Record<string, string | undefined> },
  ): { status: number | null; stdout: string; stderr: string; error?: Error };
}

declare module 'node:util' {
  /** https://nodejs.org/api/util.html#utilparseargsconfig (stable since v20.0.0; F-NODE-09). */
  export interface ParseArgsOptionConfig {
    type: 'string' | 'boolean';
    multiple?: boolean;
    short?: string;
  }
  export type ParseArgsToken =
    | { kind: 'option'; name: string; rawName: string; index: number; value: string | undefined; inlineValue: boolean | undefined }
    | { kind: 'positional'; index: number; value: string }
    | { kind: 'option-terminator'; index: number };
  export function parseArgs(config: {
    args: string[];
    options: Record<string, ParseArgsOptionConfig>;
    strict: boolean;
    allowPositionals: boolean;
    tokens: true;
  }): { values: Record<string, unknown>; positionals: string[]; tokens: ParseArgsToken[] };
  /** https://nodejs.org/api/util.html#utilstyletextformat-text-options (added v20.12.0; F-NODE-10). */
  export function styleText(format: string | readonly string[], text: string, options?: { validateStream?: boolean }): string;
}

declare module 'node:process' {
  /** https://nodejs.org/api/process.html (the CLI entry point only). */
  interface WriteStream {
    write(chunk: string): boolean;
    readonly isTTY?: boolean;
    readonly columns?: number;
  }
  const process: {
    readonly argv: string[];
    readonly stdin: { once(event: 'end' | 'close', listener: () => void): unknown };
    readonly env: Record<string, string | undefined>;
    readonly stdout: WriteStream;
    readonly stderr: WriteStream;
    cwd(): string;
    exitCode: number | undefined;
  };
  export default process;
}

/** https://nodejs.org/api/crypto.html#cryptocreatehashalgorithm-options (tests only). */
declare module 'node:crypto' {
  export interface Hash {
    update(data: Uint8Array | string): Hash;
    digest(encoding: 'hex' | 'base64'): string;
    digest(): Uint8Array;
  }
  export function createHash(algorithm: string): Hash;
}
