// Minimal declarations for the few Node.js built-in modules used by `xlvane/node`, the tests
// and the scripts. The project does not depend on @types/node (docs/decisions.md, D-006).
// Signatures follow https://nodejs.org/api/fs.html, path.html, url.html and os.html and are
// intentionally narrower than the real APIs: add a member here only when code needs it.

declare module 'node:fs/promises' {
  export interface Stats {
    readonly size: number;
    isFile(): boolean;
    isDirectory(): boolean;
  }
  export function readFile(path: string | URL): Promise<Uint8Array>;
  export function readFile(path: string | URL, encoding: 'utf8'): Promise<string>;
  export function writeFile(path: string | URL, data: Uint8Array | string, options?: { flag?: string }): Promise<void>;
  export function stat(path: string | URL): Promise<Stats>;
  export function readdir(path: string | URL): Promise<string[]>;
  export function mkdir(path: string | URL, options?: { recursive?: boolean }): Promise<string | undefined>;
  export function mkdtemp(prefix: string): Promise<string>;
  export function rm(path: string | URL, options?: { recursive?: boolean; force?: boolean }): Promise<void>;
}

declare module 'node:fs' {
  export function existsSync(path: string | URL): boolean;
  export function readdirSync(path: string | URL): string[];
  export function readFileSync(path: string | URL): Uint8Array;
  export function readFileSync(path: string | URL, encoding: 'utf8'): string;
}

declare module 'node:path' {
  export function join(...paths: string[]): string;
  export function resolve(...paths: string[]): string;
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
    options?: { timeout?: number; encoding?: 'utf8'; cwd?: string },
  ): { status: number | null; stdout: string; stderr: string; error?: Error };
}
