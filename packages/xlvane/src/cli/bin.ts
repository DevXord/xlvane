#!/usr/bin/env node
// Entry point of the `xlvane` command (package.json "bin"). The only module of the CLI that
// touches `process`; everything else gets a CliIO (docs/decisions.md D-601). No top-level await.
import { readFileSync } from 'node:fs';
import process from 'node:process';
import { run } from './main';

function version(): string {
  try {
    // dist/cli.js → ../package.json is the xlvane package manifest.
    const pkg = JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf8')) as { version?: unknown };
    return typeof pkg.version === 'string' ? pkg.version : 'unknown';
  } catch {
    return 'unknown';
  }
}

run(process.argv.slice(2), {
  stdout: (text) => {
    process.stdout.write(text);
  },
  stderr: (text) => {
    process.stderr.write(text);
  },
  env: process.env,
  cwd: process.cwd(),
  stdoutIsTTY: process.stdout.isTTY === true,
  stderrIsTTY: process.stderr.isTTY === true,
  columns: process.stdout.columns,
  version: version(),
}).then(
  (code) => {
    process.exitCode = code;
  },
  (error: unknown) => {
    process.stderr.write(`xlvane: ${error instanceof Error ? (error.stack ?? error.message) : String(error)}\n`);
    process.exitCode = 3;
  },
);
