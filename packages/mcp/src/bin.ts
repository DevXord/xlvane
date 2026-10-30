#!/usr/bin/env node
// Entry point of the `xlvane-mcp` command (package.json "bin"). No top-level await.
import process from 'node:process';
import { runCli } from './server';

runCli(process.argv.slice(2)).then(
  (code) => {
    process.exitCode = code;
  },
  (error: unknown) => {
    process.stderr.write(`xlvane-mcp: ${error instanceof Error ? (error.stack ?? error.message) : String(error)}\n`);
    process.exitCode = 3;
  },
);
