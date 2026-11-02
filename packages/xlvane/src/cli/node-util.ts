// node:util for the CLI, loaded with require instead of an ESM import. An ESM import of node:util
// evaluates all of its exports, and one of them reads the colour settings, so Node.js prints
// "Warning: The 'NO_COLOR' env is ignored due to the 'FORCE_COLOR' env being set" whenever both
// variables are set, even if no colour function runs; require loads the exports lazily and does
// not (F-NODE-15, checked with --trace-warnings in Node.js 22.23.4).
import { createRequire } from 'node:module';
import type * as NodeUtil from 'node:util';

/** The `node:util` module. */
export const nodeUtil = createRequire(import.meta.url)('node:util') as typeof NodeUtil;
