// The commands of the CLI, in the order shown in help (docs/PLAN.md section 5: at most 9 verbs).
import type { CommandSpec } from '../types';
import { convertCommand, createCommand } from './convert';
import { diffCommand } from './diff';
import { editCommand } from './edit';
import { inspectCommand } from './inspect';
import { mcpCommand } from './mcp';
import { migrateCommand } from './migrate';
import { readCommand } from './read';
import { validateCommand } from './validate';

export const COMMANDS: readonly CommandSpec[] = [
  inspectCommand,
  readCommand,
  convertCommand,
  createCommand,
  editCommand,
  diffCommand,
  validateCommand,
  migrateCommand,
  mcpCommand,
];
