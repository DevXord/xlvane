// Runs the CLI in-process with a fake terminal, and builds the fixture files the help examples use.
import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { run } from '../../packages/xlvane/src/cli/main';
import type { CliIO } from '../../packages/xlvane/src/cli/types';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';

export interface CliResult {
  readonly code: number;
  readonly stdout: string;
  readonly stderr: string;
}

export interface RunOptions {
  readonly cwd: string;
  readonly env?: Record<string, string | undefined>;
  readonly tty?: boolean;
  readonly columns?: number;
}

/** Runs `xlvane <argv>` in-process. Colour is off unless the test turns it on. */
export async function cli(argv: readonly string[], options: RunOptions): Promise<CliResult> {
  let stdout = '';
  let stderr = '';
  const io: CliIO = {
    stdout: (t) => {
      stdout += t;
    },
    stderr: (t) => {
      stderr += t;
    },
    env: options.env ?? { NO_COLOR: '1' },
    cwd: options.cwd,
    stdoutIsTTY: options.tty ?? false,
    stderrIsTTY: options.tty ?? false,
    columns: options.columns ?? 100,
    version: '0.0.0-test',
  };
  const code = await run(argv, io);
  return { code, stdout, stderr };
}

/** Splits a command line from help into arguments (double and single quotes, no escapes). */
export function splitCommand(line: string): string[] {
  const args: string[] = [];
  const re = /"([^"]*)"|'([^']*)'|(\S+)/g;
  for (const m of line.matchAll(re)) args.push(m[1] ?? m[2] ?? m[3]!);
  return args;
}

/** A temporary directory with the files the help examples refer to. */
export async function withFixtures<T>(fn: (dir: string) => Promise<T>): Promise<T> {
  const dir = await mkdtemp(join(tmpdir(), 'xlvane-cli-'));
  try {
    await writeFixtures(dir);
    return await fn(dir);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

function report(revenueNorth: number): Uint8Array {
  const wb = createWorkbook({ properties: { title: 'Sales report', creator: 'Finance' } });
  const sales = wb.addSheet('Sales');
  sales.set('A1', 'Region', { font: { bold: true } }).set('B1', 'Revenue').set('C1', 'Date');
  sales.set('A2', 'North').set('B2', revenueNorth).set('C2', new Date(Date.UTC(2024, 0, 31)));
  sales.set('A3', 'South').set('B3', 980).set('C3', new Date(Date.UTC(2024, 1, 29)));
  sales.set('A4', 'Total').set('B4', { formula: 'SUM(B2:B3)', result: revenueNorth + 980 });
  const costs = wb.addSheet('Costs');
  costs.set('A1', 'Item').set('B1', 'Cost');
  costs.set('A2', 'Rent').set('B2', 500);
  costs.set('A3', '=cmd|calc').set('B3', 1);
  costs.column('B', { style: { numFmt: '#,##0.00' } });
  wb.defineName('Total', 'Sales!$B$4');
  return writeWorkbook(wb);
}

export async function writeFixtures(dir: string): Promise<void> {
  await writeFile(join(dir, 'report.xlsx'), report(1200));
  await writeFile(join(dir, 'old.xlsx'), report(1200));
  await writeFile(join(dir, 'new.xlsx'), report(1250));
  await writeFile(join(dir, 'data.json'), JSON.stringify([
    { name: 'Ann', city: 'Oslo', joined: '2024-03-01', score: 7.5 },
    { name: 'Ben', city: 'Kraków', joined: '2023-11-15', score: 9 },
  ]));
  await writeFile(join(dir, 'rows.json'), JSON.stringify([{ Region: 'West', Revenue: 700 }]));
  await writeFile(join(dir, 'sales.csv'), 'Region,Revenue,Date\r\nNorth,1200,2024-01-31\r\nSouth,980,2024-02-29\r\n');
  await mkdir(join(dir, 'my-app', 'src'), { recursive: true });
  await writeFile(join(dir, 'my-app', 'package.json'), '{\n    "name": "my-app",\n    "dependencies": {\n        "exceljs": "^4.4.0"\n    }\n}\n');
  await writeFile(join(dir, 'my-app', 'src', 'report.js'), "const ExcelJS = require('exceljs');\nconst wb = new ExcelJS.Workbook();\nconst ws = wb.addWorksheet('S');\nwb.addImage({ filename: 'logo.png', extension: 'png' });\n");
  await writeFile(join(dir, 'package.json'), '{\n  "name": "example",\n  "dependencies": { "exceljs": "4.4.0" }\n}\n');
}
