// Warnings in the text output of the CLI: grouped, with their count and cells, wrapped, and a
// pointer to docs/errors.md (docs/decisions.md D-R02). The code is always written out; colour only
// marks it.
import type { XlvWarning } from '../errors';
import { DOCS_URL } from './help';
import { commandLine, escapeControls, plural, wrap, type Output } from './output';

interface Group {
  readonly code: string;
  readonly message: string;
  sheets: string[];
  parts: string[];
  count: number;
  cells: string[];
}

function group(warnings: readonly XlvWarning[]): Group[] {
  const groups = new Map<string, Group>();
  for (const w of warnings) {
    const key = `${w.code}\u0000${w.message}`;
    let g = groups.get(key);
    if (g === undefined) {
      g = { code: w.code, message: w.message, sheets: [], parts: [], count: 0, cells: [] };
      groups.set(key, g);
    }
    if (w.location.sheet !== undefined && !g.sheets.includes(w.location.sheet)) g.sheets.push(w.location.sheet);
    else if (w.location.sheet === undefined && w.location.part !== undefined && !g.parts.includes(w.location.part)) g.parts.push(w.location.part);
    g.count += w.count;
    for (const c of [...(w.cells ?? []), ...(w.location.cell !== undefined ? [w.location.cell] : [])]) if (!g.cells.includes(c)) g.cells.push(c);
  }
  return [...groups.values()];
}

function list(items: readonly string[], max: number): string {
  return items.slice(0, max).join(', ') + (items.length > max ? ', …' : '');
}

function where(g: Group): string {
  const parts: string[] = [];
  if (g.sheets.length === 1) parts.push(`sheet "${g.sheets[0]!}"`);
  else if (g.sheets.length > 1) parts.push(`${String(g.sheets.length)} sheets: ${list(g.sheets.map((s) => `"${s}"`), 4)}`);
  else if (g.parts.length > 0) parts.push(list(g.parts, 2));
  if (g.cells.length > 0) parts.push(`${plural(Math.max(g.count, g.cells.length), 'cell')}: ${list(g.cells, 5)}`);
  else if (g.count > 1) parts.push(`${String(g.count)} times`);
  return parts.join(', ');
}

/** Warning lines: a header line per group and the message wrapped below it. */
export function warningLines(out: Output, warnings: readonly XlvWarning[]): string[] {
  const width = Math.min(out.width, 100);
  return group(warnings).flatMap((g) => {
    const at = escapeControls(where(g));
    return [`  ${out.yellow('warning')}  ${out.bold(g.code)}${at ? `  ${at}` : ''}`, wrap(escapeControls(g.message), width, '    ')];
  });
}

/** "Warnings (N):", the warnings, and where the codes are explained. */
export function warningsBlock(out: Output, warnings: readonly XlvWarning[]): string[] {
  if (warnings.length === 0) return [];
  return [
    out.bold(out.yellow(`Warnings (${String(warnings.length)}):`)),
    ...warningLines(out, warnings),
    out.dim(`  Each code is explained in docs/errors.md of the xlvane package (${DOCS_URL}/blob/main/docs/errors.md).`),
  ];
}

/** One line for commands whose stdout is data: how many warnings, and where to see them. */
export function warningNote(warnings: readonly XlvWarning[], file: string): string | null {
  if (warnings.length === 0) return null;
  const codes = [...new Set(warnings.map((w) => w.code))];
  return `note: ${plural(warnings.length, 'warning')} while reading ${file} (${list(codes, 3)}); run "${commandLine(['validate', file])}" for details.`;
}
