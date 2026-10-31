import { inspect, type InspectReport } from '../../inspect';
import { str, usageError } from '../args';
import { readBook } from '../files';
import { plural, showValue, type Output } from '../output';
import type { CommandContext, CommandSpec } from '../types';
import { warningsBlock } from '../warnings';

/** A whole-number option of 1 or more. */
function countOption(ctx: CommandContext, name: string): number | undefined {
  const v = str(ctx.args, name);
  if (v === undefined) return undefined;
  if (!/^\d+$/.test(v) || Number(v) < 1 || !Number.isSafeInteger(Number(v))) {
    throw usageError(`The option "--${name}" must be a whole number of 1 or more, not "${v}".`, `For example --${name} 20.`, { command: 'inspect', option: `--${name}` });
  }
  return Number(v);
}

export const inspectCommand: CommandSpec = {
  name: 'inspect',
  summary: 'See sheets, columns and problems in a workbook',
  description:
    'Describes a workbook without changing it: its sheets (also hidden ones), used ranges, detected header rows, column types with sample values, merged cells, formulas, rich-text cells, column styles, hyperlinks, notes, document properties, defined names, and warnings such as macros or external links. Macros are never run and links are never followed. Built-in hidden names such as _xlnm._FilterDatabase are listed only with --verbose. With --json the result is the inspect() report described in node_modules/xlvane/docs/inspect-json.md.',
  usage: '<file> [options]',
  positionals: { min: 1, max: 1, names: ['file'] },
  options: {
    sheet: { type: 'string', value: 'name', description: 'Describe only this sheet' },
    'max-columns': { type: 'string', value: 'n', description: 'Columns described per sheet (default 200)' },
    'formula-samples': { type: 'string', value: 'n', description: 'Example formulas per sheet (default 10)' },
  },
  examples: [
    { command: 'xlvane inspect report.xlsx', description: 'Describe every sheet and list warnings' },
    { command: 'xlvane inspect report.xlsx --sheet Sales', description: 'Describe one sheet' },
    { command: 'xlvane inspect report.xlsx --json', description: 'The full report as JSON' },
  ],
  related: ['read', 'validate'],
  async run(ctx) {
    const file = ctx.args.positionals[0]!;
    const { workbook } = await readBook(ctx, file);
    const maxColumns = countOption(ctx, 'max-columns');
    const formulaSamples = countOption(ctx, 'formula-samples');
    const full = inspect(workbook, { ...(maxColumns !== undefined ? { maxColumns } : {}), ...(formulaSamples !== undefined ? { formulaSamples } : {}) });
    const name = str(ctx.args, 'sheet');
    let report: InspectReport = full;
    if (name !== undefined) {
      const sheet = workbook.getSheet(name);
      report = { ...full, sheets: full.sheets.filter((s) => s.name === sheet.name), warnings: full.warnings.filter((w) => w.location.sheet === undefined || w.location.sheet === sheet.name) };
    }
    return {
      result: report,
      exitCode: 0,
      render: (out) => renderReport(out, file, report, ctx.verbose),
    };
  },
};

function renderReport(out: Output, file: string, r: InspectReport, verbose: boolean): string {
  const lines = [
    `${out.bold(file)}: ${r.format}, ${plural(r.sheetCount, 'sheet')}, date system ${r.dateSystem}${r.conformance === 'strict' ? ', Strict conformance' : ''}`,
  ];
  if (r.hasMacros) lines.push('Contains macros (not run).');
  if (r.externalLinkCount > 0) lines.push(`Links to ${plural(r.externalLinkCount, 'other workbook')} (not followed).`);
  const props = Object.entries(r.properties);
  if (props.length > 0) {
    const keyWidth = Math.max(...props.map(([k]) => k.length));
    lines.push(out.bold('Properties:'), ...props.map(([k, v]) => `  ${k.padEnd(keyWidth)}  ${v}`));
  }
  // Built-in names Excel keeps hidden (_xlnm._FilterDatabase and the like) only with --verbose.
  const names = verbose ? r.definedNames : r.definedNames.filter((n) => !(n.hidden && n.name.startsWith('_xlnm.')));
  if (names.length > 0) {
    const shown = names.slice(0, 5).map((n) => `${n.name}${n.sheet === null ? '' : ` (sheet "${n.sheet}")`}${n.hidden ? ' (hidden)' : ''} =${n.refersTo}`);
    const more = names.length > 5 || r.definedNamesTruncated ? `, … (${r.definedNamesTruncated ? 'more than ' : ''}${String(names.length)} in total; see --json)` : '';
    lines.push(`${out.bold('Defined names:')} ${shown.join(', ')}${more}`);
  }
  for (const s of r.sheets) {
    lines.push('');
    const state = s.state === 'visible' ? '' : ` (${s.state})`;
    if (s.kind !== 'worksheet') {
      lines.push(`${out.bold(`Sheet "${s.name}"`)}${state}: ${s.kind}, no cells`);
      continue;
    }
    if (s.usedRange === null) {
      lines.push(`${out.bold(`Sheet "${s.name}"`)}${state}: empty`);
      continue;
    }
    lines.push(
      `${out.bold(`Sheet "${s.name}"`)}${state}: ${s.usedRange}, ${plural(s.rowCount, 'row')} × ${plural(s.columnCount, 'column')}, ${plural(s.cellCount, 'cell')}, ${plural(s.formulaCount, 'formula')}`,
    );
    lines.push(s.headerRow ? `  Header row: ${String(s.headerRow.row)} (${s.headerRow.confidence} confidence)` : '  Header row: none detected');
    if (s.mergedRangeCount > 0) lines.push(`  Merged: ${s.mergedRanges.slice(0, 10).join(', ')}${s.mergedRangeCount > 10 ? `, … (${String(s.mergedRangeCount)} in total)` : ''}`);
    if (s.columns.length > 0) {
      lines.push(
        out.table(
          ['Column', 'Header', 'Type', 'Samples'],
          s.columns.map((c) => [c.letter, c.header ?? '', c.dominantType, c.samples.map((v) => (c.dominantType === 'date' && typeof v === 'string' ? v.replace('T', ' ') : showValue(v))).join(', ')]),
          '  ',
        ),
      );
      if (s.columnsTruncated) lines.push(`  … ${plural(s.columnCount - s.columns.length, 'more column')}; use --max-columns ${String(s.columnCount)} to see them`);
    }
    if (s.formulas.length > 0) lines.push(`  Formulas: ${s.formulas.slice(0, 3).map((f) => `${f.cell} =${f.formula ?? '(shared)'}`).join(', ')}`);
    if (s.richTextCellCount > 0) lines.push(`  Rich text: ${plural(s.richTextCellCount, 'cell')} (${s.richTextCells.slice(0, 5).join(', ')}${s.richTextCellCount > 5 ? ', …' : ''})`);
    if (s.styledColumns.length > 0) lines.push(`  Column styles: ${s.styledColumns.slice(0, 10).join(', ')}${s.styledColumns.length > 10 ? ', …' : ''}`);
    if (s.hyperlinkCount > 0) {
      const shown = s.hyperlinks.slice(0, 3).map((l) => `${l.ref} → ${l.target ?? ''}${l.location === null ? '' : `${l.target === null ? '' : '#'}${l.location}`}`);
      lines.push(`  Hyperlinks: ${plural(s.hyperlinkCount, 'link')} (${shown.join(', ')}${s.hyperlinkCount > 3 ? ', …' : ''})`);
    }
    if (s.imageCount > 0) lines.push(`  Images: ${plural(s.imageCount, 'picture')} (${s.images.slice(0, 5).map((i) => `${i.cell ?? 'absolute'} ${i.format ?? 'other'}`).join(', ')}${s.imageCount > 5 ? ', …' : ''})`);
    if (s.noteCount > 0) lines.push(`  Notes: ${plural(s.noteCount, 'note')} (${s.notes.slice(0, 5).join(', ')}${s.noteCount > 5 ? ', …' : ''})`);
    if (s.tables.length > 0) lines.push(`  Tables: ${s.tables.slice(0, 5).map((t) => `${t.name} ${t.ref}`).join(', ')}${s.tables.length > 5 ? ', …' : ''}`);
    if (s.dataValidationCount > 0 || s.conditionalFormatCount > 0) {
      const parts = [s.dataValidationCount > 0 ? plural(s.dataValidationCount, 'data validation') : null, s.conditionalFormatCount > 0 ? plural(s.conditionalFormatCount, 'conditional format rule') : null];
      lines.push(`  Rules: ${parts.filter((p) => p !== null).join(', ')}`);
    }
    if (s.protected) lines.push('  Protected: yes');
    if (s.printArea !== null || s.pageBreakCount > 0) {
      const parts = [s.printArea === null ? null : `print area ${s.printArea.join(', ')}`, s.pageBreakCount > 0 ? plural(s.pageBreakCount, 'page break') : null];
      lines.push(`  Printing: ${parts.filter((p) => p !== null).join(', ')}`);
    }
  }
  if (r.warnings.length > 0) lines.push('', ...warningsBlock(out, r.warnings));
  return lines.join('\n');
}

