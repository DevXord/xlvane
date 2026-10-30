import { inspect, type InspectReport } from '../../inspect';
import { str } from '../args';
import { readBook } from '../files';
import { plural, showValue, type Output } from '../output';
import type { CommandSpec } from '../types';

/** Warnings as lines; the code is always written out, colour only marks it. */
export function warningLines(out: Output, warnings: InspectReport['warnings']): string[] {
  return warnings.map((w) => {
    const at = [w.location.sheet ? `sheet "${w.location.sheet}"` : '', w.location.cell ?? '', w.location.part ?? ''].filter((p) => p !== '').join(', ');
    const count = w.count > 1 ? ` (${String(w.count)} times)` : '';
    return `  ${out.yellow('warning')} ${w.code}: ${w.message}${at ? ` [${at}]` : ''}${count}`;
  });
}

export const inspectCommand: CommandSpec = {
  name: 'inspect',
  summary: 'See sheets, columns and problems in a workbook',
  description:
    'Describes a workbook without changing it: its sheets (also hidden ones), used ranges, detected header rows, column types with sample values, merged cells, formulas, rich-text cells, column styles, hyperlinks, notes, document properties, defined names, and warnings such as macros or external links. Macros are never run and links are never followed. With --json the result is the inspect() report described in docs/inspect-json.md.',
  usage: '<file> [options]',
  positionals: { min: 1, max: 1, names: ['file'] },
  options: {
    sheet: { type: 'string', value: 'name', description: 'Describe only this sheet' },
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
    const full = inspect(workbook);
    const name = str(ctx.args, 'sheet');
    let report: InspectReport = full;
    if (name !== undefined) {
      const sheet = workbook.getSheet(name);
      report = { ...full, sheets: full.sheets.filter((s) => s.name === sheet.name), warnings: full.warnings.filter((w) => w.location.sheet === undefined || w.location.sheet === sheet.name) };
    }
    return {
      result: report,
      exitCode: 0,
      render: (out) => renderReport(out, file, report),
    };
  },
};

function renderReport(out: Output, file: string, r: InspectReport): string {
  const lines = [
    `${out.bold(file)}: ${r.format}, ${plural(r.sheetCount, 'sheet')}, date system ${r.dateSystem}${r.conformance === 'strict' ? ', Strict conformance' : ''}`,
  ];
  if (r.hasMacros) lines.push('Contains macros (not run).');
  if (r.externalLinkCount > 0) lines.push(`Links to ${plural(r.externalLinkCount, 'other workbook')} (not followed).`);
  const props = Object.entries(r.properties);
  if (props.length > 0) lines.push(`Properties: ${props.map(([k, v]) => `${k} ${JSON.stringify(v)}`).join(', ')}`);
  if (r.definedNames.length > 0) {
    const shown = r.definedNames.slice(0, 5).map((n) => `${n.name}${n.sheet === null ? '' : ` (sheet "${n.sheet}")`}${n.hidden ? ' (hidden)' : ''} =${n.refersTo}`);
    const more = r.definedNames.length > 5 || r.definedNamesTruncated ? `, … (${r.definedNamesTruncated ? 'more than ' : ''}${String(r.definedNames.length)} in total; see --json)` : '';
    lines.push(`Defined names: ${shown.join(', ')}${more}`);
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
    if (s.mergedRanges.length > 0) lines.push(`  Merged: ${s.mergedRanges.slice(0, 10).join(', ')}${s.mergedRanges.length > 10 ? `, … (${String(s.mergedRanges.length)} in total)` : ''}`);
    if (s.columns.length > 0) {
      lines.push(
        out.table(
          ['Column', 'Header', 'Type', 'Samples'],
          s.columns.map((c) => [c.letter, c.header ?? '', c.dominantType, c.samples.map((v) => (c.dominantType === 'date' && typeof v === 'string' ? v.replace(/T00:00:00\.000$/, '') : showValue(v))).join(', ')]),
          '  ',
        ),
      );
      if (s.columnsTruncated) lines.push(`  … more columns; see --json`);
    }
    if (s.formulas.length > 0) lines.push(`  Formulas: ${s.formulas.slice(0, 3).map((f) => `${f.cell} =${f.formula ?? '(shared)'}`).join(', ')}`);
    if (s.richTextCellCount > 0) lines.push(`  Rich text: ${plural(s.richTextCellCount, 'cell')} (${s.richTextCells.slice(0, 5).join(', ')}${s.richTextCellCount > 5 ? ', …' : ''})`);
    if (s.styledColumns.length > 0) lines.push(`  Column styles: ${s.styledColumns.slice(0, 10).join(', ')}${s.styledColumns.length > 10 ? ', …' : ''}`);
    if (s.hyperlinkCount > 0) {
      const shown = s.hyperlinks.slice(0, 3).map((l) => `${l.ref} → ${l.target ?? ''}${l.location === null ? '' : `${l.target === null ? '' : '#'}${l.location}`}`);
      lines.push(`  Hyperlinks: ${plural(s.hyperlinkCount, 'link')} (${shown.join(', ')}${s.hyperlinkCount > 3 ? ', …' : ''})`);
    }
    if (s.noteCount > 0) lines.push(`  Notes: ${plural(s.noteCount, 'note')} (${s.notes.slice(0, 5).join(', ')}${s.noteCount > 5 ? ', …' : ''})`);
  }
  if (r.warnings.length > 0) {
    lines.push('', 'Warnings:', ...warningLines(out, r.warnings));
  }
  return lines.join('\n');
}

