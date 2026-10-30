// Defined names (ECMA-376 Part 1 §18.2.5 definedName, §18.17.2.5 names; [MS-OI29500] §2.1.568;
// facts W2-23 … W2-25 in docs/facts/sml-write.md and F2-05 in docs/facts/formulas.md, D-A03).
import { XlvError } from './errors';
import { encodeXstring, escapeAttr, escapeText, firstLoneSurrogate } from './xml/writer';

/**
 * A defined name of a workbook that was read: a name for a cell, a range, a constant or a
 * formula. `refersTo` is the formula text as stored, without a leading `=`. Built-in names such
 * as `_xlnm._FilterDatabase` and `_xlnm.Print_Area` are included as they are in the file.
 *
 * @example
 * ```ts
 * import { readWorkbook, type DefinedName } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const names: readonly DefinedName[] = readWorkbook(bytes).definedNames;
 * for (const n of names) console.log(n.name, n.sheet ?? '(workbook)', n.refersTo);
 * ```
 */
export interface DefinedName {
  readonly name: string;
  /** Formula text, for example `Sheet1!$A$1:$B$10`. */
  readonly refersTo: string;
  /** The sheet the name belongs to (`localSheetId`), or `null` for a workbook-level name. */
  readonly sheet: string | null;
  readonly hidden: boolean;
  readonly comment: string | null;
}

/**
 * Options for {@link NewWorkbook.defineName}.
 *
 * @example
 * ```ts
 * import { createWorkbook, type DefinedNameOptions } from 'xlvane';
 *
 * const workbook = createWorkbook();
 * workbook.addSheet('Rates');
 * const options: DefinedNameOptions = { sheet: 'Rates', comment: 'VAT rate' };
 * workbook.defineName('Vat', 'Rates!$B$1', options);
 * ```
 */
export interface DefinedNameOptions {
  /** Makes the name local to this sheet (exact sheet name). Default: the whole workbook. */
  readonly sheet?: string;
  /** Hides the name in spreadsheet applications. Default `false`. */
  readonly hidden?: boolean;
  /** A comment, at most 255 characters. */
  readonly comment?: string;
}

/** A checked defined name ready to be written; the same shape as a name that was read. */
export type NormalizedDefinedName = DefinedName;

/** [MS-OI29500] §2.1.568 k (name) and j (comment). */
const MAX_NAME_LENGTH = 255;
const MAX_COMMENT_LENGTH = 255;
/** The writer's formula limit, as for cell formulas (write.ts). */
const MAX_REFERS_TO_LENGTH = 8192;
// Name characters: P1 §18.17.2.5 with the Office extensions of [MS-XLSX] §2.2.2 (F2-05): a
// letter, `_` or `\` first; then letters, digits, `_`, `\`, `.` and `?`; any character above
// U+007F counts as a letter.
export const NAME_RE: RegExp = /^(?:[A-Za-z_\\]|[^\x00-\x7F])(?:[A-Za-z0-9_\\.?]|[^\x00-\x7F])*$/u;
/** A1 cell references (also beyond the grid, to stay on the safe side) and R1C1 forms (F2-09). */
export const A1_NAME_RE: RegExp = /^[A-Za-z]{1,3}\d+$/;
export const R1C1_NAME_RE: RegExp = /^(?:[Rr]\d*(?:[Cc]\d*)?|[Cc]\d*(?:[Rr]\d*)?)$/;

function invalidName(name: string, message: string, hint: string): XlvError {
  return new XlvError('XLV_E_INVALID_DEFINED_NAME', message, { hint, details: { name: name.slice(0, 100) } });
}

/**
 * Checks a defined name given by a caller. `sheets` lists the workbook's sheet names, or is
 * `null` when the sheets are not known yet (the sheet is then checked later by the caller).
 */
export function checkDefinedName(name: unknown, refersTo: unknown, options: unknown, sheets: readonly string[] | null): NormalizedDefinedName {
  if (typeof name !== 'string') throw invalidName(String(name), 'A defined name must be a string.', 'For example workbook.defineName("Rates", "Sheet1!$A$1:$A$10").');
  const example = 'Use letters, digits, "_", "." and "\\", starting with a letter, "_" or "\\", for example "Sales_2026".';
  if (name.length < 1 || name.length > MAX_NAME_LENGTH) {
    throw invalidName(name, `The defined name "${name.slice(0, 40)}" has ${String(name.length)} characters; names have 1 to ${String(MAX_NAME_LENGTH)}.`, example);
  }
  if (!NAME_RE.test(name) || firstLoneSurrogate(name) !== -1) {
    throw invalidName(name, `"${name.slice(0, 40)}" is not a valid defined name.`, `${example} Spaces and characters such as "-", "!" or "$" are not allowed.`);
  }
  const upper = name.toUpperCase();
  if (upper === 'TRUE' || upper === 'FALSE' || A1_NAME_RE.test(name) || R1C1_NAME_RE.test(name)) {
    throw invalidName(name, `"${name}" cannot be a defined name because it reads as ${upper === 'TRUE' || upper === 'FALSE' ? 'a boolean' : 'a cell reference'}.`, 'Add a word or an underscore, for example "Rate_A1" or "_A1".');
  }
  if (upper.startsWith('_XL')) {
    throw invalidName(name, `Defined names starting with "_xl" are reserved for built-in names ("${name.slice(0, 40)}").`, 'Choose another name. Built-in names such as print areas are set with their own options.');
  }
  if (name.length === 2 && name.startsWith('\\')) {
    throw invalidName(name, `"${name}" cannot be a defined name: Office does not load names of a backslash and one character.`, 'Use a longer name.');
  }
  const formula = typeof refersTo === 'string' && refersTo.startsWith('=') ? refersTo.slice(1) : refersTo;
  if (typeof formula !== 'string' || formula.trim() === '' || formula.length > MAX_REFERS_TO_LENGTH || firstLoneSurrogate(formula) !== -1) {
    throw invalidName(name, `The formula of the defined name "${name}" is empty, too long or not valid text.`, `Pass what the name refers to, with or without "=", for example "Sheet1!$A$1:$B$10" or "0.19"; at most ${String(MAX_REFERS_TO_LENGTH)} characters.`);
  }
  let sheet: string | null = null;
  let hidden = false;
  let comment: string | null = null;
  if (options !== undefined) {
    if (typeof options !== 'object' || options === null || Array.isArray(options)) {
      throw new XlvError('XLV_E_INVALID_OPTION', 'Defined name options must be an object.', { hint: 'For example { sheet: "Data", hidden: true }.', details: { option: 'options' } });
    }
    for (const key of Object.keys(options)) {
      if (key !== 'sheet' && key !== 'hidden' && key !== 'comment') {
        throw new XlvError('XLV_E_INVALID_OPTION', `Unknown option "${key.slice(0, 40)}" for the defined name "${name}".`, {
          hint: 'Known options: sheet, hidden, comment.',
          details: { option: key.slice(0, 40), known: ['sheet', 'hidden', 'comment'] },
        });
      }
    }
    const o = options as Record<string, unknown>;
    if (o['sheet'] !== undefined) {
      if (typeof o['sheet'] !== 'string') throw new XlvError('XLV_E_INVALID_OPTION', 'The option "sheet" must be a sheet name.', { hint: 'For example { sheet: "Data" }.', details: { option: 'sheet' } });
      sheet = o['sheet'];
    }
    if (o['hidden'] !== undefined) {
      if (typeof o['hidden'] !== 'boolean') throw new XlvError('XLV_E_INVALID_OPTION', 'The option "hidden" must be true or false.', { hint: 'Use hidden: true or leave it out.', details: { option: 'hidden' } });
      hidden = o['hidden'];
    }
    if (o['comment'] !== undefined) {
      const c = o['comment'];
      if (typeof c !== 'string' || c.length > MAX_COMMENT_LENGTH || firstLoneSurrogate(c) !== -1) {
        throw invalidName(name, `The comment of the defined name "${name}" must be text of at most ${String(MAX_COMMENT_LENGTH)} characters.`, 'Shorten the comment.');
      }
      comment = c;
    }
  }
  if (sheet !== null && sheets !== null) checkNameSheet(name, sheet, sheets);
  return { name, refersTo: formula, sheet, hidden, comment };
}

/** Throws `XLV_E_SHEET_NOT_FOUND` if the sheet of a local name does not exist. */
export function checkNameSheet(name: string, sheet: string, sheets: readonly string[]): void {
  if (sheets.includes(sheet)) return;
  throw new XlvError('XLV_E_SHEET_NOT_FOUND', `The defined name "${name}" belongs to the sheet "${sheet.slice(0, 100)}", which the workbook does not have.`, {
    location: { sheet: sheet.slice(0, 100) },
    hint: `Add the sheet first, or use one of: ${sheets.length > 0 ? sheets.map((s) => `"${s}"`).join(', ') : '(no sheets yet)'}.`,
    details: { name, available: [...sheets] },
  });
}

/**
 * Throws `XLV_E_INVALID_DEFINED_NAME` if `added` has the same name and scope as one of
 * `existing`. Names are not case-sensitive (P1 §18.17.2.5) and are unique per sheet for local
 * names and among workbook-level names otherwise ([MS-OI29500] §2.1.568 n).
 */
export function checkUniqueName(added: NormalizedDefinedName, existing: readonly NormalizedDefinedName[]): void {
  const key = added.name.toLowerCase();
  const clash = existing.find((n) => n.sheet === added.sheet && n.name.toLowerCase() === key);
  if (clash === undefined) return;
  throw invalidName(
    added.name,
    `The ${added.sheet === null ? 'workbook' : `sheet "${added.sheet}"`} already has a defined name "${clash.name}".`,
    'Names are not case-sensitive and must be unique in their scope. Choose another name, or make it local to a sheet with { sheet: "…" }.',
  );
}

/** One `definedName` element. `localSheetId` is the 0-based index of the name's sheet. */
export function definedNameXml(n: NormalizedDefinedName, localSheetId: number | null): string {
  let xml = `<definedName name="${escapeAttr(encodeXstring(n.name))}"`;
  if (n.comment !== null) xml += ` comment="${escapeAttr(encodeXstring(n.comment))}"`;
  if (localSheetId !== null) xml += ` localSheetId="${String(localSheetId)}"`;
  if (n.hidden) xml += ' hidden="1"';
  return `${xml}>${escapeText(encodeXstring(n.refersTo))}</definedName>`;
}
