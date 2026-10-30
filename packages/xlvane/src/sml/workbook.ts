// Parses the workbook part (ECMA-376 Part 1 §18.2.27 workbook, §18.2.19 sheet, §18.2.28
// workbookPr; docs/facts/ecma-376.md E-04, E-09 … E-11).
import { XlvError } from '../errors';
import type { ReadLimits } from '../limits';
import type { Conformance, SheetState } from '../model';
import type { WarningCollector } from '../util/warnings';
import { childNamed, childrenNamed, parseXmlTree } from '../xml/tree';
import { REL_NS, SML_NS } from './namespaces';
import { decodeXstring } from './xstring';

/** One `sheet` element of the workbook part. */
export interface WorkbookSheetEntry {
  readonly name: string;
  readonly sheetId: number | null;
  readonly state: SheetState;
  /** Relationship id (`r:id`) of the sheet part; `null` if the attribute is missing. */
  readonly relId: string | null;
}

/** One `definedName` element; `localSheetId` is resolved to a sheet by the caller. */
export interface WorkbookDefinedName {
  readonly name: string;
  readonly refersTo: string;
  readonly localSheetId: number | null;
  readonly hidden: boolean;
  readonly comment: string | null;
}

/** What the reader needs from the workbook part. */
export interface WorkbookInfo {
  readonly conformance: Conformance;
  readonly date1904: boolean;
  readonly sheets: readonly WorkbookSheetEntry[];
  readonly definedNames: readonly WorkbookDefinedName[];
}

/** Context for parsing one part. */
export interface WorkbookContext {
  readonly part: string;
  readonly limits: ReadLimits;
  readonly warnings: WarningCollector;
}

const SHEET_STATES: readonly SheetState[] = ['visible', 'hidden', 'veryHidden'];

/** Parses an `xsd:boolean` lexical value; `null` if it is not one. */
export function parseXsdBoolean(value: string): boolean | null {
  const v = value.trim();
  if (v === 'true' || v === '1') return true;
  if (v === 'false' || v === '0') return false;
  return null;
}

/** Parses the workbook part. */
export function parseWorkbookPart(xml: string, ctx: WorkbookContext): WorkbookInfo {
  const root = parseXmlTree(xml, ctx);
  let conformance: Conformance;
  if (root.local === 'workbook' && root.ns === SML_NS.strict) conformance = 'strict';
  else if (root.local === 'workbook' && root.ns === SML_NS.transitional) conformance = 'transitional';
  else {
    throw new XlvError('XLV_E_PART_INVALID', `The workbook part has the root element <${root.local}> in namespace "${root.ns}".`, {
      location: { part: ctx.part },
      hint: 'The main part of the package is not a SpreadsheetML workbook. Check that the file is an Excel workbook (.xlsx/.xlsm), not a Word or PowerPoint file.',
      details: { element: root.local, namespace: root.ns },
    });
  }
  const ns = root.ns;
  const relIdKey = `{${REL_NS[conformance]}}id`;

  let date1904 = false;
  const pr = childNamed(root, ns, 'workbookPr');
  const rawDate1904 = pr?.attrs.get('date1904');
  if (rawDate1904 !== undefined) {
    const parsed = parseXsdBoolean(rawDate1904);
    if (parsed === null) {
      ctx.warnings.add('XLV_W_INVALID_VALUE', `workbookPr@date1904 has the invalid value "${rawDate1904}"; using the 1900 date system.`, {
        part: ctx.part,
      });
    } else {
      date1904 = parsed;
    }
  }

  const sheetsElement = childNamed(root, ns, 'sheets');
  if (sheetsElement === undefined) {
    throw new XlvError('XLV_E_PART_INVALID', 'The workbook part has no <sheets> element.', {
      location: { part: ctx.part },
      hint: 'Every workbook must list its sheets. Re-save the file in a spreadsheet application.',
    });
  }
  const sheets = childrenNamed(sheetsElement, ns, 'sheet').map((el): WorkbookSheetEntry => {
    const name = decodeXstring(el.attrs.get('name') ?? '');
    const rawState = el.attrs.get('state') ?? 'visible';
    let state = SHEET_STATES.find((s) => s === rawState);
    if (state === undefined) {
      ctx.warnings.add('XLV_W_INVALID_VALUE', `Sheet "${name}" has the unknown state "${rawState}"; treating it as visible.`, {
        part: ctx.part,
        sheet: name,
      });
      state = 'visible';
    }
    const rawId = el.attrs.get('sheetId');
    const sheetId = rawId !== undefined && /^\d{1,10}$/.test(rawId) ? Number(rawId) : null;
    return { name, sheetId, state, relId: el.attrs.get(relIdKey) ?? null };
  });
  const namesElement = childNamed(root, ns, 'definedNames');
  const definedNames = (namesElement === undefined ? [] : childrenNamed(namesElement, ns, 'definedName')).map((el): WorkbookDefinedName => {
    const name = decodeXstring(el.attrs.get('name') ?? '');
    const rawId = el.attrs.get('localSheetId');
    let localSheetId: number | null = null;
    if (rawId !== undefined) {
      if (/^\d{1,10}$/.test(rawId.trim())) localSheetId = Number(rawId.trim());
      else ctx.warnings.add('XLV_W_INVALID_VALUE', `The defined name "${name.slice(0, 100)}" has the invalid localSheetId "${rawId.slice(0, 20)}"; it is read as a workbook-level name.`, { part: ctx.part });
    }
    const rawHidden = el.attrs.get('hidden');
    const comment = el.attrs.get('comment');
    return {
      name,
      refersTo: decodeXstring(el.text),
      localSheetId,
      hidden: rawHidden !== undefined && parseXsdBoolean(rawHidden) === true,
      comment: comment === undefined ? null : decodeXstring(comment),
    };
  });
  return { conformance, date1904, sheets, definedNames };
}
