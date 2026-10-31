// One XLV_E_SHEET_NOT_FOUND error for every API that looks sheets up by name (D-R03).
import { XlvError } from '../errors';
import { suggest } from './suggest';

/**
 * The `XLV_E_SHEET_NOT_FOUND` error for a sheet name, shared by every API that looks sheets up by
 * name: the closest names, a note when only the case differs, and the available names.
 */
export function sheetNotFound(name: unknown, names: readonly string[]): XlvError {
  const shown = String(name).slice(0, 100);
  const sameIgnoringCase = names.filter((n) => n.toLowerCase() === shown.toLowerCase());
  const close = sameIgnoringCase.length > 0 ? sameIgnoringCase : suggest(shown, names);
  const listed = names.slice(0, 50).map((n) => `"${n}"`).join(', ');
  return new XlvError('XLV_E_SHEET_NOT_FOUND', `The workbook has no sheet named "${shown}".`, {
    location: { sheet: shown },
    hint:
      (close.length > 0 ? `Did you mean ${close.map((n) => `"${n}"`).join(' or ')}? ` : '') +
      `Available sheets: ${listed || '(none)'}${names.length > 50 ? ', …' : ''}. ` +
      (sameIgnoringCase.length > 0 ? 'Sheet names are case-sensitive here: write the name exactly as shown.' : 'Sheet names are case-sensitive.'),
    details: { name: shown, available: [...names] },
  });
}
