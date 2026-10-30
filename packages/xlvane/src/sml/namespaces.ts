// Namespace and relationship type URIs (docs/facts/ecma-376.md E-01 … E-05).
// Strict values come from ECMA-376 Part 1 (sml.xsd targetNamespace and the part summaries in
// §12.3); Transitional values from ECMA-376 Part 4 (transitional sml.xsd and §10.2).
// Microsoft extensions come from [MS-OFFMACRO2] (docs/facts/zip-xml-web.md W-08 … W-10).
import type { Conformance } from '../model';

/** SpreadsheetML main namespace. */
export const SML_NS: Readonly<Record<Conformance, string>> = {
  strict: 'http://purl.oclc.org/ooxml/spreadsheetml/main',
  transitional: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
};

/** Namespace of `r:id` and similar relationship-reference attributes. */
export const REL_NS: Readonly<Record<Conformance, string>> = {
  strict: 'http://purl.oclc.org/ooxml/officeDocument/relationships',
  transitional: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
};

/** OPC namespaces (ECMA-376 Part 2 schemas opc-contentTypes.xsd and opc-relationships.xsd). */
export const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
export const PACKAGE_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

const STRICT_REL = 'http://purl.oclc.org/ooxml/officeDocument/relationships/';
const TRANSITIONAL_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

/** Relationship type names used by the reader. */
export type RelationshipName =
  | 'officeDocument'
  | 'worksheet'
  | 'chartsheet'
  | 'dialogsheet'
  | 'sharedStrings'
  | 'styles'
  | 'externalLink';

/** Full relationship type URIs for a name, in both conformance classes. */
export function relationshipTypes(name: RelationshipName): readonly [strict: string, transitional: string] {
  return [STRICT_REL + name, TRANSITIONAL_REL + name];
}

/** Returns the relationship name and conformance class of a relationship type URI, if known. */
export function classifyRelationship(type: string): { name: string; conformance: Conformance } | null {
  if (type.startsWith(STRICT_REL)) return { name: type.slice(STRICT_REL.length), conformance: 'strict' };
  if (type.startsWith(TRANSITIONAL_REL)) return { name: type.slice(TRANSITIONAL_REL.length), conformance: 'transitional' };
  return null;
}

/** [MS-OFFMACRO2] relationship types for macros. */
export const VBA_PROJECT_REL = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject';
export const MACROSHEET_REL = 'http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet';
export const INTL_MACROSHEET_REL = 'http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet';

/** [MS-OFFMACRO2] content types for macros. */
export const VBA_PROJECT_CONTENT_TYPE = 'application/vnd.ms-office.vbaProject';
export const MACROSHEET_CONTENT_TYPES: readonly string[] = [
  'application/vnd.ms-excel.macrosheet+xml',
  'application/vnd.ms-excel.intlmacrosheet+xml',
];
