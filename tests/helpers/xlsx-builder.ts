// Builds small XLSX packages from hand-written XML for tests. Part names, content types,
// namespaces and relationship types follow ECMA-376 (docs/facts/ecma-376.md E-01 … E-07).
import { zipSync } from 'fflate';

const enc = new TextEncoder();

export const NS = {
  transitional: {
    main: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/',
  },
  strict: {
    main: 'http://purl.oclc.org/ooxml/spreadsheetml/main',
    r: 'http://purl.oclc.org/ooxml/officeDocument/relationships',
    rel: 'http://purl.oclc.org/ooxml/officeDocument/relationships/',
  },
} as const;

export const CT = {
  workbook: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  template: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml',
  xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
  worksheet: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
  chartsheet: 'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml',
  sharedStrings: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
  styles: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
  externalLink: 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml',
  relationships: 'application/vnd.openxmlformats-package.relationships+xml',
  vba: 'application/vnd.ms-office.vbaProject',
} as const;

export interface SheetSpec {
  name: string;
  /** Contents of <sheetData>, or a complete worksheet part when it starts with "<?xml" or "<worksheet". */
  data?: string;
  /** Extra elements placed after </sheetData> (e.g. "<mergeCells>…</mergeCells>"). */
  after?: string;
  /** Extra elements placed before <sheetData> (e.g. '<dimension ref="A1:B2"/>'). */
  before?: string;
  state?: string;
  kind?: 'worksheet' | 'chartsheet';
  /** Leaves the sheet's part out of the package. */
  missingPart?: boolean;
}

export interface XlsxSpec {
  sheets?: SheetSpec[];
  /** Shared strings as plain strings (escaped automatically) or a complete part. */
  sharedStrings?: string[] | string;
  /** Contents of <styleSheet> or a complete part. */
  styles?: string;
  /** Attributes for <workbookPr>, e.g. 'date1904="1"'. */
  workbookPr?: string;
  strict?: boolean;
  /** Content type of the workbook part. */
  workbookContentType?: string;
  /** Adds a VBA project part and relationship. */
  vba?: boolean;
  /** Number of external link parts to add. */
  externalLinks?: number;
  /** Additional raw parts (name without leading slash). */
  extraParts?: Record<string, string | Uint8Array>;
  /** Parts to leave out entirely (name without leading slash). */
  omit?: string[];
  /** Raw XML placed in the workbook part after </sheets> (e.g. "<definedNames>…</definedNames>"). */
  workbookAfterSheets?: string;
  /** Extra package relationships ([type, target]), with ids rId2, rId3, … */
  rootRels?: [string, string][];
}

export function escapeXml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

export function buildXlsx(spec: XlsxSpec = {}): Uint8Array {
  const ns = spec.strict ? NS.strict : NS.transitional;
  const sheets = spec.sheets ?? [{ name: 'Sheet1', data: '' }];
  const files: Record<string, Uint8Array> = {};
  const put = (name: string, content: string | Uint8Array): void => {
    files[name] = typeof content === 'string' ? enc.encode(content) : content;
  };

  const overrides: [string, string][] = [['/xl/workbook.xml', spec.workbookContentType ?? (spec.vba ? CT.xlsm : CT.workbook)]];
  const wbRels: [string, string, string][] = [];

  sheets.forEach((sheet, i) => {
    const kind = sheet.kind ?? 'worksheet';
    const path = kind === 'worksheet' ? `worksheets/sheet${i + 1}.xml` : `chartsheets/sheet${i + 1}.xml`;
    wbRels.push([`rId${i + 1}`, ns.rel + kind, path]);
    if (sheet.missingPart) return;
    overrides.push([`/xl/${path}`, kind === 'worksheet' ? CT.worksheet : CT.chartsheet]);
    const data = sheet.data ?? '';
    if (kind === 'chartsheet') {
      put(`xl/${path}`, `${DECL}<chartsheet xmlns="${ns.main}" xmlns:r="${ns.r}"><sheetView/></chartsheet>`);
    } else if (data.startsWith('<?xml') || data.startsWith('<worksheet')) {
      put(`xl/${path}`, data);
    } else {
      put(
        `xl/${path}`,
        `${DECL}<worksheet xmlns="${ns.main}" xmlns:r="${ns.r}">${sheet.before ?? ''}<sheetData>${data}</sheetData>${sheet.after ?? ''}</worksheet>`,
      );
    }
  });

  let nextRel = sheets.length + 1;
  if (spec.sharedStrings !== undefined) {
    wbRels.push([`rId${nextRel++}`, `${ns.rel}sharedStrings`, 'sharedStrings.xml']);
    overrides.push(['/xl/sharedStrings.xml', CT.sharedStrings]);
    const sst =
      typeof spec.sharedStrings === 'string'
        ? spec.sharedStrings
        : `${DECL}<sst xmlns="${ns.main}" count="${spec.sharedStrings.length}" uniqueCount="${spec.sharedStrings.length}">${spec.sharedStrings
            .map((s) => `<si><t xml:space="preserve">${escapeXml(s)}</t></si>`)
            .join('')}</sst>`;
    put('xl/sharedStrings.xml', sst);
  }
  if (spec.styles !== undefined) {
    wbRels.push([`rId${nextRel++}`, `${ns.rel}styles`, 'styles.xml']);
    overrides.push(['/xl/styles.xml', CT.styles]);
    put('xl/styles.xml', spec.styles.startsWith('<?xml') ? spec.styles : `${DECL}<styleSheet xmlns="${ns.main}">${spec.styles}</styleSheet>`);
  }
  if (spec.vba) {
    wbRels.push([`rId${nextRel++}`, 'http://schemas.microsoft.com/office/2006/relationships/vbaProject', 'vbaProject.bin']);
    overrides.push(['/xl/vbaProject.bin', CT.vba]);
    put('xl/vbaProject.bin', new Uint8Array([0xd0, 0xcf, 0x11, 0xe0]));
  }
  for (let i = 1; i <= (spec.externalLinks ?? 0); i++) {
    wbRels.push([`rId${nextRel++}`, `${ns.rel}externalLink`, `externalLinks/externalLink${i}.xml`]);
    overrides.push([`/xl/externalLinks/externalLink${i}.xml`, CT.externalLink]);
    put(`xl/externalLinks/externalLink${i}.xml`, `${DECL}<externalLink xmlns="${ns.main}"/>`);
  }

  const sheetElements = sheets
    .map(
      (s, i) =>
        `<sheet name="${escapeXml(s.name)}" sheetId="${i + 1}"${s.state ? ` state="${s.state}"` : ''} r:id="rId${i + 1}"/>`,
    )
    .join('');
  put(
    'xl/workbook.xml',
    `${DECL}<workbook xmlns="${ns.main}" xmlns:r="${ns.r}"${spec.strict ? ' conformance="strict"' : ''}>` +
      `<workbookPr ${spec.workbookPr ?? ''}/><sheets>${sheetElements}</sheets>${spec.workbookAfterSheets ?? ''}</workbook>`,
  );
  put(
    'xl/_rels/workbook.xml.rels',
    `${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${wbRels
      .map(([id, type, target]) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`)
      .join('')}</Relationships>`,
  );
  put(
    '_rels/.rels',
    `${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${ns.rel}officeDocument" Target="xl/workbook.xml"/>` +
      (spec.rootRels ?? []).map(([type, target], i) => `<Relationship Id="rId${i + 2}" Type="${type}" Target="${target}"/>`).join('') +
      `</Relationships>`,
  );
  put(
    '[Content_Types].xml',
    `${DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="${CT.relationships}"/><Default Extension="xml" ContentType="application/xml"/>` +
      overrides.map(([part, type]) => `<Override PartName="${part}" ContentType="${type}"/>`).join('') +
      `</Types>`,
  );
  for (const [name, content] of Object.entries(spec.extraParts ?? {})) put(name, content);
  for (const name of spec.omit ?? []) delete files[name];
  return zipSync(files);
}

/** Shorthand for one cell element. */
export function c(ref: string, value: string | number | null, attrs = ''): string {
  return value === null ? `<c r="${ref}" ${attrs}/>` : `<c r="${ref}" ${attrs}><v>${String(value)}</v></c>`;
}

/** Shorthand for an inline string cell. */
export function inline(ref: string, text: string): string {
  return `<c r="${ref}" t="inlineStr"><is><t>${escapeXml(text)}</t></is></c>`;
}

/** Shorthand for a row. */
export function row(r: number, ...cells: string[]): string {
  return `<row r="${r}">${cells.join('')}</row>`;
}
