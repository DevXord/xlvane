// Defined names and document properties come from untrusted files: names such as "__proto__"
// stay data, and no prototype gains a property (CLAUDE.md: no prototype pollution).
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectPrototypesUntouched } from '../helpers/prototypes';
import { buildXlsx } from '../helpers/xlsx-builder';

const CP = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';

describe('workbook metadata from untrusted files', () => {
  it('defined names and core properties named like prototype keys stay data', () => {
    const bytes = buildXlsx({
      workbookAfterSheets:
        '<definedNames><definedName name="__proto__">polluted</definedName><definedName name="constructor" comment="x">prototype</definedName></definedNames>',
      rootRels: [['http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml']],
      extraParts: {
        'docProps/core.xml': `<coreProperties xmlns="${CP}" xmlns:dc="http://purl.org/dc/elements/1.1/"><__proto__>x</__proto__><polluted>yes</polluted><dc:title>t</dc:title></coreProperties>`,
      },
    });
    const wb = readWorkbook(bytes);
    expect(wb.definedNames.map((n) => n.name)).toEqual(['__proto__', 'constructor']);
    expect(Object.keys(wb.properties)).toEqual(['title']);
    expect(Object.getPrototypeOf(wb.properties)).toBe(Object.prototype);
    expectPrototypesUntouched();
  });

  it('writing a name "__proto__" is allowed and harmless', () => {
    const wb = createWorkbook();
    wb.addSheet('S');
    wb.defineName('__proto__', '1');
    expect(readWorkbook(writeWorkbook(wb)).definedNames[0]?.name).toBe('__proto__');
    expectPrototypesUntouched();
  });
});
