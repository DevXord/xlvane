// Ported from exceljs v4.4.0 spec/utils/test-data-validation-sheet.js,
// spec/utils/test-conditional-formatting-sheet.js, spec/utils/data/data-validations.json and
// spec/utils/data/conditional-formatting.json (MIT, Copyright (c) 2014-2019 Guyon Roche;
// tests/exceljs/LICENSE-exceljs.md), changes: Chai → Vitest `expect`; the JSON data inlined and
// `tools.fix` applied by hand (ISO strings → Date); exported as a TypeScript module. Data that
// xlvane deliberately does not write as exceljs does is changed in place and marked "xlvane:"
// (docs/exceljs-compat.md).
/* eslint-disable */
import { expect } from 'vitest';

const concatenateFormula = (...args: string[]) => ({ formula: `CONCATENATE(${args.map((v) => `"${v}"`).join(',')})` });

export const dataValidations: Record<string, any> = {
  B1: { type: 'list', allowBlank: true, formulae: ['Nephews'] },
  B3: { type: 'list', allowBlank: true, showInputMessage: true, showErrorMessage: true, formulae: ['"One,Two,Three,Four"'] },
  B5: { type: 'list', allowBlank: true, showInputMessage: true, showErrorMessage: true, formulae: ['$D$5:$F$5'] },
  B13: { type: 'whole', operator: 'equal', allowBlank: true, showInputMessage: true, showErrorMessage: true, formulae: [5], promptTitle: 'Five', prompt: 'The value must be Five' },
  // xlvane: errorStyle "error" is not a value of ST_DataValidationErrorStyle (stop, warning,
  // information; SR-03), and xlvane refuses it; "warning" is used instead.
  E13: { type: 'whole', operator: 'notEqual', allowBlank: true, showInputMessage: true, showErrorMessage: true, formulae: [5], errorStyle: 'warning', errorTitle: 'Five', error: 'The value must not be Five' },
  B15: { type: 'whole', operator: 'notEqual', formulae: [5] },
  // xlvane: a decimal validation needs its bound ([MS-XLSX] §2.6.5, D-C01); exceljs writes it
  // without one. The bounds 0 and 1 (between, the default operator) are added.
  B17: { type: 'decimal', operator: 'between', allowBlank: true, showInputMessage: true, promptTitle: 'Decimal', prompt: 'Value must be a decimal', formulae: [0, 1] },
  B19: { type: 'any', showInputMessage: true, promptTitle: 'Any', prompt: 'You can put any value here' },
  shared: { type: 'whole', operator: 'greaterThanOrEqual', formulae: [10] },
  types: ['whole', 'decimal', 'date', 'textLength'],
  values: {
    whole: { v1: 1, v2: 10 },
    decimal: { v1: 1.5, v2: 10.2 },
    date: { v1: new Date('2015-01-01T00:00:00.000Z'), v2: new Date('2016-01-01T00:00:00.000Z') },
    textLength: { v1: 5, v2: 15 },
  },
  operators: ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'],
};

function createDataValidations(type: string, operator: string) {
  const dataValidation: any = {
    type,
    operator,
    allowBlank: true,
    showInputMessage: true,
    showErrorMessage: true,
    formulae: [dataValidations.values[type].v1],
  };
  if (operator === 'between' || operator === 'notBetween') dataValidation.formulae.push(dataValidations.values[type].v2);
  return dataValidation;
}

export const dataValidationSheet = {
  addSheet(wb: any) {
    const ws = wb.addWorksheet('data-validations');

    // named list
    ws.getCell('D1').value = 'Hewie';
    ws.getCell('D1').name = 'Nephews';
    ws.getCell('E1').value = 'Dewie';
    ws.getCell('E1').name = 'Nephews';
    ws.getCell('F1').value = 'Louie';
    ws.getCell('F1').name = 'Nephews';
    ws.getCell('A1').value = concatenateFormula('Named List');
    ws.getCell('B1').dataValidation = dataValidations.B1;

    ws.getCell('A3').value = concatenateFormula('Literal List');
    ws.getCell('B3').dataValidation = dataValidations.B3;

    ws.getCell('D5').value = 'Tom';
    ws.getCell('E5').value = 'Dick';
    ws.getCell('F5').value = 'Harry';
    ws.getCell('A5').value = concatenateFormula('Range List');
    ws.getCell('B5').dataValidation = dataValidations.B5;

    dataValidations.operators.forEach((operator: string, cIndex: number) => {
      ws.getCell(7, 3 + cIndex).value = concatenateFormula(operator);
    });
    dataValidations.types.forEach((type: string, rIndex: number) => {
      const row = 8 + rIndex;
      ws.getCell(row, 1).value = concatenateFormula(type);
      dataValidations.operators.forEach((operator: string, cIndex: number) => {
        ws.getCell(row, 3 + cIndex).dataValidation = createDataValidations(type, operator);
      });
    });

    ws.getCell('A13').value = concatenateFormula('Prompt');
    ws.getCell('B13').dataValidation = dataValidations.B13;

    ws.getCell('D13').value = concatenateFormula('Error');
    ws.getCell('E13').dataValidation = dataValidations.E13;

    ws.getCell('A15').value = concatenateFormula('Terse');
    ws.getCell('B15').dataValidation = dataValidations.B15;

    ws.getCell('A17').value = concatenateFormula('Decimal');
    ws.getCell('B17').dataValidation = dataValidations.B17;

    ws.getCell('A19').value = concatenateFormula('Any');
    ws.getCell('B19').dataValidation = dataValidations.B19;

    ws.getCell('A20').value = new Date();
    ws.getCell('A20').dataValidation = {
      type: 'date',
      operator: 'greaterThan',
      showErrorMessage: true,
      allowBlank: true,
      formulae: [new Date(2016, 0, 1)],
    };

    // two rows of the same validation to test dataValidation optimisation
    ['A22', 'A23'].forEach((address) => {
      ws.getCell(address).value = concatenateFormula('Five Numbers');
    });
    ['B22', 'C22', 'D22', 'E22', 'F22', 'B23', 'C23', 'D23', 'E23', 'F23'].forEach((address) => {
      ws.getCell(address).dataValidation = JSON.parse(JSON.stringify(dataValidations.shared));
    });
  },

  checkSheet(wb: any) {
    const ws = wb.getWorksheet('data-validations');
    expect(ws).toBeDefined();

    expect(ws.getCell('B1').dataValidation).toEqual(dataValidations.B1);
    expect(ws.getCell('B3').dataValidation).toEqual(dataValidations.B3);
    expect(ws.getCell('B5').dataValidation).toEqual(dataValidations.B5);

    dataValidations.types.forEach((type: string, rIndex: number) => {
      const row = 8 + rIndex;
      dataValidations.operators.forEach((operator: string, cIndex: number) => {
        expect(ws.getCell(row, 3 + cIndex).dataValidation).toEqual(createDataValidations(type, operator));
      });
    });

    expect(ws.getCell('B13').dataValidation).toEqual(dataValidations.B13);
    expect(ws.getCell('E13').dataValidation).toEqual(dataValidations.E13);
    expect(ws.getCell('B15').dataValidation).toEqual(dataValidations.B15);
    expect(ws.getCell('B17').dataValidation).toEqual(dataValidations.B17);
    expect(ws.getCell('B19').dataValidation).toEqual(dataValidations.B19);

    // two rows of the same validation to test dataValidation optimisation
    ['B22', 'C22', 'D22', 'E22', 'F22', 'B23', 'C23', 'D23', 'E23', 'F23'].forEach((address) => {
      expect(ws.getCell(address).dataValidation).toEqual(dataValidations.shared);
    });
  },
};

export const conditionalFormattings: Record<string, any> = {
  expression: {
    ref: 'A1:E7',
    rules: [
      { type: 'expression', formulae: ['MOD(ROW()+COLUMN(),2)=0'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FF00FF00' } } } },
      { type: 'expression', priority: 2, formulae: ['TRUE'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFF0000' } } } },
    ],
  },
  top10: {
    ref: 'A8:E15',
    rules: [
      { type: 'top10', percent: true, rank: 10, style: { font: { bold: true } } },
      { type: 'top10', percent: true, bottom: true, rank: 10, style: { font: { italic: true } } },
    ],
  },
  abbreviation: {
    ref: 'A:A',
    rules: [
      { type: 'cellIs', operator: 'between', formulae: [1000, 1000000], style: { numFmt: '#,##0.000,\\K;-#,##0.000,\\K' } },
      { type: 'cellIs', operator: 'between', formulae: [-1000, -1000000], style: { numFmt: '#,##0.000,\\K;-#,##0.000,\\K' } },
      { type: 'cellIs', operator: 'greaterThan', formulae: [1000000], style: { numFmt: '#,##0.000,,\\M;-#,##0.000,,\\M' } },
      { type: 'cellIs', operator: 'lessThan', formulae: [-1000000], style: { numFmt: '#,##0.000,,\\M;-#,##0.000,,\\M' } },
    ],
  },
  types: ['expression', 'cellIs', 'top10', 'aboveAverage', 'colorScale', 'iconSet', 'containsText', 'timePeriod'],
};

export const conditionalFormattingSheet = {
  addSheet(wb: any) {
    const ws = wb.addWorksheet('conditional-formatting');
    conditionalFormattings.types.forEach((type: string) => {
      const cf = conditionalFormattings[type];
      if (cf) ws.addConditionalFormatting(cf);
    });
  },

  checkSheet(wb: any) {
    const ws = wb.getWorksheet('conditional-formatting');
    expect(ws).toBeDefined();
    expect(ws.conditionalFormattings).toBeDefined();
    ws.conditionalFormattings.forEach((item: any) => {
      const type = item.rules && item.rules[0].type;
      const cf = conditionalFormattings[type];
      expect(item).toHaveProperty('ref');
      expect(item).toHaveProperty('rules');
      expect(cf).toHaveProperty('ref');
      expect(cf).toHaveProperty('rules');
      expect(item.ref).toEqual(cf.ref);
      expect(item.rules.length).toBe(cf.rules.length);
    });
  },
};
