const {readPart} = require('../../utils/xlsx-tamper');

const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  it('issue 2736 - conditional formatting stopIfTrue is written and read back', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell('A1').value = 1;
    const style = {font: {bold: true}};
    const rules = [
      {type: 'expression', formulae: ['A1>0'], style, stopIfTrue: true},
      {type: 'cellIs', operator: 'equal', formulae: ['1'], style, stopIfTrue: true},
      {type: 'top10', rank: 3, style, stopIfTrue: true},
      {type: 'aboveAverage', style, stopIfTrue: true},
      {type: 'containsText', operator: 'containsText', text: 'x', style, stopIfTrue: true},
      {type: 'timePeriod', timePeriod: 'today', style, stopIfTrue: true},
      {type: 'expression', formulae: ['A1<0'], style},
    ];
    ws.addConditionalFormatting({ref: 'A1:A3', rules});

    const buffer = await wb.xlsx.writeBuffer();
    const xml = await readPart(buffer, 'xl/worksheets/sheet1.xml');
    expect(xml.match(/stopIfTrue="1"/g)).to.have.length(6);

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(buffer);
    const readRules = wb2.getWorksheet('Sheet1').conditionalFormattings[0].rules;
    expect(readRules.map(rule => rule.stopIfTrue)).to.deep.equal([
      true,
      true,
      true,
      true,
      true,
      true,
      undefined,
    ]);
  });
});
