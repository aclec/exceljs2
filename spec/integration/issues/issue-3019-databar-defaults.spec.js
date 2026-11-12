const {readPart} = require('../../utils/xlsx-tamper');

const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  it('issue 3019 - data bar without cfvo or color is written with defaults', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell('A1').value = 1;
    ws.addConditionalFormatting({ref: 'A1:A3', rules: [{type: 'dataBar'}]});

    const buffer = await wb.xlsx.writeBuffer();
    const xml = await readPart(buffer, 'xl/worksheets/sheet1.xml');
    // CT_DataBar requires exactly two cfvo and a color
    expect(xml).to.include('<dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>');
    expect(xml).to.include('<x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/>');

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(buffer);
    const [rule] = wb2.getWorksheet('Sheet1').conditionalFormattings[0].rules;
    expect(rule.type).to.equal('dataBar');
    expect(rule.cfvo).to.have.length(2);
  });
});
