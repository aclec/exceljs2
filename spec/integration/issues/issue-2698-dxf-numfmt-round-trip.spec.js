const {readPart} = require('../../utils/xlsx-tamper');

const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  it('issue 2698 - a dxf number format survives read and write', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell('A1').value = 0.5;
    ws.addConditionalFormatting({
      ref: 'A1',
      rules: [{type: 'expression', formulae: ['A1>0'], style: {numFmt: '0.000%'}}],
    });

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(await wb.xlsx.writeBuffer());
    const [rule] = wb2.getWorksheet('Sheet1').conditionalFormattings[0].rules;
    expect(rule.style.numFmt).to.equal('0.000%');

    const styles = await readPart(await wb2.xlsx.writeBuffer(), 'xl/styles.xml');
    expect(styles).not.to.include('[object Object]');
    expect(styles).to.match(/<dxf><numFmt numFmtId="\d+" formatCode="0.000%"\/><\/dxf>/);
  });
});
