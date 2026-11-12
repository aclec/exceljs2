const {baseWorkbookBuffer, tamper, readPart} = require('../../utils/xlsx-tamper');

const ExcelJS = verquire('exceljs');

const definedNames = xml => xml.match(/<definedName [^>]*>[^<]*<\/definedName>/g) || [];

async function load(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  return wb;
}

describe('github issues', () => {
  describe('issue 1516 - multiple print areas', () => {
    it('writes one _xlnm.Print_Area per sheet and reads every range back', async () => {
      const buffer = await baseWorkbookBuffer((wb, ws) => {
        ws.pageSetup.printArea = 'A1:B2&&D4:E5';
      });
      const names = definedNames(await readPart(buffer, 'xl/workbook.xml'));
      expect(names).to.deep.equal([
        '<definedName name="_xlnm.Print_Area" localSheetId="0">' +
          '&apos;Sheet1&apos;!$A1:$B2,&apos;Sheet1&apos;!$D4:$E5</definedName>',
      ]);

      const wb = await load(buffer);
      expect(wb.getWorksheet('Sheet1').pageSetup.printArea).to.equal('A1:B2&&D4:E5');
    });

    it('reads a defined name with several ranges and legacy duplicate names', async () => {
      const buffer = await tamper(await baseWorkbookBuffer(), {
        'xl/workbook.xml': xml =>
          xml.replace(
            '</sheets>',
            '</sheets><definedNames>' +
              '<definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$B$2,Sheet1!$D$4:$E$5</definedName>' +
              '<definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$G$1:$H$2</definedName>' +
              '</definedNames>'
          ),
      });
      const wb = await load(buffer);
      expect(wb.getWorksheet('Sheet1').pageSetup.printArea).to.equal('A1:B2&&D4:E5&&G1:H2');

      const names = definedNames(await readPart(await wb.xlsx.writeBuffer(), 'xl/workbook.xml'));
      expect(names).to.have.length(1);
    });
  });
});
