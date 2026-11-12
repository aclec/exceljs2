// Reading xlsx files written by other producers (Excel, LibreOffice, openpyxl, POI, hand-edited):
// legal-but-unusual XML must load, malformed parts must not crash with a TypeError.
// Fixtures are built with ExcelJS and patched with xlsx-tamper so they stay readable.
const {baseWorkbookBuffer, tamper} = require('../../utils/xlsx-tamper');

const ExcelJS = verquire('exceljs');

const SHEET = 'xl/worksheets/sheet1.xml';

const replaceSheetData = sheetData => xml =>
  xml.replace(/<sheetData>.*<\/sheetData>|<sheetData\/>/s, `<sheetData>${sheetData}</sheetData>`);

async function load(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  return wb;
}

describe('xlsx read robustness', function() {
  this.timeout(20000);

  describe('rows and cells without an r attribute', () => {
    it('infers row numbers and cell addresses from their position', async () => {
      const buffer = await tamper(await baseWorkbookBuffer(), {
        [SHEET]: replaceSheetData(
          '<row><c t="inlineStr"><is><t>a1</t></is></c><c><v>2</v></c></row>' +
            '<row r="4"><c r="C4"><v>3</v></c><c><v>4</v></c></row>' +
            '<row><c><v>5</v></c></row>'
        ),
      });
      const ws = (await load(buffer)).getWorksheet('Sheet1');
      expect(ws.getCell('A1').value).to.equal('a1');
      expect(ws.getCell('B1').value).to.equal(2);
      expect(ws.getCell('C4').value).to.equal(3);
      expect(ws.getCell('D4').value).to.equal(4);
      expect(ws.getCell('A5').value).to.equal(5);
      expect(ws.rowCount).to.equal(5);
    });

    it('takes the row number from the cells when only the row omits r', async () => {
      const buffer = await tamper(await baseWorkbookBuffer(), {
        [SHEET]: replaceSheetData('<row><c r="B3"><v>7</v></c></row>'),
      });
      const ws = (await load(buffer)).getWorksheet('Sheet1');
      expect(ws.getCell('B3').value).to.equal(7);
      expect(ws.rowCount).to.equal(3);
    });
  });
});
