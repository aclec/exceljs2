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

  describe('table autoFilter content that is not modelled', () => {
    const TABLE = 'xl/tables/table1.xml';
    const tableBuffer = () =>
      baseWorkbookBuffer((wb, ws) => {
        ws.addTable({
          name: 'T1',
          ref: 'A3',
          headerRow: true,
          columns: [
            {name: 'a', filterButton: true},
            {name: 'b', filterButton: true},
          ],
          rows: [[1, 2]],
        });
      });
    const withAutoFilterContent = content => xml =>
      xml.replace(/<autoFilter ref="([^"]+)">.*<\/autoFilter>/, `<autoFilter ref="$1">${content}</autoFilter>`);

    const filters = {
      dynamicFilter: '<dynamicFilter type="today"/>',
      top10: '<top10 top="1" val="5"/>',
      colorFilter: '<colorFilter dxfId="0"/>',
      iconFilter: '<iconFilter iconSet="3Arrows" iconId="1"/>',
      extLst:
        '<extLst><ext uri="{0}"><x14:filter val="x"/><customFilters><customFilter val="no"/></customFilters></ext></extLst>',
    };
    Object.entries(filters).forEach(([name, filter]) => {
      it(`ignores ${name} in a filterColumn`, async () => {
        const buffer = await tamper(await tableBuffer(), {
          [TABLE]: withAutoFilterContent(
            `<filterColumn colId="0" hiddenButton="0">${filter}</filterColumn><filterColumn colId="1" hiddenButton="0"/>`
          ),
        });
        const table = (await load(buffer)).getWorksheet('Sheet1').getTable('T1');
        expect(table.table.columns.map(column => column.name)).to.deep.equal(['a', 'b']);
        expect(table.table.columns[0].filterButton).to.equal(true);
        expect(table.table.columns[0].customFilters).to.equal(undefined);
      });
    });

    it('ignores sortState inside the autoFilter', async () => {
      const buffer = await tamper(await tableBuffer(), {
        [TABLE]: withAutoFilterContent(
          '<filterColumn colId="0" hiddenButton="0"/><sortState ref="A4:B4"><sortCondition ref="A4"/></sortState>'
        ),
      });
      const table = (await load(buffer)).getWorksheet('Sheet1').getTable('T1');
      expect(table.table.autoFilterRef).to.equal('A3:B4');
    });
  });

  describe('document properties', () => {
    it('loads when docProps/app.xml uses a prefixed root element', async () => {
      const buffer = await tamper(await baseWorkbookBuffer(), {
        'docProps/app.xml': () =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<ep:Properties xmlns:ep="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">' +
          '<ep:Application>Other</ep:Application></ep:Properties>',
      });
      const wb = await load(buffer);
      expect(wb.getWorksheet('Sheet1').getCell('A1').value).to.equal('hello');
      expect(wb.company).to.equal(undefined);
      expect(wb.manager).to.equal(undefined);
    });

    it('loads when docProps/app.xml has no root element', async () => {
      const buffer = await tamper(await baseWorkbookBuffer(), {
        'docProps/app.xml': () => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Other/>',
      });
      const wb = await load(buffer);
      expect(wb.getWorksheet('Sheet1').getCell('A1').value).to.equal('hello');
    });

    it('ignores core properties it does not model', async () => {
      const buffer = await tamper(await baseWorkbookBuffer(wb => {
        wb.creator = 'me';
      }), {
        'docProps/core.xml': xml =>
          xml.replace(
            '</cp:coreProperties>',
            '<dc:contributor>x</dc:contributor><dcterms:available><b>y</b></dcterms:available></cp:coreProperties>'
          ),
      });
      const wb = await load(buffer);
      expect(wb.creator).to.equal('me');
    });
  });
});
