// Reading xlsx files written by other producers (Excel, LibreOffice, openpyxl, POI, hand-edited):
// legal-but-unusual XML must load, malformed parts must not crash with a TypeError.
// Fixtures are built with ExcelJS and patched with xlsx-tamper so they stay readable.
const JSZip = require('jszip');
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

  describe('worksheet relationships', () => {
    const RELS = 'xl/worksheets/_rels/sheet1.xml.rels';
    // 1x1 transparent png
    const PNG = Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
      'base64'
    );
    const richBuffer = () =>
      baseWorkbookBuffer((wb, ws) => {
        ws.getCell('B1').value = 'noted';
        ws.getCell('B1').note = {
          texts: [{text: 'a '}, {text: 'note'}],
          margins: {insetmode: 'custom', inset: [0.25, 0.25, 0.35, 0.35]},
        };
        ws.getCell('C1').value = {text: 'link', hyperlink: 'https://example.com/'};
        ws.addTable({name: 'T1', ref: 'A3', columns: [{name: 'a'}], rows: [[1]]});
        ws.addImage(wb.addImage({buffer: PNG, extension: 'png'}), 'E5:F6');
      });
    const setTarget = (type, target) => xml =>
      xml.replace(new RegExp(`(relationships/${type}" Target=")[^"]+"`), `$1${target}"`);
    const moveParts = async (buffer, moves) => {
      const zip = await JSZip.loadAsync(buffer);
      for (const [from, to] of Object.entries(moves)) {
        zip.file(to, await zip.file(from).async('nodebuffer'));
        zip.remove(from);
      }
      return zip.generateAsync({type: 'nodebuffer', compression: 'DEFLATE'});
    };
    const expectEverything = ws => {
      expect(ws.getCell('B1').note.texts).to.deep.equal([{text: 'a '}, {text: 'note'}]);
      expect(ws.getCell('B1').note.margins.inset).to.deep.equal([0.25, 0.25, 0.35, 0.35]);
      expect(ws.getCell('C1').hyperlink).to.equal('https://example.com/');
      expect(ws.getTable('T1').table.name).to.equal('T1');
      expect(ws.getImages()).to.have.length(1);
    };

    it('reads every part of the reference workbook', async () => {
      expectEverything((await load(await richBuffer())).getWorksheet('Sheet1'));
    });

    it('ignores relationships to parts that do not exist', async () => {
      const buffer = await tamper(await richBuffer(), {
        [RELS]: xml =>
          [
            setTarget('comments', '../missing.xml'),
            setTarget('vmlDrawing', '../drawings/missing.vml'),
            setTarget('drawing', '../drawings/missing.xml'),
            setTarget('table', '../tables/missing.xml'),
          ].reduce((x, patch) => patch(x), xml),
      });
      const ws = (await load(buffer)).getWorksheet('Sheet1');
      expect(ws.getCell('A1').value).to.equal('hello');
      expect(ws.getCell('B1').note).to.equal(undefined);
      expect(ws.getTables()).to.have.length(0);
      expect(ws.getImages()).to.have.length(0);
    });

    it('ignores references to relationships that do not exist', async () => {
      const buffer = await tamper(await richBuffer(), {
        [SHEET]: xml =>
          xml
            .replace('<hyperlink ref="C1" r:id="rId1"/>', '<hyperlink ref="C1" r:id="rId91"/>')
            .replace('<drawing r:id="rId4"/>', '<drawing r:id="rId94"/>')
            .replace('<tablePart r:id="rId5"/>', '<tablePart r:id="rId95"/>'),
      });
      const ws = (await load(buffer)).getWorksheet('Sheet1');
      expect(ws.getCell('C1').value).to.equal('link');
      expect(ws.getTables()).to.have.length(0);
      expect(ws.getImages()).to.have.length(0);
    });

    it('ignores a drawing part that is not a spreadsheet drawing', async () => {
      const buffer = await tamper(await richBuffer(), {
        'xl/drawings/drawing1.xml': () =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<c:userShapes xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>',
      });
      const ws = (await load(buffer)).getWorksheet('Sheet1');
      expect(ws.getCell('A1').value).to.equal('hello');
      expect(ws.getImages()).to.have.length(0);
    });

    it('resolves absolute relationship targets', async () => {
      const buffer = await tamper(await richBuffer(), {
        [RELS]: xml => xml.replace(/Target="\.\.\//g, 'Target="/xl/'),
      });
      expectEverything((await load(buffer)).getWorksheet('Sheet1'));
    });

    it('reads comments at the paths used by openpyxl', async () => {
      let buffer = await moveParts(await richBuffer(), {
        'xl/comments1.xml': 'xl/comments/comment1.xml',
        'xl/drawings/vmlDrawing1.vml': 'xl/drawings/commentsDrawing1.vml',
      });
      buffer = await tamper(buffer, {
        [RELS]: xml =>
          setTarget('vmlDrawing', '/xl/drawings/commentsDrawing1.vml')(
            setTarget('comments', '/xl/comments/comment1.xml')(xml)
          ),
      });
      expectEverything((await load(buffer)).getWorksheet('Sheet1'));
    });

    it('merges the vml note properties whatever the relationship order', async () => {
      const buffer = await tamper(await richBuffer(), {
        [RELS]: xml => {
          const comments = xml.match(/<Relationship [^>]*relationships\/comments"[^>]*\/>/)[0];
          return xml.replace(comments, '').replace('</Relationships>', `${comments}</Relationships>`);
        },
      });
      expectEverything((await load(buffer)).getWorksheet('Sheet1'));
    });
  });
});
