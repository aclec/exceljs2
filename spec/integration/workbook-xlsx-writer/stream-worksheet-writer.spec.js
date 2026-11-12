// Streaming writer (ExcelJS.stream.xlsx.WorkbookWriter): sheet XML structure and shared strings.
const {PassThrough} = require('stream');
const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

const IMAGE_FILENAME = `${__dirname}/../data/image.png`;

// commit the workbook into memory and return the xlsx buffer
async function writeBuffer(build, options = {}) {
  const stream = new PassThrough();
  const chunks = [];
  stream.on('data', chunk => chunks.push(chunk));
  const wb = new ExcelJS.stream.xlsx.WorkbookWriter({stream, ...options});
  await build(wb);
  await wb.commit();
  return Buffer.concat(chunks);
}

async function readPart(buffer, path) {
  const zip = await JSZip.loadAsync(buffer);
  return zip.file(path).async('string');
}

describe('WorkbookWriter (stream)', () => {
  it('writes worksheet elements in CT_Worksheet order', async () => {
    const buffer = await writeBuffer(wb => {
      const ws = wb.addWorksheet('Sheet', {
        properties: {tabColor: {argb: 'FFFF0000'}},
        views: [{state: 'frozen', ySplit: 1}],
        pageSetup: {showGridLines: true, orientation: 'landscape'},
        headerFooter: {oddHeader: 'header'},
      });
      ws.autoFilter = 'A1:B1';
      ws.mergeCells('C1:D1');
      ws.getCell('A1').value = 'text';
      ws.getCell('A1').note = 'a note';
      ws.getCell('B1').value = {text: 'link', hyperlink: 'https://example.com'};
      ws.getRow(1).addPageBreak();
      ws.addConditionalFormatting({
        ref: 'A2:A3',
        rules: [{type: 'dataBar', cfvo: [{type: 'min'}, {type: 'max'}], color: {argb: 'FF638EC6'}}],
      });
      ws.dataValidations.add('A2', {type: 'whole', operator: 'between', formulae: [1, 9]});
      ws.addBackgroundImage(wb.addImage({filename: IMAGE_FILENAME, extension: 'png'}));
      ws.getCell('A2').value = 1;
      ws.getCell('A3').value = 2;
      return ws.protect('', {selectLockedCells: true});
    });
    const xml = await readPart(buffer, 'xl/worksheets/sheet1.xml');

    // ECMA-376 CT_Worksheet sequence (elements the stream writer can produce)
    const order = [
      'sheetPr',
      'sheetViews',
      'sheetFormatPr',
      'sheetData',
      'sheetProtection',
      'autoFilter',
      'mergeCells',
      'conditionalFormatting',
      'dataValidations',
      'hyperlinks',
      'printOptions',
      'pageMargins',
      'pageSetup',
      'headerFooter',
      'rowBreaks',
      'legacyDrawing',
      'picture',
      'extLst',
    ];
    // last occurrence: a dataBar cfRule has its own nested extLst
    const positions = order.map(tag => ({
      tag,
      index: Math.max(xml.lastIndexOf(`<${tag} `), xml.lastIndexOf(`<${tag}>`), xml.lastIndexOf(`<${tag}/>`)),
    }));
    expect(positions.filter(({index}) => index < 0).map(({tag}) => tag)).to.deep.equal([]);
    const sorted = [...positions].sort((a, b) => a.index - b.index).map(({tag}) => tag);
    expect(sorted).to.deep.equal(order);

    // the dataBar cfRule references its x14 extension by id
    const ids = [...xml.matchAll(/<x14:id>([^<]+)<\/x14:id>|<x14:cfRule [^>]*id="([^"]+)"/g)];
    expect(ids).to.have.length(2);
    expect(ids[0][1]).to.be.a('string');
    expect(ids[0][1]).to.equal(ids[1][2]);

    // and the document reader still understands it
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(buffer);
    const ws2 = wb2.getWorksheet('Sheet');
    expect(ws2.pageSetup.showGridLines).to.equal(true);
    expect(ws2.conditionalFormattings[0].rules[0].type).to.equal('dataBar');
    expect(ws2.getCell('B1').hyperlink).to.equal('https://example.com');
  });
});
