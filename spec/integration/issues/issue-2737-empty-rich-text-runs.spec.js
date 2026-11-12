const {PassThrough} = require('stream');
const {readPart} = require('../../utils/xlsx-tamper');

const ExcelJS = verquire('exceljs');

const VALUES = [
  {richText: [{font: {bold: true}, text: ''}, {text: 'kept'}, {text: ''}]},
  {richText: [{font: {bold: true}, text: ''}]},
];

function writeWorkbook() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  VALUES.forEach((value, i) => {
    ws.getCell(`A${i + 1}`).value = value;
  });
  return wb.xlsx.writeBuffer();
}

async function writeStreamInline() {
  const stream = new PassThrough();
  const chunks = [];
  stream.on('data', chunk => chunks.push(chunk));
  const wb = new ExcelJS.stream.xlsx.WorkbookWriter({stream, useSharedStrings: false});
  const ws = wb.addWorksheet('Sheet1');
  VALUES.forEach((value, i) => {
    ws.getCell(`A${i + 1}`).value = value;
  });
  ws.commit();
  await wb.commit();
  return Buffer.concat(chunks);
}

async function writeAndCheck(write, part) {
  const buffer = await write();
  const xml = await readPart(buffer, part);
  // Excel reports an empty <r> run as corrupt
  expect(xml).not.to.match(/<r>(<rPr>.*?<\/rPr>)?<t\/>|<t><\/t><\/r>/);
  expect(xml).to.include('<r><t>kept</t></r>');

  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.load(buffer);
  const ws2 = wb2.getWorksheet('Sheet1');
  expect(ws2.getCell('A1').text).to.equal('kept');
  expect(ws2.getCell('A2').text).to.equal('');
}

describe('github issues', () => {
  it('issue 2737 - empty rich text runs are not written (shared strings)', () =>
    writeAndCheck(writeWorkbook, 'xl/sharedStrings.xml'));

  it('issue 2737 - empty rich text runs are not written (streamed inline strings)', () =>
    writeAndCheck(writeStreamInline, 'xl/worksheets/sheet1.xml'));
});
