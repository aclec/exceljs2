const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

// 1x1 transparent PNG
const PNG = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
  'base64'
);

describe('github issues', () => {
  it('issue 3092 - every picture gets its own a16:creationId', async () => {
    const wb = new ExcelJS.Workbook();
    const imageId = wb.addImage({buffer: PNG, extension: 'png'});
    ['One', 'Two'].forEach(name => {
      const ws = wb.addWorksheet(name);
      ws.addImage(imageId, 'A1:B2');
      ws.addImage(imageId, 'D1:E2');
    });

    const zip = await JSZip.loadAsync(await wb.xlsx.writeBuffer());
    const ids = [];
    for (const drawing of ['drawing1', 'drawing2']) {
      // eslint-disable-next-line no-await-in-loop
      const xml = await zip.file(`xl/drawings/${drawing}.xml`).async('string');
      xml.replace(/<a16:creationId [^>]*id="([^"]+)"/g, (m, id) => ids.push(id));
    }
    expect(ids).to.have.length(4);
    ids.forEach(id => expect(id).to.match(/^\{[0-9A-F]{8}(-[0-9A-F]{4}){3}-[0-9A-F]{12}\}$/));
    expect(new Set(ids).size).to.equal(4);
  });
});
