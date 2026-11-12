const {Readable} = require('stream');

const ExcelJS = verquire('exceljs');

async function readCsv(text, options) {
  const wb = new ExcelJS.Workbook();
  return wb.csv.read(Readable.from([text]), options);
}

async function readCells(cells) {
  const ws = await readCsv(`${cells.map(cell => `"${cell}"`).join(',')}\n`);
  return ws.getRow(1).values.slice(1);
}

describe('CSV read', () => {
  describe('default value mapping', () => {
    it('keeps whitespace-only cells as strings (#2752)', async () => {
      const ws = await readCsv('a;    ;b;\t;\n', {parserOptions: {delimiter: ';'}});
      expect(ws.getCell('A1').value).to.equal('a');
      expect(ws.getCell('B1').value).to.equal('    ');
      expect(ws.getCell('C1').value).to.equal('b');
      expect(ws.getCell('D1').value).to.equal('\t');
      expect(ws.getCell('E1').value).to.equal(null); // empty cells stay empty
    });

    it('still reads safe integers and other numeric forms as numbers', async () => {
      const values = await readCells([
        '9007199254740991',
        '-9007199254740991',
        '.5',
        '1e5',
        '+5',
        '007',
        '-0',
        ' 12',
        '1.10',
        '1.5e300',
        '1e21',
        '12345678901234567890.5',
        '0x1F',
      ]);
      expect(values).to.deep.equal([
        9007199254740991,
        -9007199254740991,
        0.5,
        100000,
        5,
        7,
        -0,
        12,
        1.1,
        1.5e300,
        1e21,
        12345678901234567890.5,
        31,
      ]);
      expect(Object.is(values[6], -0)).to.equal(true);
    });

    it('still maps booleans, errors and dates', async () => {
      const values = await readCells(['true', 'false', '#N/A', '2019-11-04', 'text']);
      expect(values.slice(0, 3)).to.deep.equal([true, false, {error: '#N/A'}]);
      expect(values[3]).to.be.an.instanceof(Date);
      expect(values[4]).to.equal('text');
    });
  });
});
