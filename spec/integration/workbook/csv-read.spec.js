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

    it('keeps integers beyond Number.MAX_SAFE_INTEGER as strings (#1743)', async () => {
      const values = await readCells([
        '56343416020533614003',
        '9007199254740993',
        '-9007199254740993',
        '9007199254740992',
      ]);
      expect(values).to.deep.equal([
        '56343416020533614003',
        '9007199254740993',
        '-9007199254740993',
        '9007199254740992',
      ]);
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
        9007199254740991, -9007199254740991, 0.5, 100000, 5, 7, -0, 12, 1.1, 1.5e300, 1e21,
        12345678901234567890.5, 31,
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

  describe('parserOptions.headers (#2080)', () => {
    const CSV = 'Name,Full Name,Birthday,Age\nJim,Jim Green,2019-11-04,3\nAnn,Ann Lee,,40\n';
    const column = (ws, key) => ws.getColumn(key).values.slice(1);

    it('headers: true keys the columns by the header row', async () => {
      const ws = await readCsv(CSV, {parserOptions: {headers: true}});
      expect(ws.rowCount).to.equal(3);
      expect(column(ws, 'Name')).to.deep.equal(['Name', 'Jim', 'Ann']);
      expect(column(ws, 'Full Name')).to.deep.equal(['Full Name', 'Jim Green', 'Ann Lee']);
      expect(ws.getCell('C2').value).to.be.an.instanceof(Date);
      expect(ws.getCell('C3').value).to.equal(null);
      expect(column(ws, 'Age')).to.deep.equal(['Age', 3, 40]);
      expect(ws.getColumn('Age')).to.equal(ws.getColumn('D'));
    });

    it('applies a header transform function', async () => {
      const ws = await readCsv(CSV, {
        parserOptions: {headers: names => names.map(name => name.toUpperCase())},
      });
      expect(column(ws, 'FULL NAME')).to.deep.equal(['FULL NAME', 'Jim Green', 'Ann Lee']);
    });

    it('renames headers and skips undefined ones', async () => {
      const ws = await readCsv(CSV, {
        parserOptions: {headers: ['name', 'full', undefined, 'age'], renameHeaders: true},
      });
      expect(ws.rowCount).to.equal(3);
      expect(column(ws, 'name')).to.deep.equal(['name', 'Jim', 'Ann']);
      expect(column(ws, 'age')).to.deep.equal(['age', 3, 40]);
      expect(ws.getColumn(3).values).to.deep.equal([]);
    });

    it('uses custom headers on a file without a header row', async () => {
      const ws = await readCsv('Jim,3\n', {parserOptions: {headers: ['name', 'age']}});
      expect(ws.getRow(1).values.slice(1)).to.deep.equal(['name', 'age']);
      expect(ws.getRow(2).values.slice(1)).to.deep.equal(['Jim', 3]);
    });

    it('passes the value and column index to a custom map', async () => {
      const calls = [];
      await readCsv('a,b\n1,2\n', {
        parserOptions: {headers: true},
        map: (value, index) => calls.push([value, index]) && value,
      });
      expect(calls).to.deep.equal([
        ['1', 0],
        ['2', 1],
      ]);
    });

    it('treats __proto__ and constructor headers as plain names', async () => {
      const ws = await readCsv('__proto__,constructor,name\n1,2,3\n', {
        parserOptions: {headers: true},
      });
      expect(ws.getRow(1).values.slice(1)).to.deep.equal(['__proto__', 'constructor', 'name']);
      expect(ws.getColumn('__proto__').number).to.equal(1);
      expect(ws.getColumn('constructor').number).to.equal(2);
      // fast-csv builds rows as plain objects, which cannot hold a '__proto__' field
      expect(ws.getCell('A2').value).to.equal(null);
      expect(ws.getCell('B2').value).to.equal(2);
      expect(ws.getCell('C2').value).to.equal(3);
      expect({}.name).to.equal(undefined);
      await ws.workbook.xlsx.writeBuffer();
    });

    it('does not copy a value into several columns with an empty header', async () => {
      const ws = await readCsv('a,,b,\n1,2,3,4\n', {parserOptions: {headers: true}});
      // fast-csv keys both empty headers by '' and keeps the last value only
      expect(ws.getRow(2).values).to.deep.equal([undefined, 1, undefined, 3, 4]);
    });

    it('still reads rows as arrays without the option', async () => {
      const ws = await readCsv(CSV);
      expect(ws.getRow(1).values.slice(1)).to.deep.equal(['Name', 'Full Name', 'Birthday', 'Age']);
      expect(ws.getColumnKey('Name')).to.equal(undefined);
    });
  });
});
