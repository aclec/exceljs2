// Column keys and table names can come from data (CSV headers, loaded files, JSON):
// names of Object.prototype members must behave like any other name.
const Excel = verquire('exceljs');

function proto() {
  return Object.getPrototypeOf({});
}

describe('Worksheet keys named like Object.prototype members', () => {
  it('uses __proto__ and constructor as ordinary column keys', () => {
    const ws = new Excel.Workbook().addWorksheet('s');
    ws.columns = ['__proto__', 'constructor', 'name'].map(key => ({key, header: key}));

    expect(ws.getColumnKey('__proto__').number).to.equal(1);
    expect(ws.getColumn('constructor').number).to.equal(2);
    const keys = [];
    ws.eachColumnKey((column, key) => keys.push([key, column.number]));
    expect(keys).to.deep.equal([
      ['__proto__', 1],
      ['constructor', 2],
      ['name', 3],
    ]);

    ws.addRow(JSON.parse('{"__proto__": 1, "constructor": 2, "name": 3}'));
    expect(ws.getRow(2).values.slice(1)).to.deep.equal([1, 2, 3]);
    expect(ws.getRow(2).getCell('__proto__').value).to.equal(1);
    expect(ws.getRow(2).getCell('constructor').value).to.equal(2);

    ws.getColumn('__proto__').key = 'renamed';
    expect(ws.getColumnKey('__proto__')).to.equal(undefined);
    expect(ws.getColumnKey('renamed').number).to.equal(1);
    expect(proto()).to.equal(Object.prototype);
  });

  it('does not resolve unknown keys to inherited members', () => {
    const ws = new Excel.Workbook().addWorksheet('s');
    ws.columns = [{key: 'id'}];
    ['constructor', 'toString', 'hasOwnProperty', '__proto__'].forEach(key => {
      expect(ws.getColumnKey(key), key).to.equal(undefined);
      expect(() => ws.getColumn(key), key).to.throw(/Out of bounds/);
      expect(() => ws.getRow(1).getCell(key), key).to.throw(/Out of bounds/);
    });
  });

  it('only reads own properties of row objects', () => {
    const ws = new Excel.Workbook().addWorksheet('s');
    ws.columns = ['id', 'toString', 'constructor'].map(key => ({key}));

    ws.addRow({id: 1});
    ws.getRow(2).values = Object.assign(Object.create({toString: 'inherited'}), {id: 2});

    expect(ws.getRow(1).values.slice(1)).to.deep.equal([1]);
    expect(ws.getRow(2).values.slice(1)).to.deep.equal([2]);
  });

  it('still reads getters of class instances', () => {
    class Person {
      constructor(first) {
        this.first = first;
      }

      get greeting() {
        return `hi ${this.first}`;
      }
    }
    const ws = new Excel.Workbook().addWorksheet('s');
    ws.columns = ['first', 'greeting'].map(key => ({key}));
    ws.addRow(new Person('ann'));
    expect(ws.getRow(1).values.slice(1)).to.deep.equal(['ann', 'hi ann']);
  });

  it('keeps a table named __proto__', () => {
    const ws = new Excel.Workbook().addWorksheet('s');
    ws.addTable({
      name: '__proto__',
      ref: 'A1',
      columns: [{name: 'a'}],
      rows: [[1]],
    });
    expect(ws.getTables().map(table => table.name)).to.deep.equal(['__proto__']);
    expect(ws.getTable('__proto__').name).to.equal('__proto__');
    expect(ws.getTable('constructor')).to.equal(undefined);

    const ws2 = new Excel.Workbook().addWorksheet('s2');
    ws2.model = Object.assign(ws.model, {name: 's2'});
    expect(ws2.getTables().map(table => table.name)).to.deep.equal(['__proto__']);
    expect(ws2.getTable('toString')).to.equal(undefined);
  });
});
