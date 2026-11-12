const {baseWorkbookBuffer, tamper} = require('../../utils/xlsx-tamper');

const ExcelJS = verquire('exceljs');

async function loadWithFilterColumns(filterColumns) {
  const source = await baseWorkbookBuffer((wb, ws) => {
    ws.addTable({
      name: 'T1',
      ref: 'C1',
      headerRow: true,
      columns: [{name: 'a'}, {name: 'b'}, {name: 'c'}],
      rows: [[1, 2, 3]],
    });
  });
  const buffer = await tamper(source, {
    'xl/tables/table1.xml': xml => {
      const patched = xml.replace(/(<autoFilter [^>]*>).*?(<\/autoFilter>)/, `$1${filterColumns}$2`);
      expect(patched).not.to.equal(xml);
      return patched;
    },
  });
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  return wb.getWorksheet('Sheet1').getTable('T1').table;
}

describe('github issues', () => {
  describe('issue 2978 - table autoFilter columns', () => {
    it('ignores filterColumns beyond the table columns', async () => {
      const table = await loadWithFilterColumns(
        '<filterColumn colId="0" hiddenButton="0"/><filterColumn colId="1" hiddenButton="1"/>' +
          '<filterColumn colId="2" hiddenButton="1"/><filterColumn colId="7" hiddenButton="0"/>'
      );
      expect(table.columns.map(column => column.name)).to.deep.equal(['a', 'b', 'c']);
      expect(table.columns[0].filterButton).to.equal(true);
    });

    it('applies a filterColumn to the column named by its colId', async () => {
      // columns without a filterColumn keep the default (visible) button
      const table = await loadWithFilterColumns('<filterColumn colId="2" hiddenButton="1"/>');
      expect(table.columns.map(column => Boolean(column.filterButton))).to.deep.equal([true, true, false]);
    });

    it('does not touch Array.prototype for a hostile colId', async () => {
      await loadWithFilterColumns('<filterColumn colId="__proto__" hiddenButton="0"/>');
      expect([].filterButton).to.equal(undefined);
    });
  });
});
