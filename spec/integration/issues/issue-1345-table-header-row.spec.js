const {baseWorkbookBuffer, tamper, readPart} = require('../../utils/xlsx-tamper');

const ExcelJS = verquire('exceljs');

const TABLE = 'xl/tables/table1.xml';

function tableBuffer(headerRow) {
  return baseWorkbookBuffer((wb, ws) => {
    ws.addTable({
      name: 'T1',
      ref: 'C1',
      headerRow,
      columns: [{name: 'a'}, {name: 'b'}],
      rows: [
        [1, 2],
        [3, 4],
      ],
    });
  });
}

async function load(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  return wb;
}

describe('github issues', () => {
  describe('issue 1345 - tables without a header row', () => {
    it('does not write an autoFilter when the table has no header row', async () => {
      const buffer = await tableBuffer(false);
      const xml = await readPart(buffer, TABLE);
      expect(xml).to.include('headerRowCount="0"');
      expect(xml).not.to.include('<autoFilter');

      const wb = await load(buffer);
      expect(wb.getWorksheet('Sheet1').getTable('T1').table.headerRow).to.equal(false);
    });

    it('reads a missing headerRowCount as one header row (the ECMA default)', async () => {
      const buffer = await tamper(await tableBuffer(true), {
        [TABLE]: xml => {
          const patched = xml.replace(' headerRowCount="1"', '');
          expect(patched).not.to.equal(xml);
          return patched;
        },
      });
      const wb = await load(buffer);
      expect(wb.getWorksheet('Sheet1').getTable('T1').table.headerRow).to.equal(true);

      const xml = await readPart(await wb.xlsx.writeBuffer(), TABLE);
      expect(xml).to.include('headerRowCount="1"');
      expect(xml).to.include('<autoFilter');
    });
  });
});
