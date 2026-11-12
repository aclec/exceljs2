const {baseWorkbookBuffer, tamper, readPart} = require('../../utils/xlsx-tamper');

const ExcelJS = verquire('exceljs');

const TABLE = 'xl/tables/table1.xml';

async function tableBuffer() {
  return baseWorkbookBuffer((wb, ws) => {
    ws.addTable({
      name: 'T1',
      ref: 'C1',
      headerRow: true,
      columns: [
        {name: 'A', filterButton: true},
        {name: 'B', filterButton: true},
      ],
      rows: [
        [1, 2],
        [3, 4],
      ],
    });
  });
}

async function roundTripTableXml(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  return readPart(await wb.xlsx.writeBuffer(), TABLE);
}

describe('table filter buttons on round trip', () => {
  it('keeps buttons visible when filterColumn has no hiddenButton attribute (spec default)', async () => {
    const buffer = await tamper(await tableBuffer(), {
      [TABLE]: xml => xml.replace(/ hiddenButton="0"/g, ''),
    });
    expect(await roundTripTableXml(buffer)).not.to.include('hiddenButton="1"');
  });

  it('keeps buttons visible for columns without a filterColumn element', async () => {
    const buffer = await tamper(await tableBuffer(), {
      [TABLE]: xml => xml.replace(/<filterColumn [^>]*\/>/g, ''),
    });
    expect(await roundTripTableXml(buffer)).not.to.include('hiddenButton="1"');
  });

  it('keeps explicitly hidden buttons hidden', async () => {
    const buffer = await tamper(await tableBuffer(), {
      [TABLE]: xml => xml.replace('colId="1" hiddenButton="0"', 'colId="1" hiddenButton="1"'),
    });
    const xml = await roundTripTableXml(buffer);
    expect(xml).to.include('colId="0" hiddenButton="0"');
    expect(xml).to.include('colId="1" hiddenButton="1"');
  });
});
