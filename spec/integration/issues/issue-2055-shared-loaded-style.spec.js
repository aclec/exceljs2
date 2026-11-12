// Cells read from a file share one style object per style id (see StylesXform.getStyleModel).
// Assigning a style property must only restyle that cell (or row/column), while merged cells
// keep sharing their master's style as documented in README "Merged Cells".
const {PassThrough} = require('stream');

const ExcelJS = verquire('exceljs');

const font = {name: 'Arial', size: 10};
const red = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFF0000'}};
const thick = {top: {style: 'thick'}, left: {style: 'thick'}};

async function reload(build) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('s');
  build(ws);
  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.load(await wb.xlsx.writeBuffer());
  return wb2.getWorksheet('s');
}

describe('github issues', () => {
  describe('issue 2055 - setting a style on one loaded cell changes other cells', () => {
    const setters = {
      font: {name: 'Courier', size: 20},
      fill: red,
      border: thick,
      numFmt: '0.000',
      alignment: {horizontal: 'center'},
      protection: {locked: false},
    };

    Object.keys(setters).forEach(name => {
      it(`cell.${name} = ... only changes that cell`, async () => {
        const ws = await reload(sheet => {
          ['A1', 'A2', 'B7'].forEach(address => {
            sheet.getCell(address).value = 1;
            sheet.getCell(address).font = font;
          });
        });
        expect(ws.getCell('A1').style).to.equal(ws.getCell('A2').style); // shared after load
        const before = JSON.stringify(ws.getCell('A2').style);

        ws.getCell('A1')[name] = setters[name];

        expect(ws.getCell('A1')[name]).to.deep.equal(setters[name]);
        expect(JSON.stringify(ws.getCell('A2').style)).to.equal(before);
        expect(JSON.stringify(ws.getCell('B7').style)).to.equal(before);
      });
    });

    it('keeps the change on write and read back', async () => {
      const ws = await reload(sheet => {
        sheet.getCell('A1').font = font;
        sheet.getCell('A2').font = font;
      });
      ws.getCell('A1').fill = red;
      const wb3 = new ExcelJS.Workbook();
      await wb3.xlsx.load(await ws.workbook.xlsx.writeBuffer());
      const ws3 = wb3.getWorksheet('s');
      expect(ws3.getCell('A1').fill).to.deep.equal(red);
      expect(ws3.getCell('A2').fill).not.to.deep.equal(red);
    });

    it('column.font = ... does not restyle cells of other columns sharing its loaded style', async () => {
      const ws = await reload(sheet => {
        sheet.getColumn(1).font = font;
        sheet.getColumn(2).font = font;
        sheet.getCell('C1').value = 1;
        sheet.getCell('C1').font = font;
      });
      ws.getColumn(1).font = {name: 'Courier'};
      expect(ws.getColumn(1).font).to.deep.equal({name: 'Courier'});
      expect(ws.getColumn(2).font).to.deep.equal(font);
      expect(ws.getCell('C1').font).to.deep.equal(font);
    });

    it('row.font = ... does not restyle the row it was duplicated from', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('s');
      ws.getRow(1).font = font;
      ws.getCell('A1').value = 1;
      ws.duplicateRow(1, 1, true);

      ws.getRow(2).font = {name: 'Courier'};
      ws.getCell('B2').fill = red;

      expect(ws.getRow(2).font).to.deep.equal({name: 'Courier'});
      expect(ws.getRow(1).font).to.deep.equal(font);
      expect(ws.getCell('A1').font).to.deep.equal(font);
      expect(ws.getCell('B1').fill).to.equal(undefined);
    });
  });

  describe('merged cells keep sharing the master style', () => {
    it('styling the master styles the whole merge (README example)', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('s');
      ws.mergeCells('A4:B5');

      ws.getCell('A4').border = thick;
      ws.getCell('B5').fill = red;
      ws.getCell('B5').style.font = font;

      ['A4', 'B4', 'A5', 'B5'].forEach(address => {
        expect(ws.getCell(address).style).to.equal(ws.getCell('A4').style);
        expect(ws.getCell(address).border).to.equal(thick);
        expect(ws.getCell(address).fill).to.equal(red);
        expect(ws.getCell(address).font).to.equal(font);
      });
      expect(ws.getCell('C5').border).to.equal(undefined);

      ws.unMergeCells('A4');
      expect(ws.getCell('B5').style).not.to.equal(ws.getCell('A4').style);
      ws.getCell('A4').numFmt = '0.0';
      expect(ws.getCell('B5').numFmt).to.equal(undefined);
    });

    it('styling a loaded merge restyles the merge but not unrelated cells', async () => {
      const ws = await reload(sheet => {
        ['A1', 'B1', 'A2', 'B2', 'D1'].forEach(address => {
          sheet.getCell(address).font = font;
        });
        sheet.mergeCells('A1:B2');
      });
      expect(ws.getCell('B2').style).to.equal(ws.getCell('D1').style);

      ws.getCell('A1').border = thick;

      ['A1', 'B1', 'A2', 'B2'].forEach(address => {
        expect(ws.getCell(address).border).to.equal(thick);
      });
      expect(ws.getCell('D1').border).not.to.deep.equal(thick);
    });

    it('mergeCellsWithoutStyle leaves differently styled slaves alone', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('s');
      ws.getCell('B1').fill = red;
      ws.mergeCellsWithoutStyle('A1:B1');

      ws.getCell('A1').font = font;

      expect(ws.getCell('A1').font).to.equal(font);
      expect(ws.getCell('B1').font).to.equal(undefined);
      expect(ws.getCell('B1').fill).to.equal(red);
    });

    it('re-merged cells follow their new master only', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('s');
      ws.mergeCells('A1:B1');
      ws.unMergeCells('A1');
      ws.mergeCells('B1:C1');

      ws.getCell('A1').font = font;
      ws.getCell('B1').fill = red;

      expect(ws.getCell('B1').font).to.equal(undefined);
      expect(ws.getCell('C1').fill).to.equal(red);
      expect(ws.getCell('A1').fill).to.equal(undefined);
    });

    it('works with the streaming writer', async () => {
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter({stream: new PassThrough().resume()});
      const ws = wb.addWorksheet('s');
      ws.mergeCells('A1:B2');

      ws.getCell('A1').border = thick;

      expect(ws.getCell('B2').border).to.equal(thick);
      expect(ws.getCell('B2').style).to.equal(ws.getCell('A1').style);
      await wb.commit();
    });

    it('column and row styles reach every cell of a merge they cross', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('s');
      ws.mergeCells('A1:B4000');

      ws.getColumn(1).font = font; // touches 4000 cells of an 8000-cell merge
      ws.getRow(2).fill = red;

      expect(ws.getCell('B4000').font).to.equal(font);
      expect(ws.getCell('B4000').fill).to.equal(red);
      expect(ws.getCell('B4000').style).to.equal(ws.getCell('A1').style);
      expect(ws.getCell('C3').fill).to.equal(undefined);
    });
  });
});
