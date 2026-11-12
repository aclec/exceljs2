// Regression tests for hostile xlsx/csv input: each case used to crash the process
// (heap OOM), block the event loop, or pollute Object.prototype.
const {Readable} = require('stream');
const {baseWorkbookBuffer, tamper, readPart} = require('../../utils/xlsx-tamper');

const ExcelJS = verquire('exceljs');

const SHEET = 'xl/worksheets/sheet1.xml';
const afterSheetData = xml => fragment => xml.replace('</sheetData>', `</sheetData>${fragment}`);

async function load(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  return wb;
}

async function timed(fn) {
  const start = Date.now();
  const result = await fn();
  return {result, ms: Date.now() - start};
}

describe('hostile input', function() {
  this.timeout(20000);

  describe('merge cells', () => {
    it('rejects a merge covering the whole grid instead of exhausting memory', async () => {
      const buffer = await tamper(await baseWorkbookBuffer(), {
        [SHEET]: xml => afterSheetData(xml)('<mergeCells count="1"><mergeCell ref="A3:XFD1048576"/></mergeCells>'),
      });
      let error;
      await load(buffer).catch(e => {
        error = e;
      });
      expect(error).to.be.an('error');
      expect(error.message).to.match(/merged cells/);
    });

    it('rejects references beyond the last Excel row', async () => {
      const buffer = await tamper(await baseWorkbookBuffer(), {
        [SHEET]: xml => afterSheetData(xml)('<mergeCells count="1"><mergeCell ref="A1:B99999999999"/></mergeCells>'),
      });
      let error;
      await load(buffer).catch(e => {
        error = e;
      });
      expect(error).to.be.an('error');
      expect(error.message).to.match(/Out of bounds/);
    });

    it('loads many small merges in linear time', async () => {
      const merges = [];
      for (let row = 1; row <= 20000; row++) {
        merges.push(`<mergeCell ref="B${row}:C${row}"/>`);
      }
      const buffer = await tamper(await baseWorkbookBuffer(), {
        [SHEET]: xml => afterSheetData(xml)(`<mergeCells count="${merges.length}">${merges.join('')}</mergeCells>`),
      });
      const {result: wb, ms} = await timed(() => load(buffer));
      expect(wb.getWorksheet('Sheet1').getCell('C20000').isMerged).to.equal(true);
      expect(ms).to.be.below(5000);
    });

    it('still rejects overlapping merges without modifying the sheet', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('s');
      ws.mergeCells('B2:C3');
      expect(() => ws.mergeCells('A1:B2')).to.throw(/already merged/);
      expect(() => ws.mergeCells('C3:D4')).to.throw(/already merged/);
      expect(ws.getCell('A1').isMerged).to.equal(false);
      expect(ws.getCell('D4').isMerged).to.equal(false);
      ws.unMergeCells('B2');
      ws.mergeCells('A1:B2');
      expect(ws.getCell('B2').master.address).to.equal('A1');
    });
  });

  describe('data validations', () => {
    it('keeps a huge sqref as a range instead of expanding it per cell', async () => {
      const buffer = await tamper(await baseWorkbookBuffer(), {
        [SHEET]: xml =>
          afterSheetData(xml)(
            '<dataValidations count="1"><dataValidation type="list" allowBlank="1" sqref="B1:XFD1048576">' +
              '<formula1>"a,b"</formula1></dataValidation></dataValidations>'
          ),
      });
      const {result: wb, ms} = await timed(() => load(buffer));
      expect(ms).to.be.below(5000);
      const ws = wb.getWorksheet('Sheet1');
      expect(ws.getCell('C5').dataValidation).to.deep.include({type: 'list', formulae: ['"a,b"']});
      expect(ws.getCell('A5').dataValidation).to.equal(undefined);

      // the range survives a write / read round trip
      const out = await wb.xlsx.writeBuffer();
      expect(await readPart(out, SHEET)).to.include('sqref="B1:XFD1048576"');
    });

    it('still expands small ranges per cell', async () => {
      const buffer = await tamper(await baseWorkbookBuffer(), {
        [SHEET]: xml =>
          afterSheetData(xml)(
            '<dataValidations count="1"><dataValidation type="whole" operator="greaterThan" sqref="A2:B3 D4">' +
              '<formula1>5</formula1></dataValidation></dataValidations>'
          ),
      });
      const ws = (await load(buffer)).getWorksheet('Sheet1');
      expect(Object.keys(ws.dataValidations.model).sort()).to.deep.equal(['A2', 'A3', 'B2', 'B3', 'D4']);
    });

    it('ignores sqref tokens that are not cell references', async () => {
      const buffer = await tamper(await baseWorkbookBuffer(), {
        [SHEET]: xml =>
          afterSheetData(xml)(
            '<dataValidations count="1"><dataValidation type="list" sqref="__proto__ constructor A1">' +
              '<formula1>"a"</formula1></dataValidation></dataValidations>'
          ),
      });
      const ws = (await load(buffer)).getWorksheet('Sheet1');
      expect(Object.keys(ws.dataValidations.model)).to.deep.equal(['A1']);
      expect(Object.getPrototypeOf(ws.dataValidations.model)).to.equal(Object.prototype);
    });

    it('does not write validations that were removed', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('s');
      ws.getCell('A1').dataValidation = {type: 'list', formulae: ['"x"']};
      ws.getCell('A2').dataValidation = {type: 'list', formulae: ['"x"']};
      ws.dataValidations.remove('A2');
      const xml = await readPart(await wb.xlsx.writeBuffer(), SHEET);
      expect(xml).to.include('<dataValidations count="1">');
      expect(xml).to.include('sqref="A1"');
    });
  });

  it('clamps <col max> to the last Excel column', async () => {
    const buffer = await tamper(await baseWorkbookBuffer(), {
      [SHEET]: xml => xml.replace('<sheetData>', '<cols><col min="1" max="2000000000" width="12" customWidth="1"/></cols><sheetData>'),
    });
    const ws = (await load(buffer)).getWorksheet('Sheet1');
    expect(ws.columns.length).to.equal(16384);
    expect(ws.getColumn(16384).width).to.equal(12);
  });

  describe('defined names', () => {
    const withDefinedNames = names => ({
      'xl/workbook.xml': xml => xml.replace('</sheets>', `</sheets><definedNames>${names}</definedNames>`),
    });

    it('keeps a huge defined-name range as a rectangle', async () => {
      const buffer = await tamper(
        await baseWorkbookBuffer(),
        withDefinedNames('<definedName name="everything">Sheet1!$A$1:$XFD$1048576</definedName>')
      );
      const {result: wb, ms} = await timed(() => load(buffer));
      expect(ms).to.be.below(5000);
      expect(wb.getWorksheet('Sheet1').getCell('Z99').names).to.deep.equal(['everything']);
      expect(wb.definedNames.model).to.deep.equal([
        {name: 'everything', ranges: ['Sheet1!$A$1:$XFD$1048576']},
      ]);
    });

    it('shifts and removes large ranges with row splices', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('s');
      wb.definedNames.add('s!$A$1:$B$100000', 'big');
      ws.spliceRows(1, 2);
      expect(wb.definedNames.getRanges('big').ranges).to.deep.equal(['s!$A$1:$B$99998']);
      ws.spliceRows(1, 0, [], []);
      expect(wb.definedNames.getRanges('big').ranges).to.deep.equal(['s!$A$3:$B$100000']);
      ws.getCell('A50').removeName('big');
      expect(wb.definedNames.getRanges('big').ranges).to.deep.equal([]);
    });

    it('does not pollute Object.prototype through a "__proto__" sheet name', async () => {
      const buffer = await tamper(
        await baseWorkbookBuffer(),
        withDefinedNames(
          '<definedName name="a">__proto__!$A$7</definedName><definedName name="b">\'__proto__\'!$B$9</definedName>'
        )
      );
      await load(buffer);
      expect({}[7]).to.equal(undefined);
      expect({}[9]).to.equal(undefined);
      expect(7 in {}).to.equal(false);
    });
  });

  it('parses a drawing relationship with a pathological Target in linear time', async () => {
    const target = `..${'/media/'.repeat(20000)}`;
    const base = await baseWorkbookBuffer((wb, ws) => {
      const imageId = wb.addImage({filename: './spec/integration/data/image.png', extension: 'png'});
      ws.addImage(imageId, 'B2:C3');
    });
    const buffer = await tamper(base, {
      'xl/drawings/_rels/drawing1.xml.rels': xml => xml.replace('../media/image1.png', target),
    });
    const {ms} = await timed(() => load(buffer));
    expect(ms).to.be.below(2000);
  });

  it('rejects formulas longer than the Excel limit instead of scanning them quadratically', () => {
    const {slideFormula} = verquire('utils/shared-formula');
    expect(() => slideFormula('a-'.repeat(5000), 'A1', 'A2')).to.throw(/Excel limit/);
    expect(slideFormula('A1+$B$2+Sheet1!C3', 'A1', 'B2')).to.equal('B2+$B$2+Sheet1!D4');
  });

  it('does not merge "__proto__" keys from app-supplied objects (cell notes)', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('s');
    ws.getCell('A1').note = JSON.parse('{"texts":[{"text":"hi"}],"__proto__":{"polluted":"yes"}}');
    await wb.xlsx.writeBuffer();
    expect({}.polluted).to.equal(undefined);
  });

  it('reads CSV cells named like Object.prototype members as plain text', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = await wb.csv.read(Readable.from(['__proto__,constructor,toString,true,#N/A\n']));
    expect(ws.getRow(1).values.slice(1)).to.deep.equal([
      '__proto__',
      'constructor',
      'toString',
      true,
      {error: '#N/A'},
    ]);
    await wb.xlsx.writeBuffer(); // used to throw on the Object.prototype cell value
  });

  it('strips characters that are not allowed in XML 1.0 when writing', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('s');
    ws.getCell('A1').value = 'a￾b￿c\uD800d';
    ws.getCell('A2').value = 'emoji 😀 stays';
    const wb2 = await load(await wb.xlsx.writeBuffer());
    const ws2 = wb2.getWorksheet('s');
    expect(ws2.getCell('A1').value).to.equal('abcd');
    expect(ws2.getCell('A2').value).to.equal('emoji 😀 stays');
  });

  it('escapes pivot table names and values', async () => {
    const wb = new ExcelJS.Workbook();
    const source = wb.addWorksheet('src');
    const evil = 'x" injected="1"/><evil/><s v="';
    source.addRows([
      ['A', 'B', 'C'],
      [evil, 'b1', 5],
      ['a2', 'b2', 7],
    ]);
    const pivot = wb.addWorksheet('pivot');
    pivot.addPivotTable({
      sourceSheet: source,
      rows: ['A'],
      columns: ['B'],
      values: ['C'],
      metric: 'sum',
    });
    const out = await wb.xlsx.writeBuffer();
    const definition = await readPart(out, 'xl/pivotCache/pivotCacheDefinition1.xml');
    expect(definition).not.to.include('<evil/>');
    expect(definition).to.include('x&quot; injected=&quot;1&quot;/&gt;&lt;evil/&gt;');
    await load(out); // output is well-formed
  });
});
