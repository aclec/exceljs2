// Type-checked by `bun run test:types` (tsc) and run by `bun run test:typescript` (tsx).
// 'exceljs' is a self-reference: it resolves through the "exports" map of package.json, so this
// also checks that the published entry point and its types line up.
import {PassThrough} from 'stream';
import {expect} from 'chai';
import * as ExcelJS from 'exceljs';

describe('typescript', () => {
  it('can create and buffer xlsx', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');
    ws.getCell('A1').value = 7;
    const buffer = await wb.xlsx.writeBuffer({
      useStyles: true,
      useSharedStrings: true,
    });

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(buffer);
    const ws2 = wb2.getWorksheet('blort')!;
    expect(ws2.getCell('A1').value).to.equal(7);
  });

  it('can create and stream xlsx', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');
    ws.getCell('A1').value = 7;

    const stream = new PassThrough();
    const wb2 = new ExcelJS.Workbook();
    const [, loaded] = await Promise.all([wb.xlsx.write(stream).then(() => stream.end()), wb2.xlsx.read(stream)]);

    expect(loaded).to.equal(wb2);
    expect(wb2.getWorksheet('blort')!.getCell('A1').value).to.equal(7);
  });

  it('exposes the enums', () => {
    expect(ExcelJS.ValueType.Number).to.equal(2);
  });
});
