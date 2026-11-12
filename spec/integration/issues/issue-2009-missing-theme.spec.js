const JSZip = require('jszip');
const {baseWorkbookBuffer, tamper} = require('../../utils/xlsx-tamper');

const ExcelJS = verquire('exceljs');

async function withoutTheme(buffer) {
  const patched = await tamper(buffer, {
    'xl/_rels/workbook.xml.rels': xml => xml.replace(/<Relationship [^>]*theme1\.xml"\/>/, ''),
    '[Content_Types].xml': xml => xml.replace(/<Override PartName="\/xl\/theme\/theme1\.xml"[^>]*\/>/, ''),
  });
  const zip = await JSZip.loadAsync(patched);
  zip.remove('xl/theme/theme1.xml');
  return zip.generateAsync({type: 'nodebuffer'});
}

describe('github issues', () => {
  it('issue 2009 - a workbook read without a theme writes the theme it references', async () => {
    const source = await withoutTheme(await baseWorkbookBuffer());
    expect((await JSZip.loadAsync(source)).file('xl/theme/theme1.xml')).to.equal(null);

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(source);
    const zip = await JSZip.loadAsync(await wb.xlsx.writeBuffer());

    const rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    const contentTypes = await zip.file('[Content_Types].xml').async('string');
    expect(rels).to.include('Target="theme/theme1.xml"');
    expect(contentTypes).to.include('/xl/theme/theme1.xml');
    expect(zip.file('xl/theme/theme1.xml')).not.to.equal(null);
  });
});
