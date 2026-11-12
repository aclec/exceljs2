const { baseWorkbookBuffer, editParts, readPart, tamper, unzipParts } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

async function withoutTheme(buffer) {
    const patched = await tamper(buffer, {
        "xl/_rels/workbook.xml.rels": (xml) => xml.replace(/<Relationship [^>]*theme1\.xml"\/>/, ""),
        "[Content_Types].xml": (xml) => xml.replace(/<Override PartName="\/xl\/theme\/theme1\.xml"[^>]*\/>/, ""),
    });
    return editParts(patched, (parts) => {
        delete parts["xl/theme/theme1.xml"];
    });
}

describe("github issues", () => {
    it("issue 2009 - a workbook read without a theme writes the theme it references", async () => {
        const source = await withoutTheme(await baseWorkbookBuffer());
        expect(unzipParts(source)["xl/theme/theme1.xml"]).to.equal(undefined);

        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(source);
        const buffer = await wb.xlsx.writeBuffer();

        const rels = await readPart(buffer, "xl/_rels/workbook.xml.rels");
        const contentTypes = await readPart(buffer, "[Content_Types].xml");
        expect(rels).to.include('Target="theme/theme1.xml"');
        expect(contentTypes).to.include("/xl/theme/theme1.xml");
        expect(unzipParts(buffer)["xl/theme/theme1.xml"]).not.to.equal(undefined);
    });
});
