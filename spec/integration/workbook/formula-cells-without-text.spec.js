// A clone of a shared formula whose master is not in the file is read as a formula cell without formula text: it was
// written back as an empty normal formula (<f/>), which is not a formula.
const { baseWorkbookBuffer, tamper, readPart } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

describe("Formula cells without formula text", () => {
    it("writes an orphan shared formula clone as its cached value", async () => {
        const buffer = await tamper(await baseWorkbookBuffer(), {
            "xl/worksheets/sheet1.xml": (xml) =>
                xml.replace("</sheetData>", '<row r="2"><c r="A2"><f t="shared" si="7"/><v>3</v></c></row></sheetData>'),
        });
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buffer);
        const written = await wb.xlsx.writeBuffer();
        const sheet = await readPart(written, "xl/worksheets/sheet1.xml");
        expect(sheet).to.include('<c r="A2"><v>3</v></c>');
        expect(sheet).not.to.include("<f");

        const back = new ExcelJS.Workbook();
        await back.xlsx.load(written);
        expect(back.getWorksheet("Sheet1").getCell("A2").value).to.equal(3);
    });
});
