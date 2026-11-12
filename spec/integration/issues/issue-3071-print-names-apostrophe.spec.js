const { readPart } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

describe("github issues", () => {
    it("issue 3071 - apostrophes in sheet names are doubled in print areas and titles", async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet("Bob's sheet");
        ws.getCell("A1").value = "x";
        ws.pageSetup.printArea = "A1:B2&&D4:E5";
        ws.pageSetup.printTitlesRow = "1:2";
        ws.pageSetup.printTitlesColumn = "A:B";

        const buffer = await wb.xlsx.writeBuffer();
        const xml = (await readPart(buffer, "xl/workbook.xml")).replace(/&apos;/g, "'");
        expect(xml).to.include(
            '<definedName name="_xlnm.Print_Area" localSheetId="0">' + "'Bob''s sheet'!$A$1:$B$2,'Bob''s sheet'!$D$4:$E$5</definedName>"
        );
        expect(xml).to.include(
            '<definedName name="_xlnm.Print_Titles" localSheetId="0">' + "'Bob''s sheet'!$A:$B,'Bob''s sheet'!$1:$2</definedName>"
        );

        const wb2 = new ExcelJS.Workbook();
        await wb2.xlsx.load(buffer);
        const { pageSetup } = wb2.getWorksheet("Bob's sheet");
        expect(pageSetup.printArea).to.equal("A1:B2&&D4:E5");
        expect(pageSetup.printTitlesRow).to.equal("1:2");
        expect(pageSetup.printTitlesColumn).to.equal("A:B");
    });
});
