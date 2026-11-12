const { readPart } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

describe("shared strings", () => {
    it("count every reference in the workbook, and uniqueCount the strings", async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet("Sheet1");
        ws.getCell("A1").value = "a";
        ws.getCell("A2").value = "a";
        ws.getCell("A3").value = "b";
        ws.getCell("B1").value = { richText: [{ text: "rich" }] };
        ws.getCell("B2").value = { richText: [{ text: "rich" }] };
        // a formula result is not a shared string
        ws.getCell("C1").value = { formula: "A1", result: "a" };
        wb.addWorksheet("Sheet2").getCell("A1").value = "a";

        const buffer = await wb.xlsx.writeBuffer();
        expect(await readPart(buffer, "xl/sharedStrings.xml")).to.match(/<sst [^>]* count="6" uniqueCount="3">/);

        // and so once read back
        const back = new ExcelJS.Workbook();
        await back.xlsx.load(buffer);
        expect(await readPart(await back.xlsx.writeBuffer(), "xl/sharedStrings.xml")).to.match(/<sst [^>]* count="6" uniqueCount="3">/);
    });
});
