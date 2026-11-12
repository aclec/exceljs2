const { baseWorkbookBuffer, tamper, readPart } = require("../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

const TABLE = "xl/tables/table1.xml";

// CT_Table: <tableStyleInfo> is optional, a table without it has no table style
async function tableWithoutStyle() {
    const buffer = await baseWorkbookBuffer((wb, ws) => {
        ws.addTable({ name: "T", ref: "C1", columns: [{ name: "A" }, { name: "B" }], rows: [[1, 2]] });
    });
    return tamper(buffer, { [TABLE]: (xml) => xml.replace(/<tableStyleInfo[^>]*\/>/, "") });
}

describe("Table without table style", () => {
    it("loads with no style and writes back without <tableStyleInfo>", async () => {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(await tableWithoutStyle());
        const table = wb.getWorksheet("Sheet1").getTable("T");
        expect(table.theme).to.equal(null);

        const xml = await readPart(await wb.xlsx.writeBuffer(), TABLE);
        expect(xml).to.include('name="T"');
        expect(xml).not.to.include("<tableStyleInfo");

        const wb2 = new ExcelJS.Workbook();
        await wb2.xlsx.load(await wb.xlsx.writeBuffer());
        expect(wb2.getWorksheet("Sheet1").getTable("T").theme).to.equal(null);
    });
});
