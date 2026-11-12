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

// CT_Table requires a name and <tableColumns>, CT_TableColumn a name: Excel repairs a file without them by removing the
// table; the workbook loads without that table (its cells are kept) and writes back.
describe("Invalid table part", () => {
    const tableBuffer = () =>
        baseWorkbookBuffer((wb, ws) => {
            ws.addTable({ name: "T", ref: "C1", columns: [{ name: "A" }, { name: "B" }], rows: [[1, 2]] });
        });
    const patches = {
        "a table without name": (xml) => xml.replace(/ name="T"/, ""),
        "a table without <tableColumns>": (xml) => xml.replace(/<tableColumns.*<\/tableColumns>/, ""),
        "a column without name": (xml) => xml.replace(/<tableColumn id="1" name="A"/, '<tableColumn id="1"'),
    };
    for (const [title, patch] of Object.entries(patches)) {
        it(`is skipped: ${title}`, async () => {
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.load(await tamper(await tableBuffer(), { [TABLE]: patch }));
            const ws = wb.getWorksheet("Sheet1");
            expect(ws.getTables()).to.deep.equal([]);
            expect(ws.getCell("C2").value).to.equal(1);
            const wb2 = new ExcelJS.Workbook();
            await wb2.xlsx.load(await wb.xlsx.writeBuffer());
            expect(wb2.getWorksheet("Sheet1").getCell("D2").value).to.equal(2);
        });
    }
});
