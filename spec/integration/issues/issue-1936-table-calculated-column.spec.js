const { baseWorkbookBuffer, tamper } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

describe("github issues", () => {
    it("issue 1936 - table columns after a calculatedColumnFormula are kept", async () => {
        const source = await baseWorkbookBuffer((wb, ws) => {
            ws.addTable({
                name: "T1",
                ref: "C1",
                headerRow: true,
                columns: [{ name: "a" }, { name: "b" }, { name: "c" }],
                rows: [
                    [1, 2, 3],
                    [4, 5, 6],
                ],
            });
        });
        const buffer = await tamper(source, {
            "xl/tables/table1.xml": (xml) => {
                const patched = xml.replace(
                    /<tableColumn id="2" name="b"([^>]*)\/>/,
                    '<tableColumn id="2" name="b"$1>' +
                        "<calculatedColumnFormula>T1[[#This Row],[a]]+1</calculatedColumnFormula></tableColumn>"
                );
                expect(patched).not.to.equal(xml);
                return patched;
            },
        });

        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buffer);
        const { table } = wb.getWorksheet("Sheet1").getTable("T1");
        expect(table.columns.map((column) => column.name)).to.deep.equal(["a", "b", "c"]);
    });
});
