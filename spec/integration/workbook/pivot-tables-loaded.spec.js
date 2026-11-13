// A pivot table can be added to a worksheet read from a file (its pivotTables used to be undefined after a load).
const Excel = verquire("exceljs");

describe("Workbook", () => {
    describe("Pivot Tables of a loaded worksheet", () => {
        it("adds a pivot table to a worksheet read from a file", async () => {
            const source = new Excel.Workbook();
            const ws = source.addWorksheet("Sheet1");
            ws.addRows([
                ["A", "B", "C"],
                ["a1", "b1", 5],
                ["a2", "b2", 24],
            ]);
            source.addWorksheet("Pivot");
            const buffer = await source.xlsx.writeBuffer();

            const loaded = new Excel.Workbook();
            await loaded.xlsx.load(buffer);
            const sheet = loaded.getWorksheet("Sheet1");
            expect(sheet.pivotTables).to.deep.equal([]);

            const target = loaded.getWorksheet("Pivot");
            target.addPivotTable({ sourceSheet: sheet, rows: ["A"], columns: ["B"], values: ["C"], metric: "sum" });
            expect(target.pivotTables).to.have.length(1);
            expect(loaded.pivotTables).to.have.length(1);

            // and the workbook is written
            const written = await loaded.xlsx.writeBuffer();
            expect(written.length).to.be.greaterThan(0);
        });
    });
});
