const Excel = verquire("exceljs");

describe("Worksheet", () => {
    describe("Page Breaks", () => {
        it("adds multiple row breaks", () => {
            const wb = new Excel.Workbook();
            const ws = wb.addWorksheet("blort");

            // initial values
            ws.getCell("A1").value = "A1";
            ws.getCell("B1").value = "B1";
            ws.getCell("A2").value = "A2";
            ws.getCell("B2").value = "B2";
            ws.getCell("A3").value = "A3";
            ws.getCell("B3").value = "B3";

            let row = ws.getRow(1);
            row.addPageBreak();
            row = ws.getRow(2);
            row.addPageBreak();

            expect(ws.rowBreaks.length).to.equal(2);
        });

        it("breaks across all the columns by default (max: 16383, the last 0-based column)", () => {
            const ws = new Excel.Workbook().addWorksheet("blort");
            ws.getRow(3).addPageBreak();
            ws.getRow(5).addPageBreak(2, 4);
            expect(ws.rowBreaks).to.deep.equal([
                { id: 3, max: 16383, man: 1 },
                { id: 5, max: 3, min: 1, man: 1 },
            ]);
        });

        it("rejects a column out of the worksheet", () => {
            const ws = new Excel.Workbook().addWorksheet("blort");
            for (const [lft, rght] of [
                [1, 16385],
                [-1, 2],
                [1.5, 2],
                [Number.NaN, 2],
            ]) {
                expect(() => ws.getRow(3).addPageBreak(lft, rght)).to.throw(Error, /^Invalid page break column/);
            }
            expect(ws.rowBreaks).to.deep.equal([]);
        });
    });
});
