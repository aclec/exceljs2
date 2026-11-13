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

        it("breaks on one column only when lft and rght are the same, the first one", () => {
            const ws = new Excel.Workbook().addWorksheet("blort");
            ws.getRow(3).addPageBreak(1, 1);
            ws.getRow(4).addPageBreak(2, 2);
            ws.getRow(5).addPageBreak(0, 1);
            ws.getRow(6).addPageBreak(3);
            expect(ws.rowBreaks).to.deep.equal([
                { id: 3, max: 0, man: 1 },
                { id: 4, max: 1, min: 1, man: 1 },
                { id: 5, max: 0, man: 1 },
                { id: 6, max: 16383, min: 2, man: 1 },
            ]);
        });

        it("rejects a row break whose first column is after its last one", () => {
            const ws = new Excel.Workbook().addWorksheet("blort");
            expect(() => ws.getRow(3).addPageBreak(10, 5)).to.throw(
                Error,
                "Invalid page break columns: 10 to 5 (the first column is after the last one)"
            );
            expect(ws.rowBreaks).to.deep.equal([]);
        });

        it("adds column breaks across all the rows by default (max: 1048575, the last 0-based row)", () => {
            const ws = new Excel.Workbook().addWorksheet("blort");
            ws.getColumn(3).addPageBreak();
            ws.getColumn("F").addPageBreak(5, 10);
            ws.getColumn(16384).addPageBreak(1, 1048576);
            expect(ws.colBreaks).to.deep.equal([
                { id: 3, max: 1048575, man: 1 },
                { id: 6, max: 9, min: 4, man: 1 },
                { id: 16384, max: 1048575, man: 1 },
            ]);
            expect(ws.rowBreaks).to.deep.equal([]);
            expect(ws.model.colBreaks).to.equal(ws.colBreaks);
        });

        it("breaks on one row only when top and bottom are the same, the first one", () => {
            const ws = new Excel.Workbook().addWorksheet("blort");
            ws.getColumn(2).addPageBreak(1, 1);
            ws.getColumn(3).addPageBreak(7, 7);
            ws.getColumn(4).addPageBreak(0, 1);
            ws.getColumn(5).addPageBreak(3);
            expect(ws.colBreaks).to.deep.equal([
                { id: 2, max: 0, man: 1 },
                { id: 3, max: 6, min: 6, man: 1 },
                { id: 4, max: 0, man: 1 },
                { id: 5, max: 1048575, min: 2, man: 1 },
            ]);
        });

        it("rejects a column break whose first row is after its last one", () => {
            const ws = new Excel.Workbook().addWorksheet("blort");
            expect(() => ws.getColumn(3).addPageBreak(10, 5)).to.throw(
                Error,
                "Invalid page break rows: 10 to 5 (the first row is after the last one)"
            );
            expect(ws.colBreaks).to.deep.equal([]);
        });

        it("rejects a row out of the worksheet for a column break", () => {
            const ws = new Excel.Workbook().addWorksheet("blort");
            for (const [top, bottom] of [
                [1, 1048577],
                [-1, 2],
                [1.5, 2],
                [Number.NaN, 2],
            ]) {
                expect(() => ws.getColumn(3).addPageBreak(top, bottom)).to.throw(Error, /^Invalid page break row/);
            }
            expect(ws.colBreaks).to.deep.equal([]);
        });

        it("copies the column breaks with the model, none when the model has none", () => {
            const wb = new Excel.Workbook();
            const ws = wb.addWorksheet("a");
            ws.getColumn(2).addPageBreak();
            const copy = wb.addWorksheet("b");
            copy.model = { ...ws.model, name: "b" };
            expect(copy.colBreaks).to.deep.equal([{ id: 2, max: 1048575, man: 1 }]);

            const { colBreaks, ...rest } = ws.model;
            const other = wb.addWorksheet("c");
            other.model = { ...rest, name: "c" };
            expect(colBreaks).to.have.length(1);
            expect(other.colBreaks).to.deep.equal([]);
        });

        it("keeps the 4.x pageSetup.colBreaks (null): the breaks are worksheet.colBreaks", () => {
            const ws = new Excel.Workbook().addWorksheet("blort");
            expect(ws.pageSetup.colBreaks).to.equal(null);
            ws.getColumn(2).addPageBreak();
            expect(ws.pageSetup.colBreaks).to.equal(null);
        });
    });
});
