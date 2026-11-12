// Columns loaded from <col min max> ranges, or reached through a far cell, are created on demand: these
// tests pin the behaviour of ws.columns, getColumn and the written <cols> against the dense model.
const ExcelJS = verquire("exceljs");

function sheetFromCols(cols) {
    const ws = new ExcelJS.Workbook().addWorksheet("s");
    const { model } = ws;
    model.cols = cols;
    ws.model = model;
    return ws;
}

describe("Worksheet column ranges", () => {
    it("reads the columns of a range on demand", () => {
        const ws = sheetFromCols([{ min: 2, max: 16384, width: 20, style: { font: { bold: true } } }]);
        expect(ws.getColumn(3).width).to.equal(20);
        expect(ws.getColumn(3).font).to.deep.equal({ bold: true });
        expect(ws.getColumn(1).width).to.equal(undefined);
        expect(ws.getColumn("XFD").width).to.equal(20);
        expect(ws.columns.length).to.equal(16384);
        expect(ws.columns[1]).to.equal(ws.getColumn(2));
        expect(ws.columns.map((c) => c.number).slice(0, 3)).to.deep.equal([1, 2, 3]);
    });

    it("writes a range as one <col>, and splits it around a changed column", () => {
        const ws = sheetFromCols([{ min: 1, max: 16384, width: 20 }]);
        expect(ws.model.cols).to.deep.equal([
            {
                min: 1,
                max: 16384,
                width: 20,
                style: {},
                isCustomWidth: true,
                hidden: false,
                outlineLevel: 0,
                collapsed: false,
            },
        ]);
        ws.getColumn(5).width = 30;
        expect(ws.model.cols.map(({ min, max, width }) => [min, max, width])).to.deep.equal([
            [1, 4, 20],
            [5, 5, 30],
            [6, 16384, 20],
        ]);
    });

    it("keeps the first of overlapping ranges, and gaps before a range", () => {
        const ws = sheetFromCols([
            { min: 5, max: 20, width: 30 },
            { min: 3, max: 10, width: 20 },
            { min: 40, max: 30, width: 50 },
        ]);
        expect(ws.columns.length).to.equal(39);
        expect([2, 3, 10, 11, 20, 21].map((n) => ws.getColumn(n).width)).to.deep.equal([undefined, 20, 20, 30, 30, undefined]);
        expect(ws.model.cols.map(({ min, max, width }) => [min, max, width])).to.deep.equal([
            [3, 10, 20],
            [11, 20, 30],
        ]);
    });

    it("does not join columns across a column that was never created", () => {
        const ws = new ExcelJS.Workbook().addWorksheet("s");
        ws.getColumn(1).width = 20;
        ws.getColumn(3).width = 20;
        expect(ws.model.cols.map(({ min, max }) => [min, max])).to.deep.equal([
            [1, 1],
            [3, 3],
        ]);
        expect(ws.columns.length).to.equal(3);
        expect(ws.columns[1].number).to.equal(2);
    });

    it("keeps the columns array returned earlier", () => {
        const ws = new ExcelJS.Workbook().addWorksheet("s");
        ws.columns = [{ key: "a", width: 10 }, { key: "b" }];
        const { columns } = ws;
        ws.getCell("E1").value = 1;
        expect(ws.columns).to.equal(columns);
        expect(columns.map((c) => c.number)).to.deep.equal([1, 2, 3, 4, 5]);
    });

    it("splices columns loaded as a range", () => {
        const ws = sheetFromCols([{ min: 1, max: 3, width: 20 }]);
        ws.spliceColumns(2, 1);
        expect(ws.columns.map((c) => c.width)).to.deep.equal([20, 20, 9, undefined]);
        expect(ws.model.cols.map(({ min, max, width }) => [min, max, width])).to.deep.equal([[1, 2, 20]]);
    });
});
