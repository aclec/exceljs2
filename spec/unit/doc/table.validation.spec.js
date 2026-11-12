// Invalid table changes are rejected with a clear Error before anything is changed (model or cells).
const { tamper } = require("../../utils/xlsx-tamper");

const Excel = verquire("exceljs");
const colCache = verquire("utils/col-cache");

function newSheet() {
    return new Excel.Workbook().addWorksheet("s");
}

function addTable(ws, props) {
    return ws.addTable({ name: "U", ref: "E1", columns: [{ name: "A" }], rows: [[1]], ...props });
}

// the values of the cells of a range, by row
function cellValues(ws, range) {
    const { top, left, bottom, right } = colCache.decode(range);
    const values = [];
    for (let row = top; row <= bottom; row++) {
        const rowValues = [];
        for (let col = left; col <= right; col++) {
            rowValues.push(ws.getRow(row).getCell(col).value);
        }
        values.push(rowValues);
    }
    return values;
}

async function loadedTable(patches = {}) {
    const wb = new Excel.Workbook();
    addTable(wb.addWorksheet("s"), { rows: [[1], [2]] });
    const buffer = await tamper(await wb.xlsx.writeBuffer(), patches);
    const loaded = new Excel.Workbook();
    await loaded.xlsx.load(buffer);
    return loaded.getWorksheet("s").getTable("U");
}

describe("Table validation", () => {
    describe("a row with more values than the table has columns", () => {
        it("is rejected by addTable, before any cell is written", () => {
            const ws = newSheet();
            expect(() => addTable(ws, { rows: [[1], [1, 2]] })).to.throw(
                Error,
                'Table "U": row 2 has 2 values, but the table has 1 column'
            );
            expect(cellValues(ws, "E1:F3")).to.deep.equal([
                [null, null],
                [null, null],
                [null, null],
            ]);
            expect(ws.getTable("U")).to.equal(undefined);
        });

        it("is rejected by addRow, leaving the table unchanged", () => {
            const ws = newSheet();
            const table = addTable(ws);
            expect(() => table.addRow([1, 2, 3])).to.throw(Error, 'Table "U": row 2 has 3 values, but the table has 1 column');
            expect(() => table.addRow([1, 2], 0)).to.throw(Error, 'Table "U": row 1 has 2 values, but the table has 1 column');
            table.commit();
            expect(table.model.rows).to.deep.equal([[1]]);
            expect(table.model.tableRef).to.equal("E1:E2");
            expect(cellValues(ws, "E1:F3")).to.deep.equal([
                ["A", null],
                [1, null],
                [null, null],
            ]);
        });

        it("is rejected by commit when the model rows were changed directly, before any cell is cleared", () => {
            const ws = newSheet();
            const table = addTable(ws, { columns: [{ name: "A" }, { name: "B" }], rows: [[1, 2]] });
            table.removeColumns(1);
            table.model.rows.push([3, 4]);
            expect(() => table.commit()).to.throw(Error, 'Table "U": row 2 has 2 values, but the table has 1 column');
            expect(cellValues(ws, "E1:F2")).to.deep.equal([
                ["A", "B"],
                [1, 2],
            ]);
        });

        it("rejects a row that is not an array", () => {
            const ws = newSheet();
            expect(() => addTable(ws, { rows: [1] })).to.throw(Error, 'Table "U": row 1 is not an array of values');
            const table = addTable(ws);
            expect(() => table.addRow("a")).to.throw(Error, 'Table "U": row 2 is not an array of values');
            expect(table.model.rows).to.deep.equal([[1]]);
        });

        it("accepts a row with fewer values than columns", () => {
            const ws = newSheet();
            const table = addTable(ws, { columns: [{ name: "A" }, { name: "B" }], rows: [[1]] });
            table.addRow([2]);
            table.commit();
            expect(cellValues(ws, "E1:F3")).to.deep.equal([
                ["A", "B"],
                [1, null],
                [2, null],
            ]);
        });
    });

    describe("addColumn", () => {
        it("rejects more values than the table has rows, leaving the table unchanged", () => {
            const ws = newSheet();
            const table = addTable(ws);
            expect(() => table.addColumn({ name: "B" }, [1, 2, 3])).to.throw(
                Error,
                'Table "U": 3 values for a column, but the table has 1 row'
            );
            expect(table.width).to.equal(1);
            expect(table.model.rows).to.deep.equal([[1]]);
        });
    });

    describe("getColumn", () => {
        it("throws for an index outside the table", () => {
            const table = addTable(newSheet());
            expect(table.getColumn(0).name).to.equal("A");
            for (const index of [1, -1, 0.5]) {
                expect(() => table.getColumn(index)).to.throw(Error, `Table "U" has no column ${index} (0-based; it has 1 column)`);
            }
        });
    });

    describe("commit", () => {
        it("starts the next change from the table as committed", () => {
            const ws = newSheet();
            const table = addTable(ws);
            table.ref = "C3";
            table.commit();
            table.ref = "A10";
            table.commit();
            expect(cellValues(ws, "C3:C4")).to.deep.equal([[null], [null]]);
            expect(cellValues(ws, "A10:A11")).to.deep.equal([["A"], [1]]);
        });
    });

    describe("an invalid totals row function", () => {
        it("is rejected before the model is completed", () => {
            const ws = newSheet();
            const props = { columns: [{ name: "A" }, { name: "B", totalsRowFunction: "bogus" }], rows: [], totalsRow: true };
            expect(() => addTable(ws, props)).to.throw(Error, "Invalid Totals Row Function: bogus");
            expect(props.style).to.equal(undefined);
            expect(props.columns[0].totalsRowLabel).to.equal(undefined);
        });
    });

    describe("the totals row formula of a column without function", () => {
        it("is undefined (not null)", () => {
            const table = addTable(newSheet(), { columns: [{ name: "A" }, { name: "B" }], rows: [[1, 2]] });
            expect(table.getColumn(1).totalsRowFunction).to.equal("none");
            expect(table.getColumn(1).totalsRowFormula).to.equal(undefined);
        });
    });

    describe("a table read from a file", () => {
        it("has the height of its range", async () => {
            const table = await loadedTable();
            expect(table.width).to.equal(1);
            expect(table.height).to.equal(2);
        });

        it("cannot be changed by the table methods (its rows are the cells of the worksheet)", async () => {
            const table = await loadedTable();
            const message = 'Table "U" was read from a file: change the cells of its worksheet';
            expect(() => table.addRow([3])).to.throw(Error, message);
            expect(() => table.removeRows(0)).to.throw(Error, message);
            expect(() => table.addColumn({ name: "B" }, [])).to.throw(Error, message);
            expect(() => table.removeColumns(0)).to.throw(Error, message);
            expect(() => {
                table.ref = "C3";
            }).to.throw(Error, message);
            expect(() => {
                table.getColumn(0).name = "B";
            }).to.throw(Error, message);
            expect(table.model.ref).to.equal(undefined);
            expect(table.getColumn(0).name).to.equal("A");
            // nothing changed: nothing to commit
            table.commit();
        });

        it("without table style has no theme, and gets a style when one is set", async () => {
            const table = await loadedTable({ "xl/tables/table1.xml": (xml) => xml.replace(/<tableStyleInfo[^>]*\/>/, "") });
            expect(table.model.style).to.equal(null);
            expect(table.theme).to.equal(null);
            expect(table.showRowStripes).to.equal(undefined);
            table.showRowStripes = true;
            expect(table.model.style).to.deep.equal({ showRowStripes: true });
            expect(table.showRowStripes).to.equal(true);
        });
    });
});
