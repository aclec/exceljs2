// Merged cells must move with the rows / columns spliced around them: the merge index (worksheet._merges, which the
// writers use), the cells' merge values and their masters must stay consistent.
const colCache = verquire("utils/col-cache");
const Range = verquire("doc/range");

const Excel = verquire("exceljs");

const ROW_MERGES = { "B2:C3": "m1", "E5:E8": "m2", "A10:D10": "m3" };
const COL_MERGES = { "B2:C3": "m1", "E1:G1": "m2", "A5:A9": "m3" };

function buildSheet(merges) {
    const ws = new Excel.Workbook().addWorksheet("sheet");
    for (let row = 1; row <= 12; row++) {
        ws.getRow(row).values = ["a", "b", "c", "d", "e", "f", "g"].map((col) => `${col}${row}`);
    }
    Object.entries(merges).forEach(([range, value]) => {
        ws.mergeCells(range);
        ws.getCell(range.split(":")[0]).value = value;
    });
    return ws;
}

// every merge of the index has its master at the top-left and every other cell of its area merged into that very
// cell (not a detached copy); no cell outside the index is a merge slave
function checkConsistency(ws) {
    const slaves = new Set();
    Object.entries(ws._merges).forEach(([key, range]) => {
        expect(key, "merge key").to.equal(range.tl);
        const master = ws.findCell(range.top, range.left);
        expect(master, `master of ${range.range}`).to.not.equal(undefined);
        expect(master.type, `master of ${range.range}`).to.not.equal(Excel.ValueType.Merge);
        for (let row = range.top; row <= range.bottom; row++) {
            for (let col = range.left; col <= range.right; col++) {
                if (row !== range.top || col !== range.left) {
                    const cell = ws.findCell(row, col);
                    const address = colCache.encodeAddress(row, col);
                    expect(cell && cell.isMergedTo(master), `${address} merged into ${range.tl}`).to.equal(true);
                    slaves.add(address);
                }
            }
        }
    });
    ws.eachRow({ includeEmpty: true }, (row) => {
        row.eachCell((cell) => {
            if (cell.type === Excel.ValueType.Merge) {
                expect(slaves.has(cell.address), `${cell.address} is a merge slave outside any merge`).to.equal(true);
            }
        });
    });
}

function mergeRanges(ws) {
    return ws.model.merges.slice().sort();
}

const row = (n) => [`x${n}`];
// a column of values for rows 1 to 12, without values in the rows of emptyRows
const column = (n, emptyRows = []) => Array.from({ length: 12 }, (_, i) => (emptyRows.includes(i + 1) ? null : `y${n}${i + 1}`));

describe("Worksheet merges after splicing", () => {
    describe("rows", () => {
        const cases = [
            { name: "insertRow above every merge", op: (ws) => ws.insertRow(1, row(1)), merges: ["A11:D11", "B3:C4", "E6:E9"] },
            { name: "insertRow inside a merge expands it", op: (ws) => ws.insertRow(3, row(1)), merges: ["A11:D11", "B2:C4", "E6:E9"] },
            { name: "insertRow between merges", op: (ws) => ws.insertRow(4, row(1)), merges: ["A11:D11", "B2:C3", "E6:E9"] },
            {
                name: "insertRows of several rows",
                op: (ws) => ws.insertRows(5, [row(1), row(2)]),
                merges: ["A12:D12", "B2:C3", "E7:E10"],
            },
            { name: "delete the master row of a merge", op: (ws) => ws.spliceRows(2, 1), merges: ["A9:D9", "B2:C2", "E4:E7"] },
            { name: "delete rows inside a merge", op: (ws) => ws.spliceRows(6, 2), merges: ["A8:D8", "B2:C3", "E5:E6"] },
            { name: "delete a whole single-row merge", op: (ws) => ws.spliceRows(10, 1), merges: ["B2:C3", "E5:E8"] },
            { name: "delete a whole merge", op: (ws) => ws.spliceRows(2, 2), merges: ["A8:D8", "E3:E6"] },
            {
                name: "delete across two merges",
                op: (ws) => ws.spliceRows(3, 5, row(1)),
                // E5:E8 keeps a single cell (E8 -> E4): not a merge any more
                merges: ["A6:D6", "B2:C2"],
            },
            {
                name: "delete and insert more rows (net offset)",
                op: (ws) => ws.spliceRows(2, 1, row(1), row(2), row(3)),
                merges: ["A12:D12", "B5:C5", "E7:E10"],
            },
            {
                name: "delete and insert inside a merge",
                op: (ws) => ws.spliceRows(6, 1, row(1), row(2)),
                merges: ["A11:D11", "B2:C3", "E5:E9"],
            },
            {
                name: "delete and insert fewer rows",
                op: (ws) => ws.spliceRows(1, 3, row(1)),
                merges: ["A8:D8", "E3:E6"],
            },
            { name: "replace rows one for one", op: (ws) => ws.spliceRows(4, 1, row(1)), merges: ["A10:D10", "B2:C3", "E5:E8"] },
            { name: "splice nothing", op: (ws) => ws.spliceRows(4, 0), merges: ["A10:D10", "B2:C3", "E5:E8"] },
            {
                name: "duplicateRow of a single-row merge (insert)",
                op: (ws) => ws.duplicateRow(10, 2, true),
                merges: ["A10:D10", "A11:D11", "A12:D12", "B2:C3", "E5:E8"],
            },
            {
                name: "duplicateRow inside a merge (insert)",
                op: (ws) => ws.duplicateRow(2, 1, true),
                merges: ["A11:D11", "B2:C4", "E6:E9"],
            },
            {
                name: "duplicateRow over a merge (overwrite)",
                op: (ws) => ws.duplicateRow(1, 2, false),
                merges: ["A10:D10", "E5:E8"],
            },
        ];

        cases.forEach(({ name, op, merges }) => {
            it(name, () => {
                const ws = buildSheet(ROW_MERGES);
                op(ws);
                expect(mergeRanges(ws)).to.deep.equal(merges);
                checkConsistency(ws);
            });
        });

        it("keeps the master value and style with the moved merge", () => {
            const ws = buildSheet(ROW_MERGES);
            const border = { bottom: { style: "thick" } };
            ws.getCell("E5").border = border;
            ws.spliceRows(2, 1, row(1), row(2), row(3));
            expect(ws.getCell("E7").value).to.equal("m2");
            expect(ws.getCell("E10").value).to.equal("m2");
            expect(ws.getCell("E10").border).to.deep.equal(border);
            expect(ws.getCell("A12").value).to.equal("m3");
            // the merge whose master row was deleted keeps the next row's (slave, empty) cells
            expect(ws.getCell("B5").value).to.equal(null);
        });

        it("lets moved merges be unmerged and merged again", () => {
            const ws = buildSheet(ROW_MERGES);
            ws.insertRows(1, [row(1), row(2)]);
            ws.unMergeCells("E7");
            ws.unMergeCells("B4");
            expect(mergeRanges(ws)).to.deep.equal(["A12:D12"]);
            ws.mergeCells("B4:E10");
            checkConsistency(ws);
        });

        it("writes and reads back the moved merges", async () => {
            const ws = buildSheet(ROW_MERGES);
            ws.spliceRows(2, 1, row(1), row(2), row(3));
            const buffer = await ws.workbook.xlsx.writeBuffer();
            const wb = new Excel.Workbook();
            await wb.xlsx.load(buffer);
            const reloaded = wb.getWorksheet("sheet");
            expect(mergeRanges(reloaded)).to.deep.equal(["A12:D12", "B5:C5", "E7:E10"]);
            expect(reloaded.getCell("E9").value).to.equal("m2");
        });
    });

    describe("columns", () => {
        const cases = [
            {
                name: "insert a column before every merge",
                op: (ws) => ws.spliceColumns(1, 0, column(1)),
                merges: ["B5:B9", "C2:D3", "F1:H1"],
            },
            {
                name: "insert an empty column inside a merge",
                op: (ws) => ws.spliceColumns(3, 0, column(1, [2, 3])),
                merges: ["A5:A9", "B2:D3", "F1:H1"],
            },
            { name: "delete the master column of a merge", op: (ws) => ws.spliceColumns(2, 1), merges: ["A5:A9", "B2:B3", "D1:F1"] },
            {
                name: "delete and insert more columns inside a merge",
                op: (ws) => ws.spliceColumns(6, 1, column(1, [1]), column(2, [1])),
                merges: ["A5:A9", "B2:C3", "E1:H1"],
            },
            { name: "delete a whole merge", op: (ws) => ws.spliceColumns(1, 1), merges: ["A2:B3", "D1:F1"] },
        ];

        cases.forEach(({ name, op, merges }) => {
            it(name, () => {
                const ws = buildSheet(COL_MERGES);
                op(ws);
                expect(mergeRanges(ws)).to.deep.equal(merges);
                checkConsistency(ws);
            });
        });

        it("keeps the master value with the moved merge", () => {
            const ws = buildSheet(COL_MERGES);
            ws.spliceColumns(1, 0, column(1));
            expect(ws.getCell("C2").value).to.equal("m1");
            expect(ws.getCell("D3").value).to.equal("m1");
            expect(ws.getCell("H1").value).to.equal("m2");
        });
    });

    describe("values inserted inside a merge", () => {
        // Excel widens a merge over the cells inserted inside it, and refuses to change part of a merged cell: a value
        // inserted there would be lost, so the splice throws before anything changes
        function expectUnchanged(ws) {
            expect(mergeRanges(ws)).to.deep.equal(Object.keys(ROW_MERGES).sort());
            expect(ws.getCell("B2").value).to.equal("m1");
            expect(ws.getCell("A3").value).to.equal("a3");
            expect(ws.getCell("C3").isMergedTo(ws.getCell("B2"))).to.equal(true);
            checkConsistency(ws);
        }

        it("insertRow refuses values in the merged columns", () => {
            const ws = buildSheet(ROW_MERGES);
            expect(() => ws.insertRow(3, ["a", "b-inserted", "c-inserted"])).to.throw(/C3.*B2:C4|B3.*B2:C4/);
            expectUnchanged(ws);
        });

        it("insertRows refuses values given by column key", () => {
            const ws = buildSheet(ROW_MERGES);
            ws.columns = [{ key: "a" }, { key: "b" }, { key: "c" }];
            expect(() => ws.insertRows(3, [{ a: 1 }, { a: 2, c: "lost" }])).to.throw(/C4.*B2:C5/);
            expectUnchanged(ws);
        });

        it("resolves column keys as row.values does", () => {
            // keys of columns that were replaced stay in use: row.values writes them, so the check must see them
            let ws = buildSheet(ROW_MERGES);
            ws.columns = [{ key: "a" }, { key: "b" }, { key: "c" }];
            ws.columns = [{ key: "x" }, { key: "y" }, { key: "z" }];
            expect(() => ws.insertRow(3, { b: "lost" })).to.throw(/B3.*B2:C4/);
            expectUnchanged(ws);

            // a key moved to another column is written to that column only
            ws = buildSheet(ROW_MERGES);
            ws.getColumn(2).key = "k";
            ws.getColumn(4).key = "k";
            ws.insertRow(3, { k: "v" });
            expect(ws.getCell("D3").value).to.equal("v");
            expect(ws.getCell("B3").isMergedTo(ws.getCell("B2"))).to.equal(true);
            expect(ws.getCell("C3").isMergedTo(ws.getCell("B2"))).to.equal(true);
            checkConsistency(ws);
        });

        it("spliceRows refuses values that replace rows inside a merge", () => {
            const ws = buildSheet(ROW_MERGES);
            expect(() => ws.spliceRows(6, 1, row(1), ["x", "y", "z", "w", "lost"])).to.throw(/E7.*E5:E9/);
            expect(mergeRanges(ws)).to.deep.equal(Object.keys(ROW_MERGES).sort());
            expect(ws.getCell("A6").value).to.equal("a6");
        });

        it("still widens the merge when the inserted cells of the merge are empty", () => {
            const ws = buildSheet(ROW_MERGES);
            ws.insertRows(3, [
                ["a", null, ""],
                [undefined, undefined, undefined, "d"],
            ]);
            expect(mergeRanges(ws)).to.deep.equal(["A12:D12", "B2:C5", "E7:E10"]);
            expect(ws.getCell("A3").value).to.equal("a");
            expect(ws.getCell("D4").value).to.equal("d");
            checkConsistency(ws);
        });

        it("spliceColumns refuses values in the merged rows", () => {
            const ws = buildSheet(COL_MERGES);
            expect(() => ws.spliceColumns(3, 0, column(1, [2]))).to.throw(/C3.*B2:D3/);
            expect(mergeRanges(ws)).to.deep.equal(Object.keys(COL_MERGES).sort());
            expect(ws.getCell("C5").value).to.equal("c5");
            checkConsistency(ws);
        });
    });

    it("duplicateRow copies only the merges of the row as it was (overwrite)", () => {
        const ws = buildSheet({ "C3:D4": "m" });
        ws.duplicateRow(3, 2, false);
        // C3:D4 lost its second row to the overwrite: it was not a single-row merge of the duplicated row
        expect(mergeRanges(ws)).to.deep.equal(["C3:D3"]);
        checkConsistency(ws);
    });

    it("refuses a splice that would exceed the merged-cell cap, before changing anything", function () {
        this.timeout(30000);
        const ws = new Excel.Workbook().addWorksheet("sheet");
        ws.mergeCells("A1:A1040000"); // the cap is 1,048,576 merged cells
        ws.getCell("A1").value = "big";
        ws.mergeCells("D2:E3");
        ws.getCell("D6").value = "below";
        expect(() => ws.spliceRows(3, 0, ...Array.from({ length: 10000 }, () => ["x"]))).to.throw(/merged cells/);
        expect(Object.values(ws._merges).map((range) => range.range)).to.deep.equal(["A1:A1040000", "D2:E3"]);
        expect(ws.getCell("D6").value).to.equal("below");
        expect(ws.getCell("A5").isMergedTo(ws.getCell("A1"))).to.equal(true);
    });

    it("keeps the merged-cell count in step", () => {
        const ws = buildSheet(ROW_MERGES);
        ws.spliceRows(3, 5, row(1));
        const count = Object.values(ws._merges).reduce((sum, range) => sum + new Range(range.range).count, 0);
        expect(ws._mergedCellCount).to.equal(count);
    });
});
