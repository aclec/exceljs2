// spliceRows / spliceColumns are checked against Array.prototype.splice on the same data.
const Excel = verquire("exceljs");

const SIZE = 10;

function rowsOf(ws) {
    const rows = [];
    ws.eachRow((row, rowNumber) => {
        rows[rowNumber - 1] = row.getCell(1).value;
    });
    return Array.from(rows, (value) => (value === undefined ? null : value));
}

function sheetWithRows() {
    const ws = new Excel.Workbook().addWorksheet("s");
    for (let r = 1; r <= SIZE; r++) {
        ws.getCell(r, 1).value = `r${r}`;
        ws.getCell(r, 1).font = { name: `r${r}` };
    }
    return ws;
}

function sheetWithColumns() {
    const ws = new Excel.Workbook().addWorksheet("s");
    const values = [];
    for (let c = 1; c <= SIZE; c++) {
        values.push(`c${c}`);
    }
    ws.addRow(values);
    ws.addRow(values);
    return ws;
}

describe("Worksheet splice", () => {
    // [start, count, number of inserted rows/columns]
    const cases = [];
    for (let start = 1; start <= SIZE + 1; start++) {
        for (let count = 0; count <= SIZE - start + 3; count++) {
            for (let nInserts = 0; nInserts <= 3; nInserts++) {
                if (count || nInserts) cases.push([start, count, nInserts]);
            }
        }
    }

    describe("spliceRows", () => {
        it("matches Array#splice for every start/count/inserts combination", () => {
            cases.forEach(([start, count, nInserts]) => {
                const ws = sheetWithRows();
                const expected = Array.from({ length: SIZE }, (v, i) => `r${i + 1}`);
                const inserts = Array.from({ length: nInserts }, (v, i) => [`new${i}`]);
                expected.splice(start - 1, count, ...inserts.map(([value]) => value));

                ws.spliceRows(start, count, ...inserts);

                const label = `spliceRows(${start}, ${count}, ${nInserts} inserts)`;
                expect(rowsOf(ws), label).to.deep.equal(expected);
                expect(ws.rowCount, label).to.equal(expected.length);
                expect(ws.actualRowCount, label).to.equal(expected.length);
                expect(ws.lastRow ? ws.lastRow.number : 0, label).to.equal(expected.length);
                expected.forEach((value, index) => {
                    const font = value.startsWith("r") ? { name: value } : undefined;
                    expect(ws.getCell(index + 1, 1).font, label).to.deep.equal(font);
                });
            });
        });

        it("removes the last two rows (spliceRows(9, 2) of 10)", () => {
            const ws = sheetWithRows();
            ws.spliceRows(9, 2);
            expect(ws.findRow(9)).to.equal(undefined);
            expect(ws.findRow(10)).to.equal(undefined);
            expect(ws.rowCount).to.equal(8);
            expect(ws.lastRow.getCell(1).value).to.equal("r8");
        });

        it("removes a few rows near the end without leaving stale rows", () => {
            const ws = sheetWithRows();
            ws.spliceRows(2, 5);
            expect(rowsOf(ws)).to.deep.equal(["r1", "r7", "r8", "r9", "r10"]);
            expect(ws.findRow(6)).to.equal(undefined);
        });

        it("does not grow the sheet when deleting past the end", () => {
            const ws = sheetWithRows();
            ws.spliceRows(9, 50);
            expect(ws._rows.length).to.be.at.most(SIZE);
            expect(ws.lastRow.number).to.equal(8);
            let count = 0;
            ws.eachRow({ includeEmpty: true }, () => count++);
            expect(count).to.equal(8);
        });

        it("can insert and delete repeatedly", () => {
            const ws = sheetWithRows();
            const expected = Array.from({ length: SIZE }, (v, i) => `r${i + 1}`);
            [
                [3, 2, ["a"], ["b"], ["c"]],
                [10, 3],
                [1, 1, ["d"]],
                [8, 0, ["e"], ["f"]],
                [7, 4],
            ].forEach(([start, count, ...inserts]) => {
                ws.spliceRows(start, count, ...inserts);
                expected.splice(start - 1, count, ...inserts.map(([value]) => value));
                expect(rowsOf(ws)).to.deep.equal(expected);
                expect(ws.rowCount).to.equal(expected.length);
            });
        });

        it("moves notes with their rows", () => {
            const ws = new Excel.Workbook().addWorksheet("s");
            ws.getCell("A3").value = "v3";
            ws.getCell("A3").note = "n3";
            ws.getCell("B4").note = "n4"; // a note on a cell without a value
            ws.getCell("A6").note = "n6";

            ws.insertRows(2, [["x"], ["y"]]);
            expect(ws.getCell("A5").note).to.equal("n3");
            expect(ws.getCell("B6").note).to.equal("n4");
            expect(ws.getCell("A8").note).to.equal("n6");
            ["A3", "B4", "A6"].forEach((address) => expect(ws.getCell(address).note, address).to.equal(undefined));

            ws.spliceRows(1, 5, ["z"]);
            expect(ws.getCell("B2").note).to.equal("n4");
            expect(ws.getCell("A4").note).to.equal("n6");
            ["A5", "B6", "A8"].forEach((address) => expect(ws.getCell(address).note, address).to.equal(undefined));
        });
    });

    describe("with far rows", () => {
        const FAR = 100000;

        function sheetWithFarRow() {
            const ws = new Excel.Workbook().addWorksheet("s");
            ws.getCell("A1").value = "r1";
            ws.getCell("A3").value = "r3";
            ws.getCell(`A${FAR}`).value = "far";
            ws.getCell(`A${FAR}`).font = { name: "far" };
            ws.getCell(`A${FAR}`).note = "note";
            return ws;
        }

        function existingRows(ws) {
            return Object.keys(ws._rows)
                .filter((key) => ws._rows[key])
                .map((key) => [Number(key) + 1, ws._rows[key].getCell(1).value]);
        }

        it("moves the rows after the splice, and only them", () => {
            const ws = sheetWithFarRow();
            ws.spliceRows(2, 1);
            expect(existingRows(ws)).to.deep.equal([
                [1, "r1"],
                [2, "r3"],
                [FAR - 1, "far"],
            ]);
            ws.spliceRows(2, 0, ["a"], ["b"]);
            expect(existingRows(ws)).to.deep.equal([
                [1, "r1"],
                [2, "a"],
                [3, "b"],
                [4, "r3"],
                [FAR + 1, "far"],
            ]);
            expect(ws.getCell(`A${FAR + 1}`).font).to.deep.equal({ name: "far" });
            expect(ws.getCell(`A${FAR + 1}`).note).to.equal("note");
            expect([ws.rowCount, ws.actualRowCount, ws.lastRow.number]).to.deep.equal([FAR + 1, 5, FAR + 1]);
        });

        it("removes the far row and shortens the sheet", () => {
            const ws = sheetWithFarRow();
            ws.spliceRows(50, FAR);
            expect(existingRows(ws)).to.deep.equal([
                [1, "r1"],
                [3, "r3"],
            ]);
            expect([ws.rowCount, ws._rows.length, ws.lastRow.number]).to.deep.equal([3, 3, 3]);
        });

        it("visits the rows in order", () => {
            const ws = sheetWithFarRow();
            const visited = [];
            ws.eachRow((row, rowNumber) => visited.push([rowNumber, row.number]));
            expect(visited).to.deep.equal([
                [1, 1],
                [3, 3],
                [FAR, FAR],
            ]);
            expect(ws.dimensions.model).to.deep.equal({ top: 1, left: 1, bottom: FAR, right: 1 });
            const values = ws.getSheetValues();
            expect([values.length, values[3], values[FAR]]).to.deep.equal([FAR + 1, [undefined, "r3"], [undefined, "far"]]);
            expect(ws.model.rows.map((row) => row.number)).to.deep.equal([1, 3, FAR]);
        });

        it("inserts columns without creating the rows that get no value", () => {
            const ws = sheetWithFarRow();
            ws.spliceColumns(1, 0, [undefined, "x"]);
            expect(existingRows(ws).map(([rowNumber]) => rowNumber)).to.deep.equal([1, 2, 3, FAR]);
            expect([1, 2, 3, FAR].map((r) => [ws.getCell(r, 1).value, ws.getCell(r, 2).value])).to.deep.equal([
                [null, "r1"],
                ["x", null],
                [null, "r3"],
                [null, "far"],
            ]);
            expect(ws.getCell(FAR, 2).font).to.deep.equal({ name: "far" });
        });

        it("inserts the values given for rows past the last one, and keeps falsy values", () => {
            const ws = new Excel.Workbook().addWorksheet("s");
            ws.spliceColumns(2, 0, ["a", 0, false, ""]);
            expect([1, 2, 3, 4].map((r) => ws.getCell(r, 2).value)).to.deep.equal(["a", 0, false, ""]);

            const filled = sheetWithFarRow();
            filled.spliceColumns(1, 0, [0, false]);
            expect([1, 2].map((r) => filled.getCell(r, 1).value)).to.deep.equal([0, false]);
            expect(filled.getCell(1, 2).value).to.equal("r1");
        });

        it("duplicates a row that does not exist as empty rows", () => {
            const ws = sheetWithFarRow();
            ws.duplicateRow(2, 2, true);
            expect(existingRows(ws)).to.deep.equal([
                [1, "r1"],
                [2, null],
                [3, null],
                [4, null],
                [5, "r3"],
                [FAR + 2, "far"],
            ]);
            expect(ws.actualRowCount).to.equal(3);
        });
    });

    describe("spliceColumns", () => {
        it("matches Array#splice for every start/count/inserts combination", () => {
            cases.forEach(([start, count, nInserts]) => {
                const ws = sheetWithColumns();
                const expected = Array.from({ length: SIZE }, (v, i) => `c${i + 1}`);
                const inserts = Array.from({ length: nInserts }, (v, i) => [`new${i}`, `new${i}`]);
                expected.splice(start - 1, count, ...inserts.map(([value]) => value));

                ws.spliceColumns(start, count, ...inserts);

                const label = `spliceColumns(${start}, ${count}, ${nInserts} inserts)`;
                [1, 2].forEach((r) => {
                    const values = ws.getRow(r).values.slice(1);
                    expect(
                        values.filter((v) => v !== null),
                        label
                    ).to.deep.equal(expected);
                    expect(values.slice(0, expected.length), label).to.deep.equal(expected);
                });
                expect(ws.actualColumnCount, label).to.equal(expected.length);
            });
        });

        it("moves column keys with their columns", () => {
            const ws = new Excel.Workbook().addWorksheet("s");
            ws.columns = ["a", "b", "c", "d"].map((key) => ({ key, header: key }));

            ws.spliceColumns(2, 1);
            expect(ws.getColumn("c").number).to.equal(2);
            expect(ws.getColumn("d").number).to.equal(3);
            expect(ws.getColumnKey("b")).to.equal(undefined);

            ws.spliceColumns(3, 1);
            expect(ws.getColumnKey("d")).to.equal(undefined);
            expect(ws.getColumn("c").number).to.equal(2);

            ws.spliceColumns(2, 0, ["x"]);
            expect(ws.getColumn("c").number).to.equal(3);
            expect(ws.getRow(1).values.slice(1)).to.deep.equal(["a", "x", "c"]);
        });
    });
});
