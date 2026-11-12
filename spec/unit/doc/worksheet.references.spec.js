const Excel = verquire("exceljs");
const Range = verquire("doc/range");

// a worksheet named "s"
function sheet() {
    return new Excel.Workbook().addWorksheet("s");
}

// the indexes of the rows that exist (a row "0" was stored at index -1)
function rowKeys(ws) {
    return Object.keys(ws._rows);
}

describe("Worksheet references", () => {
    describe("mergeCells", () => {
        it("drops the worksheet's own sheet name from a merge", () => {
            const ws = sheet();
            ws.mergeCells("s!A1:B2");
            ws.mergeCells(8, 1, 9, 2, "s");
            ws.mergeCells({ top: 5, left: 1, bottom: 6, right: 2, sheetName: "S" });
            ws.mergeCells("'s'!D1", "E2");
            ws.mergeCells(new Range("s!G1:H2"));
            expect(ws.model.merges).to.deep.equal(["A1:B2", "A8:B9", "A5:B6", "D1:E2", "G1:H2"]);
            expect(ws.getCell("B2").isMerged).to.be.true();
        });

        it("rejects a merge on another sheet, merging nothing", () => {
            const cases = [
                ["y!A1:B2"],
                [8, 1, 9, 2, "y"],
                [{ top: 5, left: 1, bottom: 6, right: 2, sheetName: "y" }],
                ["A1", "B2", "y"],
                ["y!A1", "B2"],
                [new Range("y!A1:B2")],
                [["y!A1:B2"]],
            ];
            for (const args of cases) {
                const ws = sheet();
                expect(() => ws.mergeCells(...args), JSON.stringify(args)).to.throw(Error, /on the worksheet "y", not "s"/);
                expect(() => ws.mergeCellsWithoutStyle(...args), JSON.stringify(args)).to.throw(Error, /on the worksheet "y"/);
                expect(ws.model.merges).to.deep.equal([]);
                expect(rowKeys(ws)).to.deep.equal([]);
            }
        });

        it("rejects invalid merge arguments, merging nothing", () => {
            const cases = [
                [],
                [10, 10],
                ["foo"],
                ["A1:foo"],
                ["a1:b2"],
                ["A0:B2"],
                ["XFE1"],
                ["A1048577"],
                [0, 1, 2, 2],
                [1, 1, 2, 16385],
                [1.5, 1, 2, 2],
                [Number.NaN, 1, 2, 2],
                ["A1", 2],
                ["A1:B2", "C3"],
                [{ top: 1, left: 1 }],
                [null],
                [new Range()],
                [new Range("Sheet1!#REF!")],
                [[]],
            ];
            for (const args of cases) {
                const ws = sheet();
                expect(() => ws.mergeCells(...args), JSON.stringify(args)).to.throw(Error, /^Cannot merge: invalid range/);
                expect(ws.model.merges).to.deep.equal([]);
                expect(rowKeys(ws)).to.deep.equal([]);
            }
        });

        it("keeps the valid forms of a range", () => {
            const ws = sheet();
            ws.mergeCells("A1:B2");
            ws.mergeCells("$D$1:$E$2");
            ws.mergeCells("H2:G1");
            ws.mergeCells("I1", "J2");
            ws.mergeCells("L2", "K1", "s");
            ws.mergeCells(1, 13, 2, 14);
            ws.mergeCells(["O1", "P2"]);
            ws.mergeCells([1, 17, 2, 18]);
            ws.mergeCells({ top: 1, left: 19, bottom: 2, right: 20 });
            ws.mergeCells(new Range("U1:V2"));
            ws.mergeCells("W1");
            expect(ws.model.merges).to.deep.equal([
                "A1:B2",
                "D1:E2",
                "G1:H2",
                "I1:J2",
                "K1:L2",
                "M1:N2",
                "O1:P2",
                "Q1:R2",
                "S1:T2",
                "U1:V2",
                "W1:W1",
            ]);
        });

        it("unmerges with the worksheet's own sheet name only", () => {
            const ws = sheet();
            ws.mergeCells("A1:B2");
            ws.mergeCells("D1:E2");
            expect(() => ws.unMergeCells("y!A1")).to.throw(Error, /^Cannot unmerge "y!A1": the range is on the worksheet "y", not "s"$/);
            expect(() => ws.unMergeCells()).to.throw(Error, /^Cannot unmerge: invalid range/);
            expect(() => ws.unMergeCells("A0")).to.throw(Error, /^Cannot unmerge: invalid range/);
            expect(ws.model.merges).to.deep.equal(["A1:B2", "D1:E2"]);
            ws.unMergeCells("s!B2");
            expect(ws.model.merges).to.deep.equal(["D1:E2"]);
        });
    });

    describe("fillFormula", () => {
        it("drops the worksheet's own sheet name from the formula range", () => {
            const ws = sheet();
            ws.fillFormula("s!A1:A2", "B1", [1, 2]);
            expect(ws.getCell("A1").value).to.deep.equal({ shareType: "shared", formula: "B1", ref: "A1:A2", result: 1 });
            expect(ws.getCell("A2").value).to.deep.equal({ sharedFormula: "A1", result: 2 });
        });

        it("writes the range of the formula without `$` and with its corners in order", () => {
            const ws = sheet();
            ws.fillFormula("$B$2:$A$1", "C1", undefined, "array");
            expect(ws.getCell("A1").value.ref).to.equal("A1:B2");
        });

        it("fills a single cell", () => {
            const ws = sheet();
            ws.fillFormula("C3", "SUM(A1:B2)", [10], "array");
            expect(ws.getCell("C3").value).to.deep.equal({ shareType: "array", formula: "SUM(A1:B2)", ref: "C3", result: 10 });
        });

        it("rejects a range of another sheet, or an invalid range, filling nothing", () => {
            for (const range of ["y!A1:A2", "foo", "A0:A2", "a1:a2", ""]) {
                const ws = sheet();
                expect(() => ws.fillFormula(range, "B1"), range).to.throw(Error, /^Cannot fill a formula/);
                expect(rowKeys(ws)).to.deep.equal([]);
            }
        });
    });

    describe("cells and rows", () => {
        it("rejects a cell or a row out of the worksheet", () => {
            const ws = sheet();
            const cells = [["A0"], [0, 1], [-1, 1], [1, 0], [1.5, 1], [1048577, 1], [1, 16385], ["foo"], ["a1"], ["A1:B2"], [""]];
            for (const args of cells) {
                expect(() => ws.getCell(...args), JSON.stringify(args)).to.throw(Error, /^Invalid cell/);
            }
            expect(() => ws.getCell("y!A1")).to.throw(Error, 'The cell y!A1 is not on the worksheet "s"');
            for (const r of [0, -1, 1.5, 1048577, Number.NaN]) {
                expect(() => ws.getRow(r), String(r)).to.throw(Error, /^Invalid row number/);
            }
            expect(() => ws.getRows(0, 2)).to.throw(Error, /^Invalid row number/);
            expect(rowKeys(ws)).to.deep.equal([]);
        });

        it("finds no cell or row out of the worksheet, without throwing", () => {
            const ws = sheet();
            ws.getCell("A1").value = 1;
            for (const args of [["A0"], [0, 1], [1, 0], ["foo"], ["y!A1"], [1048577, 1]]) {
                expect(ws.findCell(...args), JSON.stringify(args)).to.be.undefined();
            }
            expect(ws.findRow(0)).to.be.undefined();
            expect(ws.findCell("s!A1").value).to.equal(1);
            expect(ws.findCell("$A$1").value).to.equal(1);
            expect(rowKeys(ws)).to.deep.equal(["0"]);
        });

        it("decodes hostile references in linear time", () => {
            const ws = sheet();
            const inputs = [
                `${"$A".repeat(20000)}1`,
                `${"A".repeat(40000)}1`,
                `${"'".repeat(40000)}!A1`,
                `'${"''".repeat(20000)}x'!A1`,
                `${"!".repeat(40000)}A1`,
                `A1${":".repeat(40000)}`,
                `${"1".repeat(40000)}`,
            ];
            for (const input of inputs) {
                const start = process.hrtime.bigint();
                expect(ws.findCell(input)).to.be.undefined();
                expect(() => ws.mergeCells(input)).to.throw(Error);
                expect(() => ws.fillFormula(input, "A1")).to.throw(Error);
                const ms = Number(process.hrtime.bigint() - start) / 1e6;
                expect(ms, input.slice(0, 20)).to.be.below(50);
            }
        });

        it("gets a cell by its address, with the worksheet's own sheet name or not", () => {
            const ws = sheet();
            expect(ws.getCell("S!B2").address).to.equal("B2");
            expect(ws.getCell("$C$3").address).to.equal("C3");
            expect(ws.getCell("XFD1048576").address).to.equal("XFD1048576");
            expect(ws.getCell(2, 3).address).to.equal("C2");
        });
    });
});
