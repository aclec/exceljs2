const DefinedNames = verquire("doc/defined-names");
const colCache = verquire("utils/col-cache");

describe("DefinedNames", () => {
    it("adds names for cells", () => {
        const dn = new DefinedNames();

        dn.add("blort!A1", "foo");
        expect(dn.getNames("blort!A1")).to.deep.equal(["foo"]);
        expect(dn.getNames("blort!$A$1")).to.deep.equal(["foo"]);

        dn.add("blort!$B$4", "bar");
        expect(dn.getNames("blort!B4")).to.deep.equal(["bar"]);
        expect(dn.getNames("blort!$B$4")).to.deep.equal(["bar"]);

        dn.add("'blo rt'!$B$4", "bar");
        expect(dn.getNames("'blo rt'!$B$4")).to.deep.equal(["bar"]);
        dn.add("'blo ,!rt'!$B$4", "bar");
        expect(dn.getNames("'blo ,!rt'!$B$4")).to.deep.equal(["bar"]);
    });

    it("removes names for cells", () => {
        const dn = new DefinedNames();

        dn.add("blort!A1", "foo");
        dn.add("blort!A1", "bar");
        dn.remove("blort!A1", "foo");

        expect(dn.getNames("blort!A1")).to.deep.equal(["bar"]);
    });

    describe("removing a range", () => {
        it("removes the cells of a range reference", () => {
            const dn = new DefinedNames();
            dn.add("s!$A$1:$B$2", "r");
            dn.add("s!$A$1:$B$2", "other");
            dn.remove("s!$A$1:$B$2", "r");

            expect(dn.getRanges("r").ranges).to.deep.equal([]);
            expect(dn.model).to.deep.equal([{ name: "other", ranges: ["s!$A$1:$B$2"] }]);
        });

        it("removes only the cells inside the range, on its sheet", () => {
            const dn = new DefinedNames();
            dn.add("s!$A$1:$C$3", "r");
            dn.add("t!$A$1:$C$3", "r");
            dn.remove("s!$B$1:$C$2", "r");

            expect(dn.getRanges("r").ranges).to.deep.equal(["s!$A$1:$A$3", "s!$B$3:$C$3", "t!$A$1:$C$3"]);
            expect(dn.getNames("s!$B$2")).to.deep.equal([]);
            expect(dn.getNames("s!$C$3")).to.deep.equal(["r"]);
        });

        it("removes a range given by its bounds or decoded (removeEx)", () => {
            const dn = new DefinedNames();
            dn.add("s!$A$1:$B$2", "bounds");
            dn.add("s!$A$1:$B$2", "decoded");
            dn.removeEx({ sheetName: "s", top: 1, left: 1, bottom: 2, right: 2 }, "bounds");
            dn.removeEx(colCache.decodeEx("s!$A$1:$B$2"), "decoded");

            expect(dn.model).to.deep.equal([]);
        });

        it("removes a range from every name (removeAllNames)", () => {
            const dn = new DefinedNames();
            dn.add("s!$A$1:$B$2", "a");
            dn.add("s!$B$2:$C$3", "b");
            dn.removeAllNames({ sheetName: "s", top: 2, left: 2, bottom: 3, right: 3 });

            expect(dn.model).to.deep.equal([{ name: "a", ranges: ["s!$A$1:$A$2", "s!$B$1"] }]);
        });

        it("throws for a range whose bounds are not rows and columns of the grid", () => {
            const dn = new DefinedNames();
            dn.add("s!$A$1:$B$2", "a");
            for (const bounds of [
                { top: 0, left: 1, bottom: 2, right: 2 },
                { top: 2, left: 1, bottom: 1, right: 2 },
                { top: 1, left: 1.5, bottom: 2, right: 2 },
                { top: 1, left: 1, bottom: 2, right: 16385 },
                { top: 1, left: 1, bottom: 1048577, right: 2 },
            ]) {
                expect(() => dn.removeEx({ sheetName: "s", ...bounds }, "a")).to.throw(Error, /Invalid range/);
                expect(() => dn.removeAllNames({ sheetName: "s", ...bounds })).to.throw(Error, /Invalid range/);
            }
            expect(dn.model).to.deep.equal([{ name: "a", ranges: ["s!$A$1:$B$2"] }]);
        });

        it("removes a range from a name without cells there, or an unknown name", () => {
            const dn = new DefinedNames();
            dn.add("s!$A$1", "a");
            dn.remove("s!$C$3:$D$4", "a");
            dn.remove("s!$C$3:$D$4", "unknown");

            expect(dn.model).to.deep.equal([{ name: "a", ranges: ["s!$A$1"] }]);
        });

        it("cuts a range kept as a rectangle into the rectangles that remain, without expanding it", () => {
            const dn = new DefinedNames();
            dn.add("s!$A$1:$XFD$1048576", "everything");
            dn.remove("s!$B$2:$C$3", "everything");

            expect(dn.getRanges("everything").ranges).to.deep.equal([
                "s!$A$1:$XFD$1",
                "s!$A$2:$A$3",
                "s!$D$2:$XFD$3",
                "s!$A$4:$XFD$1048576",
            ]);
            expect(dn.getNames("s!$B$2")).to.deep.equal([]);
            expect(dn.getNames("s!$C$3")).to.deep.equal([]);
            expect(dn.getNames("s!$D$3")).to.deep.equal(["everything"]);
            expect(dn.matrixMap.everything).to.equal(undefined);
        });

        it("removes a rectangle covered by the range, and keeps one of another sheet", () => {
            const dn = new DefinedNames();
            dn.add("s!$A$1:$XFD$1048576", "big");
            dn.add("t!$A$1:$XFD$1048576", "big");
            dn.remove("s!$A$1:$XFD$1048576", "big");

            expect(dn.getRanges("big").ranges).to.deep.equal(["t!$A$1:$XFD$1048576"]);
        });

        it("cuts a cell out of a range kept as a rectangle (the rest of the range stays named)", () => {
            const dn = new DefinedNames();
            dn.add("s!$A$1:$B$100000", "big");
            dn.remove("s!$A$50", "big");

            expect(dn.getRanges("big").ranges).to.deep.equal(["s!$A$1:$B$49", "s!$B$50", "s!$A$51:$B$100000"]);
            expect(dn.getNames("s!$A$50")).to.deep.equal([]);
            expect(dn.getNames("s!$A$51")).to.deep.equal(["big"]);
        });
    });

    // get ranges example
    it("gets the right ranges for a name", () => {
        const dn = new DefinedNames();

        dn.add("blort!A1", "vertical");
        dn.add("blort!A2", "vertical");
        dn.add("blort!A3", "vertical");

        dn.add("blort!C1", "horizontal");
        dn.add("blort!D1", "horizontal");
        dn.add("blort!E1", "horizontal");

        dn.add("blort!C3", "square");
        dn.add("blort!D3", "square");
        dn.add("blort!C4", "square");
        dn.add("blort!D4", "square");

        dn.add("other!A1", "single");

        expect(dn.getRanges("vertical")).to.deep.equal({
            name: "vertical",
            ranges: ["blort!$A$1:$A$3"],
        });
        expect(dn.getRanges("horizontal")).to.deep.equal({
            name: "horizontal",
            ranges: ["blort!$C$1:$E$1"],
        });
        expect(dn.getRanges("square")).to.deep.equal({
            name: "square",
            ranges: ["blort!$C$3:$D$4"],
        });
        expect(dn.getRanges("single")).to.deep.equal({
            name: "single",
            ranges: ["other!$A$1"],
        });
    });

    it("splices", () => {
        const dn = new DefinedNames();
        dn.add("vertical!A1", "vertical");
        dn.add("vertical!A2", "vertical");
        dn.add("vertical!A3", "vertical");
        dn.add("vertical!A4", "vertical");

        dn.add("horizontal!A1", "horizontal");
        dn.add("horizontal!B1", "horizontal");
        dn.add("horizontal!C1", "horizontal");
        dn.add("horizontal!D1", "horizontal");

        ["A", "B", "C", "D"].forEach((col) => {
            [1, 2, 3, 4].forEach((row) => {
                dn.add(`square!${col}${row}`, "square");
            });
        });

        dn.add("single!A1", "singleA1");
        dn.add("single!D1", "singleD1");
        dn.add("single!A4", "singleA4");
        dn.add("single!D4", "singleD4");

        dn.spliceRows("vertical", 2, 2, 1);
        dn.spliceColumns("horizontal", 2, 2, 1);
        dn.spliceRows("square", 2, 2, 1);
        dn.spliceColumns("square", 2, 2, 1);
        dn.spliceRows("single", 2, 2, 1);
        dn.spliceColumns("single", 2, 2, 1);

        expect(dn.getRanges("vertical")).to.deep.equal({
            name: "vertical",
            ranges: ["vertical!$A$1", "vertical!$A$3"],
        });
        expect(dn.getRanges("horizontal")).to.deep.equal({
            name: "horizontal",
            ranges: ["horizontal!$A$1", "horizontal!$C$1"],
        });
        expect(dn.getRanges("square")).to.deep.equal({
            name: "square",
            ranges: ["square!$A$1", "square!$C$1", "square!$A$3", "square!$C$3"],
        });
        expect(dn.getRanges("singleA1")).to.deep.equal({
            name: "singleA1",
            ranges: ["single!$A$1"],
        });
        expect(dn.getRanges("singleD1")).to.deep.equal({
            name: "singleD1",
            ranges: ["single!$C$1"],
        });
        expect(dn.getRanges("singleA4")).to.deep.equal({
            name: "singleA4",
            ranges: ["single!$A$3"],
        });
        expect(dn.getRanges("singleD4")).to.deep.equal({
            name: "singleD4",
            ranges: ["single!$C$3"],
        });
    });

    it("creates matrix from model", () => {
        const dn = new DefinedNames();

        dn.model = [];
        dn.add("blort!A1", "bar");
        dn.remove("blort!A1", "foo");

        expect(dn.getNames("blort!A1")).to.deep.equal(["bar"]);
    });

    it("skips values with invalid range", () => {
        const dn = new DefinedNames();
        dn.model = [
            { name: "eq", ranges: ['"="'] },
            { name: "ref", ranges: ["#REF!"] },
            { name: "single", ranges: ["Sheet3!$A$1"] },
            { name: "range", ranges: ["Sheet3!$A$2:$F$2228"] },
        ];

        expect(dn.model).to.deep.equal([
            { name: "single", ranges: ["Sheet3!$A$1"] },
            { name: "range", ranges: ["Sheet3!$A$2:$F$2228"] },
        ]);
    });
});
