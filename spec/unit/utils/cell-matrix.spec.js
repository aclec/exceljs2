const CellMatrix = verquire("utils/cell-matrix");

describe("CellMatrix", () => {
    it("getCell always returns a cell", () => {
        const cm = new CellMatrix();
        expect(cm.getCell("A1")).to.be.ok();
        expect(cm.getCell("$B$2")).to.be.ok();
        expect(cm.getCell("Sheet!$C$3")).to.be.ok();
    });
    it("findCell only returns known cells", () => {
        const cm = new CellMatrix();
        expect(cm.findCell("A1")).to.be.undefined();
        expect(cm.findCell("$B$2")).to.be.undefined();
        expect(cm.findCell("Sheet!$C$3")).to.be.undefined();
    });
    it("merges a copy of the template into the cells created from a location", () => {
        const template = { mark: true, nested: { a: 1 } };
        const cm = new CellMatrix(template);
        const cell = cm.getCell("Sheet!$C$3");
        expect(cell).to.deep.include({ sheetName: "Sheet", address: "C3", row: 3, col: 3, mark: true, nested: { a: 1 } });
        expect(cell.nested).not.to.equal(template.nested);
        expect(cm.map((c) => c.address)).to.deep.equal(["C3"]);
    });
    it("keeps one location without row and column per sheet, never enumerated", () => {
        const cm = new CellMatrix();
        const error = cm.getCell("Sheet!#REF!");
        expect(error).to.deep.equal({ sheetName: "Sheet", error: "#REF!" });
        // any other such location of the sheet finds it
        const range = { sheetName: "Sheet", top: 0, left: 1, bottom: 2, right: 2 };
        cm.addCellEx(range);
        expect(cm.findCellEx(range)).to.equal(error);
        expect(cm.findCell("Other!#REF!")).to.be.undefined();
        expect(cm.map((c) => c)).to.deep.equal([]);
        // splicing columns keeps it, splicing rows drops it
        cm.spliceColumns("Sheet", 1, 1, 0);
        expect(cm.findCell("Sheet!#REF!")).to.equal(error);
        cm.spliceRows("Sheet", 1, 1, 0);
        expect(cm.findCell("Sheet!#REF!")).to.be.undefined();
        cm.addCell("Sheet!#REF!");
        cm.removeCellEx({ sheetName: "Sheet", error: "#N/A" });
        expect(cm.findCell("Sheet!#REF!")).to.be.undefined();
    });
});
