// Invalid models and arguments of the document classes (note, image, range, column, pivot table): a clear Error or a
// clean result, not a TypeError deep inside or a broken model.
const Excel = verquire("exceljs");
const Note = verquire("doc/note");
const Range = verquire("doc/range");

describe("Document invalid models", () => {
    it("a note model without the texts of its comment throws", () => {
        expect(() => Note.fromModel({ type: "note", note: {} })).to.throw(Error, "The model of a note needs the texts of its comment");
        expect(Note.fromModel({ type: "note", note: { texts: [{ text: "a" }] } }).note).to.equal("a");
    });

    it("an image whose range was removed throws when its model is read", () => {
        const wb = new Excel.Workbook();
        const ws = wb.addWorksheet("s");
        const imageId = wb.addImage({ buffer: Buffer.from([]), extension: "png" });
        ws.addImage(imageId, "B2:D4");
        const [image] = ws.getImages();
        delete image.range;
        expect(() => image.model).to.throw(Error, `The image ${imageId} has no range`);
    });

    it("a range location with a bound that is not 1-based throws", () => {
        expect(() => new Range({ top: 0, left: 1, bottom: 2, right: 2 })).to.throw(
            Error,
            "Invalid range: top 0, left 1, bottom 2, right 2"
        );
        expect(new Range({ top: 1, left: 1, bottom: 2, right: 2 }).range).to.equal("A1:B2");
    });

    it("expandToAddress ignores a reference that is not a cell; contains is false for it", () => {
        const range = new Range("B2:C3");
        range.expandToAddress("A1:D4");
        range.expandToAddress("#REF!");
        expect(range.model).to.deep.include({ top: 2, left: 2, bottom: 3, right: 3 });
        range.expandToAddress("E5");
        expect(range.range).to.equal("B2:E5");
        expect(range.contains("A1:B2")).to.equal(false);
        expect(range.contains("C3")).to.equal(true);
    });

    it("column eachCell throws without an iteratee function", () => {
        const ws = new Excel.Workbook().addWorksheet("s");
        ws.getCell("A1").value = 1;
        expect(() => ws.getColumn(1).eachCell()).to.throw(Error, "eachCell needs an iteratee function");
        expect(() => ws.getColumn(1).eachCell({ includeEmpty: true })).to.throw(Error, "eachCell needs an iteratee function");
    });

    it("a pivot table field is named by a string, a number or a boolean header, not by the text of another value", () => {
        const wb = new Excel.Workbook();
        const source = wb.addWorksheet("source");
        const date = new Date(Date.UTC(2024, 0, 1));
        source.addRows([
            ["A", 5, date],
            ["a", 1, 2],
        ]);
        expect(() =>
            wb.addWorksheet("p1").addPivotTable({ sourceSheet: source, rows: ["A"], columns: [String(date)], values: ["5"] })
        ).to.throw(Error, `The header name "${String(date)}" was not found in source.`);
        const pivot = wb.addWorksheet("p2");
        pivot.addPivotTable({ sourceSheet: source, rows: ["A"], columns: ["5"], values: ["5"] });
        expect(pivot.pivotTables[0].columns).to.deep.equal([1]);
    });
});
