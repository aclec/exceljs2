const XmlStream = verquire("utils/xml-stream");
const PhoneticTextXform = verquire("xlsx/xform/strings/phonetic-text-xform");
const PivotTableXform = verquire("xlsx/xform/pivot-table/pivot-table-xform");
const TwoCellAnchorXform = verquire("xlsx/xform/drawing/two-cell-anchor-xform");
const TableColumnXform = verquire("xlsx/xform/table/table-column-xform");

// a model missing a part that the xform needs gives a clear Error, not a TypeError from deep inside
describe("Xform render errors", () => {
    it("PhoneticTextXform: a run without text nor rich text", () => {
        expect(() => new PhoneticTextXform().render(new XmlStream(), { sb: 0, eb: 1 })).to.throw(
            Error,
            "A phonetic run (<rPh>) needs its text or its rich text"
        );
    });

    it("PivotTableXform: a pivot table without value field", () => {
        const model = { rows: [0], columns: [1], values: [], cacheFields: [{ name: "a" }, { name: "b" }], cacheId: "10" };
        expect(() => new PivotTableXform().render(new XmlStream(), model)).to.throw(
            Error,
            "A pivot table needs a value field, one of its cache fields"
        );
    });

    it("TwoCellAnchorXform: an anchor without bottom-right corner", () => {
        const tl = { nativeCol: 0, nativeColOff: 0, nativeRow: 0, nativeRowOff: 0 };
        expect(() => new TwoCellAnchorXform().render(new XmlStream(), { range: { tl }, picture: { rId: "rId1" } })).to.throw(
            Error,
            "An image anchored by two cells needs its bottom-right corner (range.br) and its picture"
        );
    });

    it("TableColumnXform: a column not prepared", () => {
        expect(() => new TableColumnXform().prepare({ name: "A" }, {})).to.throw(
            Error,
            "TableColumnXform.prepare needs the index of the column"
        );
        expect(() => new TableColumnXform().render(new XmlStream(), { name: "A" })).to.throw(Error, 'Table column "A" has no id');
    });
});
