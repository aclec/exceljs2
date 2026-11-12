const StylesXform = verquire("xlsx/xform/style/styles-xform");
const Enums = verquire("doc/enums");

// a styles part whose cell format i uses font i, fill i, border i and the alignment of its <xf>
async function styles({ fonts = ["<font/>"], fills = ["<fill/>"], borders = ["<border/>"], xfs }) {
    const xform = new StylesXform();
    await xform.parseStream([
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
            `<fonts>${fonts.join("")}</fonts><fills>${fills.join("")}</fills><borders>${borders.join("")}</borders>` +
            `<cellXfs>${xfs.join("")}</cellXfs></styleSheet>`,
    ]);
    return xform;
}

const xf = (id, alignment = "") => `<xf numFmtId="0" fontId="${id}" fillId="${id}" borderId="${id}">${alignment}</xf>`;

describe("StylesXform values read from a file", () => {
    it("reads a cell fill without (valid) patternType as the pattern none, and a <fill> without child as no fill", async () => {
        const xform = await styles({
            fills: [
                '<fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill>',
                "<fill><patternFill/></fill>",
                '<fill><patternFill patternType="Solid"/></fill>',
                "<fill/>",
            ],
            fonts: ["<font/>", "<font/>", "<font/>", "<font/>"],
            borders: ["<border/>", "<border/>", "<border/>", "<border/>"],
            xfs: [xf(0), xf(1), xf(2), xf(3)],
        });
        expect(xform.getStyleModel(0).fill).to.deep.equal({ type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } });
        expect(xform.getStyleModel(1).fill).to.deep.equal({ type: "pattern", pattern: "none" });
        expect(xform.getStyleModel(2).fill).to.deep.equal({ type: "pattern", pattern: "none" });
        expect(xform.getStyleModel(3)).not.to.have.property("fill");
    });

    it("keeps the valid vertAlign / scheme / underline of a font and drops the invalid ones", async () => {
        const xform = await styles({
            fonts: [
                '<font><u val="double"/><vertAlign val="baseline"/><scheme val="major"/></font>',
                '<font><u val="bogus"/><vertAlign val="up"/><scheme val="bogus"/></font>',
            ],
            fills: ["<fill/>", "<fill/>"],
            borders: ["<border/>", "<border/>"],
            xfs: [xf(0), xf(1)],
        });
        expect(xform.getStyleModel(0).font).to.deep.equal({ underline: "double", vertAlign: "baseline", scheme: "major" });
        expect(xform.getStyleModel(1).font).to.deep.equal({ underline: true });
    });

    it("reads a general or invalid horizontal / vertical alignment as absent", async () => {
        const xform = await styles({
            fonts: ["<font/>", "<font/>"],
            fills: ["<fill/>", "<fill/>"],
            borders: ["<border/>", "<border/>"],
            xfs: [
                xf(0, '<alignment horizontal="general" vertical="center"/>'),
                xf(1, '<alignment horizontal="middle" vertical="up" wrapText="1"/>'),
            ],
        });
        expect(xform.getStyleModel(0).alignment).to.deep.equal({ vertical: "middle" });
        expect(xform.getStyleModel(1).alignment).to.deep.equal({ wrapText: true });
    });

    it("reads a border edge of style none or of an invalid style as no style", async () => {
        const xform = await styles({
            borders: ['<border><left style="thin"/><right style="none"><color rgb="FF000000"/></right><top style="bogus"/></border>'],
            xfs: [xf(0)],
        });
        expect(xform.getStyleModel(0).border).to.deep.equal({ left: { style: "thin" }, right: { color: { argb: "FF000000" } } });
    });

    it("throws a clear Error when a style is registered without the style manager role", async () => {
        const xform = await styles({ xfs: [xf(0)] });
        expect(() => xform.addStyleModel({ font: { bold: true } }, Enums.ValueType.Number)).to.throw(
            Error,
            "StylesXform is not a style manager: call init() first"
        );
    });
});
