const StylesXform = verquire("xlsx/xform/style/styles-xform");
const Enums = verquire("doc/enums");

// StylesXform as style manager: the ids it gives to the styles of the cells (the per-cell write path)
describe("StylesXform style manager", () => {
    it("gives a style with only a number format the id of its <xf>, as a style with other parts", () => {
        const styles = new StylesXform(true);
        const percent = styles.addStyleModel({ numFmt: "0.00%" }, Enums.ValueType.Number);
        const bold = styles.addStyleModel({ font: { bold: true } }, Enums.ValueType.Number);
        const general = styles.addStyleModel({}, Enums.ValueType.Number);
        const date = styles.addStyleModel({}, Enums.ValueType.Date);
        const text = styles.addStyleModel({}, Enums.ValueType.String);

        // the same ids as the cell formats rendered for these styles
        expect(styles.addStyleModel({ numFmt: "0.00%" }, Enums.ValueType.Number)).to.equal(percent);
        expect(styles._addStyle({ numFmtId: styles._addNumFmtStr("0.00%") })).to.equal(percent);
        expect(styles._addStyle({ numFmtId: 0 })).to.equal(general);
        expect(styles._addStyle({ numFmtId: 14 })).to.equal(date);
        expect(styles._addStyle({})).to.equal(text);
        expect(new Set([percent, bold, general, date, text]).size).to.equal(4); // General (0) is the default style 0
        expect(general).to.equal(0);
        expect(text).to.equal(0);
    });

    it("renders the cell formats of the styles registered", () => {
        const styles = new StylesXform(true);
        const ids = [
            styles.addStyleModel({ numFmt: "0.0" }, Enums.ValueType.Number),
            styles.addStyleModel({}, Enums.ValueType.Date),
            styles.addStyleModel({ numFmt: "0.0", alignment: { horizontal: "center" } }, Enums.ValueType.Number),
            styles.addStyleModel({ numFmt: "0.0" }, Enums.ValueType.Number),
        ];
        expect(ids).to.deep.equal([1, 2, 3, 1]);
        const { xml } = styles;
        expect(xml).to.contain(
            '<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
                '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
                '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
                '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1">' +
                '<alignment horizontal="center"/></xf></cellXfs>'
        );
    });

    it("returns no differential format for a parsed part without dxfs", async () => {
        const styles = new StylesXform();
        await styles.parseStream([
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="0"/></styleSheet>',
        ]);
        expect(styles.getDxfStyle(0)).to.equal(undefined);
        expect(styles.getStyleModel(1)).to.equal(null);
    });
});
