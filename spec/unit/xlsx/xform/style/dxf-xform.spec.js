const DxfXform = verquire("xlsx/xform/style/dxf-xform");

describe("DxfXform", () => {
    it("renders its children in CT_Dxf order", () => {
        const xml = new DxfXform().toXml({
            protection: { locked: false },
            border: { top: { style: "thin" } },
            alignment: { horizontal: "center" },
            fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFF0000" } },
            numFmt: "0.0%",
            numFmtId: 164,
            font: { bold: true },
        });
        // font, numFmt, fill, alignment, protection, border (ECMA-376 Part 1, 18.8.17)
        expect(xml).to.equal(
            '<dxf><font><b/></font><numFmt numFmtId="164" formatCode="0.0%"/>' +
                '<fill><patternFill patternType="solid"><bgColor rgb="FFFF0000"/></patternFill></fill>' +
                '<alignment horizontal="center"/><protection locked="0"/>' +
                '<border><left/><right/><top style="thin"/><bottom/><diagonal/></border></dxf>'
        );
    });
});
