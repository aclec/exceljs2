const testXformHelper = require("../test-xform-helper");

const HyperlinkXform = verquire("xlsx/xform/sheet/hyperlink-xform");

const expectations = [
    {
        title: "Web Link",
        create() {
            return new HyperlinkXform();
        },
        preparedModel: { address: "B6", rId: "rId1" },
        get parsedModel() {
            return this.preparedModel;
        },
        xml: '<hyperlink ref="B6" r:id="rId1"/>',
        tests: ["render", "renderIn", "parse"],
    },
    {
        title: "Internal Link",
        create() {
            return new HyperlinkXform();
        },
        preparedModel: { address: "B6", target: "#sheet1!B2" },
        parsedModel: { address: "B6", target: "sheet1!B2" },
        xml: '<hyperlink ref="B6" location="sheet1!B2"/>',
        tests: ["render", "renderIn", "parse"],
    },
    {
        title: "Defined Name Link",
        create() {
            return new HyperlinkXform();
        },
        preparedModel: { address: "B6", target: "#MyName", tooltip: "go" },
        parsedModel: { address: "B6", target: "MyName", tooltip: "go" },
        xml: '<hyperlink ref="B6" tooltip="go" location="MyName"/>',
        tests: ["render", "renderIn", "parse"],
    },
    {
        title: "External Link with Location",
        create() {
            return new HyperlinkXform();
        },
        preparedModel: { address: "B6", rId: "rId1", target: "https://example.com/doc#part2" },
        parsedModel: { address: "B6", rId: "rId1", target: "part2" },
        xml: '<hyperlink ref="B6" r:id="rId1" location="part2"/>',
        tests: ["parse"],
    },
];

describe("HyperlinkXform", () => {
    testXformHelper(expectations);

    describe("getLocation", () => {
        it("reads '#'-prefixed targets as a place in the workbook", () => {
            expect(HyperlinkXform.getLocation("#Name")).to.equal("Name");
            expect(HyperlinkXform.getLocation("#Sheet1!A1")).to.equal("Sheet1!A1");
            expect(HyperlinkXform.getLocation("#'My Sheet'!B2:C3")).to.equal("'My Sheet'!B2:C3");
        });

        it("reads legacy unprefixed sheet references as a place in the workbook", () => {
            ["Sheet1!A1", "'My Sheet'!B2", "'It''s'!A1", "My Sheet!$B$2", "Data.2024!A1:B20", "'a/b#c'!XFD1048576"].forEach((target) =>
                expect(HyperlinkXform.getLocation(target), target).to.equal(target)
            );
        });

        it("reads URLs and paths whose tail looks like a sheet reference as external", () => {
            [
                "Name",
                "https://x/#!a1",
                "https://example.com/#!page1",
                "http://example.com?x!A1",
                "mailto:a@b",
                "mailto:a!b1",
                "file:///c:/x.xlsx#Sheet!A1",
                "../other.xlsx#Sheet1!A1",
                "..\\other.xlsx!A1",
                "dir/Sheet1!A1",
                "Sheet1!A1x",
                "Sheet1!ABCD1",
                "'unterminated!A1",
                undefined,
            ].forEach((target) => expect(HyperlinkXform.getLocation(target), target).to.equal(undefined));
        });
    });
});
