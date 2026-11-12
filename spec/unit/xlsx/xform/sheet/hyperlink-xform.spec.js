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
            [
                "Sheet1!A1",
                "'My Sheet'!B2",
                "'It''s'!A1",
                "My Sheet!$B$2",
                "Data.2024!A1:B20",
                "'a/b#c'!XFD1048576",
                // sheet names may hold ' (not first or last) and #
                "Bob's!A1",
                "Q#1!B2",
                "Bob's #2!A1:B2",
            ].forEach((target) => expect(HyperlinkXform.getLocation(target), target).to.equal(target));
        });

        it("reads URLs and paths whose tail looks like a sheet reference as external", () => {
            [
                "Name",
                "https://x/#!a1",
                "https://example.com/#!page1",
                "http://example.com?x!A1",
                // relative links to another workbook: a file name with an extension before the #
                "Book2.xlsx#Sheet1!A1",
                "report v1.2.xlsm#'My Sheet'!B2",
                "mailto:a@b",
                "mailto:a!b1",
                "file:///c:/x.xlsx#Sheet!A1",
                "../other.xlsx#Sheet1!A1",
                "..\\other.xlsx!A1",
                "dir/Sheet1!A1",
                "Sheet1!A1x",
                "Sheet1!ABCD1",
                "'unterminated!A1",
                "Bob'!A1",
                "'Bob!A1",
                undefined,
            ].forEach((target) => expect(HyperlinkXform.getLocation(target), target).to.equal(undefined));
        });

        it("tells legacy references from URLs in linear time", () => {
            ["'", "a", "a'", "'a", "''", "a#", "#'"].forEach((unit) => {
                const target = `${unit.repeat(10000)}!A1x`;
                const start = Date.now();
                HyperlinkXform.getLocation(target);
                expect(Date.now() - start, JSON.stringify(unit)).to.be.below(50);
            });
        });
    });
});
