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
});
