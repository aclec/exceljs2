const testXformHelper = require("../../test-xform-helper");

const CfvoXform = verquire("xlsx/xform/sheet/cf/cfvo-xform");

const expectations = [
    {
        title: "min",
        create() {
            return new CfvoXform();
        },
        preparedModel: { type: "min" },
        xml: '<cfvo type="min" />',
        get parsedModel() {
            return this.preparedModel;
        },
        tests: ["render", "parse"],
    },
    {
        title: "percent",
        create() {
            return new CfvoXform();
        },
        preparedModel: { type: "percent", value: 12.5 },
        xml: '<cfvo type="percent" val="12.5" />',
        get parsedModel() {
            return this.preparedModel;
        },
        tests: ["render", "parse"],
    },
    {
        title: "formula",
        create() {
            return new CfvoXform();
        },
        preparedModel: { type: "formula", value: "AVERAGE($A$1:$A$9)" },
        xml: '<cfvo type="formula" val="AVERAGE($A$1:$A$9)" />',
        get parsedModel() {
            return this.preparedModel;
        },
        tests: ["render", "parse"],
    },
];

describe("CfvoXform", () => {
    testXformHelper(expectations);

    it("reads number literals as numbers and anything else as a formula", () => {
        expect(["50", "-1.5", "+.5", "1E3", "2.5e-1"].map(CfvoXform.toValue)).to.deep.equal([50, -1.5, 0.5, 1000, 0.25]);
        expect(["$A$1", "1+1", "0x10", "Infinity", "1,5", ""].map(CfvoXform.toValue)).to.deep.equal([
            "$A$1",
            "1+1",
            "0x10",
            "Infinity",
            "1,5",
            "",
        ]);
        expect(CfvoXform.toValue(undefined)).to.equal(undefined);
    });
});
