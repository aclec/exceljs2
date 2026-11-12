const testXformHelper = require("../test-xform-helper");

const NumFmtXform = verquire("xlsx/xform/style/numfmt-xform");

const expectations = [
    {
        title: "date",
        create: () => new NumFmtXform(),
        preparedModel: { id: 165, formatCode: "d-mmm-yyyy" },
        xml: '<numFmt numFmtId="165" formatCode="d-mmm-yyyy"/>',
        get parsedModel() {
            return this.preparedModel;
        },
        tests: ["render", "renderIn", "parse"],
    },
    {
        title: "thing",
        create: () => new NumFmtXform(),
        preparedModel: { id: 165, formatCode: "[Green]#,##0 ;[Red](#,##0)" },
        xml: '<numFmt numFmtId="165" formatCode="[Green]#,##0 ;[Red](#,##0)"/>',
        get parsedModel() {
            return this.preparedModel;
        },
        tests: ["render", "renderIn", "parse"],
    },
    {
        title: "escaped characters",
        create: () => new NumFmtXform(),
        preparedModel: { id: 166, formatCode: '#,##0.00\\ "€";[Red]\\-#,##0.00\\ "€"' },
        xml: '<numFmt numFmtId="166" formatCode="#,##0.00\\ &quot;€&quot;;[Red]\\-#,##0.00\\ &quot;€&quot;"/>',
        get parsedModel() {
            return this.preparedModel;
        },
        tests: ["render", "renderIn", "parse"],
    },
];

describe("NumFmtXform", () => {
    testXformHelper(expectations);
});
