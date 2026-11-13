const testXformHelper = require("../test-xform-helper");

const WorkbookPropertiesXform = verquire("xlsx/xform/book/workbook-properties-xform");

const expectations = [
    {
        title: "default",
        create() {
            return new WorkbookPropertiesXform();
        },
        preparedModel: {},
        xml: '<workbookPr defaultThemeVersion="164011" filterPrivacy="1"></workbookPr>',
        parsedModel: {},
        tests: ["render", "renderIn"],
    },
    {
        title: "date1904",
        create() {
            return new WorkbookPropertiesXform();
        },
        preparedModel: { date1904: true },
        xml: '<workbookPr date1904="1" defaultThemeVersion="164011" filterPrivacy="1"></workbookPr>',
        parsedModel: { date1904: true },
        tests: ["render", "renderIn", "parse"],
    },
    // date1904 is an xsd:boolean: "true" / "1" or "false" / "0"
    {
        title: 'date1904="true"',
        create() {
            return new WorkbookPropertiesXform();
        },
        xml: '<workbookPr date1904="true"></workbookPr>',
        parsedModel: { date1904: true },
        tests: ["parse"],
    },
    {
        title: 'date1904="false"',
        create() {
            return new WorkbookPropertiesXform();
        },
        xml: '<workbookPr date1904="false"></workbookPr>',
        parsedModel: { date1904: false },
        tests: ["parse"],
    },
    {
        title: 'date1904="0"',
        create() {
            return new WorkbookPropertiesXform();
        },
        xml: '<workbookPr date1904="0"></workbookPr>',
        parsedModel: { date1904: false },
        tests: ["parse"],
    },
    {
        title: "date1904 absent",
        create() {
            return new WorkbookPropertiesXform();
        },
        xml: '<workbookPr defaultThemeVersion="164011"></workbookPr>',
        parsedModel: { date1904: false },
        tests: ["parse"],
    },
];

describe("WorkbookPropertiesXform", () => {
    testXformHelper(expectations);
});
