const testXformHelper = require("../test-xform-helper");

const FontXform = verquire("xlsx/xform/style/font-xform");

const expectations = [
    {
        title: "green bold",
        create() {
            return new FontXform();
        },
        preparedModel: {
            bold: true,
            size: 14,
            color: { argb: "FF00FF00" },
            name: "Calibri",
            family: 2,
            scheme: "minor",
        },
        xml: '<font><b/><color rgb="FF00FF00"/><family val="2"/><scheme val="minor"/><sz val="14"/><name val="Calibri"/></font>',
        get parsedModel() {
            return this.preparedModel;
        },
        tests: ["render", "renderIn", "parse"],
    },
    {
        title: "rPr tag",
        create() {
            return new FontXform({ tagName: "rPr", fontNameTag: "rFont" });
        },
        preparedModel: {
            italic: true,
            size: 14,
            color: { argb: "FF00FF00" },
            name: "Calibri",
            family: 2,
            scheme: "minor",
        },
        xml: '<rPr><i/><color rgb="FF00FF00"/><family val="2"/><scheme val="minor"/><sz val="14"/><rFont val="Calibri"/></rPr>',
        get parsedModel() {
            return this.preparedModel;
        },
        tests: ["render", "renderIn", "parse"],
    },
    {
        title: "boolean flags with an explicit val",
        create() {
            return new FontXform();
        },
        xml:
            '<font><b val="0"/><i val="false"/><strike val="1"/><outline val="true"/><shadow/>' +
            '<condense val="0"/><extend val="false"/><sz val="11"/></font>',
        parsedModel: { strike: true, outline: true, shadow: true, size: 11 },
        tests: ["parse"],
    },
    {
        title: "a size with decimals",
        create() {
            return new FontXform();
        },
        preparedModel: { size: 10.5, name: "Arial" },
        xml: '<font><sz val="10.5"/><name val="Arial"/></font>',
        get parsedModel() {
            return this.preparedModel;
        },
        tests: ["render", "renderIn", "parse"],
    },
];

describe("FontXform", () => {
    testXformHelper(expectations);
});
