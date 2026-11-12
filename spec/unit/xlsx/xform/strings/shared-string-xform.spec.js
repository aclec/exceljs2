const testXformHelper = require("../test-xform-helper");

const SharedStringXform = verquire("xlsx/xform/strings/shared-string-xform");

const expectations = [
    {
        // ECMA-376 ST_Xstring: _xHHHH_ encodes a character, so a literal one is written with its _ escaped as _x005F_
        title: "text holding _xHHHH_ sequences",
        create() {
            return new SharedStringXform();
        },
        preparedModel: "a_x0041_b _x00e9_ _x005F_ _x12_ x_ _x0041_x0042_",
        xml: "<si><t>a_x005F_x0041_b _x005F_x00e9_ _x005F_x005F_ _x12_ x_ _x005F_x0041_x005F_x0042_</t></si>",
        parsedModel: "a_x0041_b _x00e9_ _x005F_ _x12_ x_ _x0041_x0042_",
        tests: ["render", "renderIn", "parse"],
    },
    {
        title: "text",
        create() {
            return new SharedStringXform();
        },
        preparedModel: "Hello, World!",
        xml: "<si><t>Hello, World!</t></si>",
        parsedModel: "Hello, World!",
        tests: ["render", "renderIn", "parse"],
    },
    {
        title: "rich text",
        create() {
            return new SharedStringXform();
        },
        preparedModel: {
            richText: [
                {
                    font: {
                        size: 11,
                        bold: true,
                        color: { theme: 1 },
                        name: "Calibri",
                        family: 2,
                        scheme: "minor",
                    },
                    text: "Bold,",
                },
            ],
        },
        xml: '<si><r><rPr><b/><color theme="1"/><family val="2"/><scheme val="minor"/><sz val="11"/><rFont val="Calibri"/></rPr><t>Bold,</t></r></si>',
        parsedModel: {
            richText: [
                {
                    font: {
                        size: 11,
                        bold: true,
                        color: { theme: 1 },
                        name: "Calibri",
                        family: 2,
                        scheme: "minor",
                    },
                    text: "Bold,",
                },
            ],
        },
        tests: ["render", "renderIn", "parse"],
    },
    {
        title: "richText is empty",
        create() {
            return new SharedStringXform();
        },
        preparedModel: {
            richText: [],
        },
        xml: "<si><t></t></si>",
        parsedModel: "",
        tests: ["parse"],
    },
    {
        title: "text + phonetic",
        create() {
            return new SharedStringXform();
        },
        preparedModel: {
            text: "Hello, World!",
            phoneticText: { text: "Helow woruld" },
        },
        xml: '<si><t>Hello, World!</t><rPh eb="0" sb="0"><t>Helow woruld</t></rPh></si>',
        parsedModel: "Hello, World!",
        tests: ["parse"],
    },
    {
        title: "Kanji + Katakana",
        create() {
            return new SharedStringXform();
        },
        preparedModel: {
            text: "役割",
            phoneticText: {
                sb: 0,
                eb: 2,
                text: "ヤクワリ",
                properties: { fontId: 1 },
            },
        },
        xml: '<si><t>役割</t><rPh sb="0" eb="2"><t>ヤクワリ</t></rPh><phoneticPr fontId="1" /></si>',
        parsedModel: "役割",
        tests: ["parse"],
    },
    {
        title: "text with newline",
        create() {
            return new SharedStringXform();
        },
        preparedModel: "Hello,\nWorld!",
        xml: '<si><t xml:space="preserve">Hello,\nWorld!</t></si>',
        parsedModel: "Hello,\nWorld!",
        tests: ["render", "renderIn", "parse"],
    },
    {
        title: "rich text with empty runs",
        create() {
            return new SharedStringXform();
        },
        preparedModel: {
            richText: [{ font: { bold: true }, text: "" }, { text: "Hello" }, { text: "" }, { font: { italic: true } }],
        },
        xml: "<si><r><t>Hello</t></r></si>",
        tests: ["render", "renderIn"],
    },
    {
        title: "rich text with only empty runs",
        create() {
            return new SharedStringXform();
        },
        preparedModel: {
            richText: [{ font: { bold: true }, text: "" }, { text: "" }],
        },
        xml: "<si><t/></si>",
        tests: ["render", "renderIn"],
    },
    {
        title: "plain text followed by rich text runs",
        create() {
            return new SharedStringXform();
        },
        xml: "<si><t>Plain </t><r><rPr><b/></rPr><t>bold</t></r><r><t> end</t></r></si>",
        parsedModel: {
            richText: [{ text: "Plain " }, { font: { bold: true }, text: "bold" }, { text: " end" }],
        },
        tests: ["parse"],
    },
];

describe("SharedStringXform", () => {
    testXformHelper(expectations);

    it("escapes _xHHHH_ sequences in linear time", () => {
        ["_", "_x", "_x0041", "_x0041_", "_x00_"].forEach((unit) => {
            const text = unit.repeat(10000);
            const start = Date.now();
            const xml = new SharedStringXform().toXml(text);
            expect(Date.now() - start, unit).to.be.below(50);
            expect(xml.length).to.be.at.least(text.length);
        });
    });
});
