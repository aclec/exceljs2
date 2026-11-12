const testXformHelper = require("../test-xform-helper");

const CellXform = verquire("xlsx/xform/sheet/cell-xform");
const SharedStringsXform = verquire("xlsx/xform/strings/shared-strings-xform");
const Enums = verquire("doc/enums");

const fakeStyles = {
    addStyleModel(style, effectiveType) {
        if (effectiveType === Enums.ValueType.Date) {
            return 1;
        }
        return 0;
    },
    getStyleModel(styleId) {
        switch (styleId) {
            case 1:
                return { numFmt: "mm-dd-yy" };
            default:
                return null;
        }
    },
};

const fakeHyperlinkMap = {
    H1: { target: "http://www.foo.com" },
};

const expectations = [
    {
        title: "Styled Null",
        create() {
            return new CellXform();
        },
        preparedModel: { address: "A1", type: Enums.ValueType.Null, styleId: 1 },
        xml: '<c r="A1" s="1" />',
        parsedModel: { address: "A1", type: Enums.ValueType.Null, styleId: 1 },
        tests: ["render", "renderIn", "parse"],
    },
    {
        title: "Number",
        create() {
            return new CellXform();
        },
        preparedModel: { address: "A1", type: Enums.ValueType.Number, value: 5 },
        parsedModel: { address: "A1", type: Enums.ValueType.Number, value: 5 },
        xml: '<c r="A1"><v>5</v></c>',
        tests: ["render", "renderIn", "parse"],
    },
    {
        title: "Boolean",
        create() {
            return new CellXform();
        },
        preparedModel: { address: "A1", type: Enums.ValueType.Boolean, value: true },
        parsedModel: { address: "A1", type: Enums.ValueType.Boolean, value: true },
        xml: '<c r="A1" t="b"><v>1</v></c>',
        tests: ["render", "renderIn", "parse"],
    },
    {
        title: "Error",
        create() {
            return new CellXform();
        },
        preparedModel: {
            address: "A1",
            type: Enums.ValueType.Error,
            value: { error: "#N/A" },
        },
        parsedModel: {
            address: "A1",
            type: Enums.ValueType.Error,
            value: { error: "#N/A" },
        },
        xml: '<c r="A1" t="e"><v>#N/A</v></c>',
        tests: ["render", "renderIn", "parse"],
    },
    {
        title: "String",
        create() {
            return new CellXform();
        },
        initialModel: { address: "A1", type: Enums.ValueType.String, value: "Foo" },
        preparedModel: { address: "A1", type: Enums.ValueType.String, value: "Foo" },
        xml: '<c r="A1" t="str"><v>Foo</v></c>',
        parsedModel: { address: "A1", type: Enums.ValueType.String, value: "Foo" },
        reconciledModel: {
            address: "A1",
            type: Enums.ValueType.String,
            value: "Foo",
        },
        tests: ["prepare", "render", "renderIn", "parse", "reconcile"],
        options: { hyperlinkMap: fakeHyperlinkMap, styles: fakeStyles },
    },
    {
        title: "String with Invalid Number",
        create() {
            return new CellXform();
        },
        initialModel: {
            address: "A1",
            type: Enums.ValueType.String,
            value: "6E1000",
        },
        preparedModel: {
            address: "A1",
            type: Enums.ValueType.String,
            value: "6E1000",
        },
        xml: '<c r="A1" t="str"><v>6E1000</v></c>',
        parsedModel: { address: "A1", type: Enums.ValueType.String, value: "6E1000" },
        reconciledModel: {
            address: "A1",
            type: Enums.ValueType.String,
            value: "6E1000",
        },
        tests: ["prepare", "render", "renderIn", "parse", "reconcile"],
        options: {
            hyperlinkMap: fakeHyperlinkMap,
            styles: fakeStyles,
        },
    },
    {
        title: "Inline String with plain text",
        create() {
            return new CellXform();
        },
        xml: '<c r="A1" t="inlineStr"><is><t>Foo</t></is></c>',
        parsedModel: { address: "A1", type: Enums.ValueType.String, value: "Foo" },
        reconciledModel: {
            address: "A1",
            type: Enums.ValueType.String,
            value: "Foo",
        },
        tests: ["parse", "reconcile"],
        options: { hyperlinkMap: fakeHyperlinkMap, styles: fakeStyles },
    },
    {
        title: "Inline String with plain text followed by rich text runs",
        create() {
            return new CellXform();
        },
        xml: '<c r="A1" t="inlineStr"><is><t>Plain </t><r><rPr><b/></rPr><t>bold</t></r></is></c>',
        parsedModel: {
            address: "A1",
            type: Enums.ValueType.String,
            value: { richText: [{ text: "Plain " }, { font: { bold: true }, text: "bold" }] },
        },
        reconciledModel: {
            address: "A1",
            type: Enums.ValueType.RichText,
            value: { richText: [{ text: "Plain " }, { font: { bold: true }, text: "bold" }] },
        },
        tests: ["parse", "reconcile"],
        options: { hyperlinkMap: fakeHyperlinkMap, styles: fakeStyles },
    },
    {
        title: "Inline String with RichText",
        create() {
            return new CellXform();
        },
        initialModel: {
            address: "A1",
            type: Enums.ValueType.String,
            value: {
                richText: [
                    { font: { color: { argb: "FF0000" } }, text: "red" },
                    { font: { color: { argb: "00FF00" } }, text: "green" },
                ],
            },
        },
        preparedModel: {
            address: "A1",
            type: Enums.ValueType.String,
            value: {
                richText: [
                    { font: { color: { argb: "FF0000" } }, text: "red" },
                    { font: { color: { argb: "00FF00" } }, text: "green" },
                ],
            },
        },
        xml: '<c r="A1" t="inlineStr"><is><r><rPr><color rgb="FF0000"/></rPr><t>red</t></r><r><rPr><color rgb="00FF00"/></rPr><t>green</t></r></is></c>',
        parsedModel: {
            address: "A1",
            type: Enums.ValueType.String,
            value: {
                richText: [
                    { font: { color: { argb: "FF0000" } }, text: "red" },
                    { font: { color: { argb: "00FF00" } }, text: "green" },
                ],
            },
        },
        reconciledModel: {
            address: "A1",
            type: Enums.ValueType.RichText,
            value: {
                richText: [
                    { font: { color: { argb: "FF0000" } }, text: "red" },
                    { font: { color: { argb: "00FF00" } }, text: "green" },
                ],
            },
        },
        tests: ["prepare", "render", "renderIn", "parse", "reconcile"],
        options: { hyperlinkMap: fakeHyperlinkMap, styles: fakeStyles },
    },
    {
        title: "Shared String",
        create() {
            return new CellXform();
        },
        initialModel: { address: "A1", type: Enums.ValueType.String, value: "Foo" },
        preparedModel: {
            address: "A1",
            type: Enums.ValueType.String,
            value: "Foo",
            ssId: 0,
        },
        xml: '<c r="A1" t="s"><v>0</v></c>',
        parsedModel: { address: "A1", type: Enums.ValueType.String, value: 0 },
        reconciledModel: {
            address: "A1",
            type: Enums.ValueType.String,
            value: "Foo",
        },
        tests: ["prepare", "render", "renderIn", "parse", "reconcile"],
        options: {
            sharedStrings: new SharedStringsXform(),
            hyperlinkMap: fakeHyperlinkMap,
            styles: fakeStyles,
        },
    },
    {
        title: "Shared String with RichText",
        create() {
            return new CellXform();
        },
        initialModel: {
            address: "A1",
            type: Enums.ValueType.RichText,
            value: {
                richText: [
                    { font: { color: { argb: "FF0000" } }, text: "red" },
                    { font: { color: { argb: "00FF00" } }, text: "green" },
                ],
            },
        },
        preparedModel: {
            address: "A1",
            type: Enums.ValueType.RichText,
            value: {
                richText: [
                    { font: { color: { argb: "FF0000" } }, text: "red" },
                    { font: { color: { argb: "00FF00" } }, text: "green" },
                ],
            },
            ssId: 0,
        },
        xml: '<c r="A1" t="s"><v>0</v></c>',
        parsedModel: {
            address: "A1",
            type: Enums.ValueType.String,
            value: 0,
        },
        reconciledModel: {
            address: "A1",
            type: Enums.ValueType.RichText,
            value: {
                richText: [
                    { font: { color: { argb: "FF0000" } }, text: "red" },
                    { font: { color: { argb: "00FF00" } }, text: "green" },
                ],
            },
        },
        tests: ["prepare", "render", "renderIn", "parse", "reconcile"],
        options: {
            sharedStrings: new SharedStringsXform(),
            hyperlinkMap: fakeHyperlinkMap,
            styles: fakeStyles,
        },
    },
    {
        title: "Date",
        create() {
            return new CellXform();
        },
        initialModel: {
            address: "A1",
            type: Enums.ValueType.Date,
            value: new Date("2016-06-09T00:00:00.000Z"),
        },
        preparedModel: {
            address: "A1",
            type: Enums.ValueType.Date,
            value: new Date("2016-06-09T00:00:00.000Z"),
            styleId: 1,
        },
        xml: '<c r="A1" s="1"><v>42530</v></c>',
        parsedModel: {
            address: "A1",
            type: Enums.ValueType.Number,
            value: 42530,
            styleId: 1,
        },
        reconciledModel: {
            address: "A1",
            type: Enums.ValueType.Date,
            value: new Date("2016-06-09T00:00:00.000Z"),
            style: { numFmt: "mm-dd-yy" },
        },
        tests: ["prepare", "render", "renderIn", "parse", "reconcile"],
        options: {
            sharedStrings: new SharedStringsXform(),
            hyperlinks: [],
            hyperlinkMap: fakeHyperlinkMap,
            styles: fakeStyles,
        },
    },
    {
        title: "Hyperlink",
        create() {
            return new CellXform();
        },
        initialModel: {
            address: "H1",
            type: Enums.ValueType.Hyperlink,
            hyperlink: "http://www.foo.com",
            text: "www.foo.com",
        },
        preparedModel: {
            address: "H1",
            type: Enums.ValueType.Hyperlink,
            hyperlink: "http://www.foo.com",
            text: "www.foo.com",
            ssId: 0,
        },
        xml: '<c r="H1" t="s"><v>0</v></c>',
        parsedModel: { address: "H1", type: Enums.ValueType.String, value: 0 },
        reconciledModel: {
            address: "H1",
            type: Enums.ValueType.Hyperlink,
            text: "www.foo.com",
            hyperlink: "http://www.foo.com",
        },
        tests: ["prepare", "render", "renderIn", "parse", "reconcile"],
        options: {
            sharedStrings: new SharedStringsXform(),
            hyperlinks: [],
            hyperlinkMap: fakeHyperlinkMap,
            styles: fakeStyles,
        },
    },
    {
        title: "String Formula",
        create() {
            return new CellXform();
        },
        initialModel: {
            address: "A1",
            type: Enums.ValueType.Formula,
            formula: "A2",
            result: "Foo",
        },
        preparedModel: {
            address: "A1",
            type: Enums.ValueType.Formula,
            formula: "A2",
            result: "Foo",
        },
        xml: '<c r="A1" t="str"><f>A2</f><v>Foo</v></c>',
        parsedModel: {
            address: "A1",
            type: Enums.ValueType.Formula,
            formula: "A2",
            result: "Foo",
        },
        reconciledModel: {
            address: "A1",
            type: Enums.ValueType.Formula,
            formula: "A2",
            result: "Foo",
        },
        tests: ["prepare", "render", "renderIn", "parse", "reconcile"],
        options: {
            sharedStrings: new SharedStringsXform(),
            hyperlinks: [],
            hyperlinkMap: fakeHyperlinkMap,
            styles: fakeStyles,
            formulae: {},
            siFormulae: 0,
        },
    },
    {
        title: "Number Formula",
        create() {
            return new CellXform();
        },
        preparedModel: {
            address: "A1",
            type: Enums.ValueType.Formula,
            formula: "A2",
            result: 7,
        },
        xml: '<c r="A1"><f>A2</f><v>7</v></c>',
        parsedModel: {
            address: "A1",
            type: Enums.ValueType.Formula,
            formula: "A2",
            result: 7,
        },
        tests: ["render", "renderIn", "parse"],
        options: { formulae: {}, siFormulae: 0 },
    },
    {
        title: "Master Shared Formula",
        create() {
            return new CellXform();
        },
        preparedModel: {
            address: "A2",
            type: Enums.ValueType.Formula,
            shareType: "shared",
            ref: "A2:B2",
            formula: "A1",
            result: 2,
            si: 0,
        },
        xml: '<c r="A2"><f t="shared" ref="A2:B2" si="0">A1</f><v>2</v></c>',
        parsedModel: {
            address: "A2",
            type: Enums.ValueType.Formula,
            shareType: "shared",
            ref: "A2:B2",
            formula: "A1",
            result: 2,
            si: "0",
        },
        reconciledModel: {
            address: "A2",
            type: Enums.ValueType.Formula,
            shareType: "shared",
            ref: "A2:B2",
            formula: "A1",
            result: 2,
        },
        tests: ["render", "renderIn", "parse", "reconcile"],
        options: {
            styles: fakeStyles,
            hyperlinks: [],
            hyperlinkMap: fakeHyperlinkMap,
            formulae: {},
            siFormulae: 0,
        },
    },
    {
        title: "Shared Formula Slave",
        create() {
            return new CellXform();
        },
        initialModel: {
            address: "A2",
            type: Enums.ValueType.Formula,
            sharedFormula: "A1",
            result: 2,
        },
        preparedModel: {
            address: "A2",
            type: Enums.ValueType.Formula,
            sharedFormula: "A1",
            result: 2,
            si: 0,
        },
        xml: '<c r="A2"><f t="shared" si="0" /><v>2</v></c>',
        parsedModel: {
            address: "A2",
            type: Enums.ValueType.Formula,
            result: 2,
            shareType: "shared",
            si: "0",
        },
        reconciledModel: {
            address: "A2",
            type: Enums.ValueType.Formula,
            result: 2,
            sharedFormula: "A1",
        },
        tests: ["prepare", "render", "renderIn", "parse", "reconcile"],
        options: {
            styles: fakeStyles,
            hyperlinks: [],
            hyperlinkMap: fakeHyperlinkMap,
            formulae: {
                A1: {
                    address: "A1",
                    type: Enums.ValueType.Formula,
                    formula: "ROW()",
                    result: 1,
                    si: 0,
                },
                0: "A1",
            },
            siFormulae: 1,
        },
    },
    {
        title: "Array Shared Formula",
        create() {
            return new CellXform();
        },
        preparedModel: {
            address: "A2",
            type: Enums.ValueType.Formula,
            shareType: "array",
            ref: "A2:B2",
            formula: "A1",
            result: 2,
        },
        xml: '<c r="A2"><f t="array" ref="A2:B2">A1</f><v>2</v></c>',
        parsedModel: {
            address: "A2",
            type: Enums.ValueType.Formula,
            shareType: "array",
            ref: "A2:B2",
            formula: "A1",
            result: 2,
        },
        reconciledModel: {
            address: "A2",
            type: Enums.ValueType.Formula,
            shareType: "array",
            ref: "A2:B2",
            formula: "A1",
            result: 2,
        },
        tests: ["render", "renderIn", "parse", "reconcile"],
        options: {
            styles: fakeStyles,
            hyperlinks: [],
            hyperlinkMap: fakeHyperlinkMap,
            formulae: {},
            siFormulae: 0,
        },
    },
];

describe("CellXform", () => {
    testXformHelper(expectations);
});

describe("CellXform formula cells", () => {
    const XmlStream = verquire("utils/xml-stream");

    function prepareAndRender(model) {
        const xform = new CellXform();
        xform.prepare(model, { styles: fakeStyles, hyperlinks: [], comments: [], formulae: {}, siFormulae: 0 });
        const xmlStream = new XmlStream();
        xform.render(xmlStream, model);
        return xmlStream.xml;
    }

    it("writes a formula cell without formula text as its cached value (no empty <f>)", () => {
        // CT_CellFormula may be empty only for a clone of a shared formula: an empty normal formula is not a formula
        const cases = [
            [{ formula: "", result: 5 }, '<c r="A1"><v>5</v></c>'],
            [{ formula: "", result: "x" }, '<c r="A1" t="str"><v>x</v></c>'],
            [{ formula: "", result: true }, '<c r="A1" t="b"><v>1</v></c>'],
            [{ formula: "", result: { error: "#N/A" } }, '<c r="A1" t="e"><v>#N/A</v></c>'],
            [{ formula: "" }, '<c r="A1"/>'],
            // a clone of a shared formula whose master is not in the file (read as a formula without text)
            [{ formula: undefined, sharedFormula: undefined, result: 3 }, '<c r="A1"><v>3</v></c>'],
        ];
        for (const [value, xml] of cases) {
            expect(prepareAndRender({ address: "A1", type: Enums.ValueType.Formula, ...value }), JSON.stringify(value)).to.equal(xml);
        }
    });

    it("still writes the empty <f> of a shared formula clone", () => {
        const xform = new CellXform();
        const xmlStream = new XmlStream();
        xform.render(xmlStream, { address: "A2", type: Enums.ValueType.Formula, si: 0, result: 2 });
        expect(xmlStream.xml).to.equal('<c r="A2"><f t="shared" si="0"/><v>2</v></c>');
    });

    it("tells the type of a formula result by the presence and type of its markers, as Cell does", () => {
        // a hyperlink, formula or shared string object is not a formula result, whatever the truthiness of its markers
        for (const result of [{ text: "", hyperlink: "http://x" }, { formula: "", result: 1 }, { sharedString: 0 }, { richText: [] }]) {
            expect(
                () => prepareAndRender({ address: "B3", type: Enums.ValueType.Formula, formula: "A1", result }),
                JSON.stringify(result)
            ).to.throw(Error, /Cell B3: the result of a formula must be/);
        }
    });
});
