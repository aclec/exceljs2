const fs = require("fs");

const testXformHelper = require("../test-xform-helper");

const Enums = verquire("doc/enums");
const XmlStream = verquire("utils/xml-stream");
const WorksheetXform = verquire("xlsx/xform/sheet/worksheet-xform");
const RelType = verquire("xlsx/rel-type");

const SharedStringsXform = verquire("xlsx/xform/strings/shared-strings-xform");
const StylesXform = verquire("xlsx/xform/style/styles-xform");

const fakeStyles = {
    addStyleModel(style, cellType) {
        if (cellType === Enums.ValueType.Date) {
            return 1;
        }
        if (style && style.font) {
            return 2;
        }
        return 0;
    },
    getStyleModel(id) {
        switch (id) {
            case 1:
                return { numFmt: "mm-dd-yy" };
            case 2:
                return {
                    font: {
                        underline: true,
                        size: 11,
                        color: { theme: 10 },
                        name: "Calibri",
                        family: 2,
                        scheme: "minor",
                    },
                };
            default:
                return null;
        }
    },
};

const fakeHyperlinkMap = {
    B6: "https://www.npmjs.com/package/exceljs",
};

function fixDate(model) {
    model.rows[3].cells[1].value = new Date(model.rows[3].cells[1].value);
    return model;
}

const expectations = [
    {
        title: "Sheet 1",
        create: () => new WorksheetXform(),
        initialModel: fixDate(require("./data/sheet.1.0.json")),
        preparedModel: fixDate(require("./data/sheet.1.1.json")),
        xml: fs.readFileSync(`${__dirname}/data/sheet.1.2.xml`).toString(),
        parsedModel: require("./data/sheet.1.3.json"),
        reconciledModel: fixDate(require("./data/sheet.1.4.json")),
        tests: ["prepare", "render", "parse"],
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
        title: "Sheet 2 - Data Validations",
        create: () => new WorksheetXform(),
        initialModel: require("./data/sheet.2.0.json"),
        preparedModel: require("./data/sheet.2.1.json"),
        xml: fs.readFileSync(`${__dirname}/data/sheet.2.2.xml`).toString(),
        tests: ["prepare", "render"],
        options: {
            styles: new StylesXform(true),
            sharedStrings: new SharedStringsXform(),
            hyperlinks: [],
            formulae: {},
            siFormulae: 0,
        },
    },
    {
        title: "Sheet 3 - Empty Sheet",
        create: () => new WorksheetXform(),
        preparedModel: require("./data/sheet.3.1.json"),
        xml: fs.readFileSync(`${__dirname}/data/sheet.3.2.xml`).toString(),
        tests: ["render"],
        options: {
            styles: new StylesXform(true),
            sharedStrings: new SharedStringsXform(),
            hyperlinks: [],
        },
    },
    {
        title: "Sheet 5 - Shared Formulas",
        create: () => new WorksheetXform(),
        initialModel: require("./data/sheet.5.0.json"),
        preparedModel: require("./data/sheet.5.1.json"),
        xml: fs.readFileSync(`${__dirname}/data/sheet.5.2.xml`).toString(),
        parsedModel: require("./data/sheet.5.3.json"),
        reconciledModel: require("./data/sheet.5.4.json"),
        tests: ["prepare-render", "parse"],
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
        title: "Sheet 6 - AutoFilter",
        create: () => new WorksheetXform(),
        preparedModel: require("./data/sheet.6.1.json"),
        xml: fs.readFileSync(`${__dirname}/data/sheet.6.2.xml`).toString(),
        parsedModel: require("./data/sheet.6.3.json"),
        tests: ["render", "parse"],
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
        title: "Sheet 7 - Row Breaks",
        create: () => new WorksheetXform(),
        initialModel: require("./data/sheet.7.0.json"),
        preparedModel: require("./data/sheet.7.1.json"),
        xml: fs.readFileSync(`${__dirname}/data/sheet.7.2.xml`).toString(),
        tests: ["prepare", "render"],
        options: {
            sharedStrings: new SharedStringsXform(),
            hyperlinks: [],
            hyperlinkMap: fakeHyperlinkMap,
            styles: fakeStyles,
            formulae: {},
            siFormulae: 0,
        },
    },
];

describe("WorksheetXform", () => {
    testXformHelper(expectations);

    it("hyperlinks must be after dataValidations", () => {
        const xform = new WorksheetXform();
        const model = require("./data/sheet.4.0.json");
        const xmlStream = new XmlStream();
        const options = {
            styles: new StylesXform(true),
            sharedStrings: new SharedStringsXform(),
            hyperlinks: [],
        };
        xform.prepare(model, options);
        xform.render(xmlStream, model);

        const { xml } = xmlStream;
        const iHyperlinks = xml.indexOf("hyperlinks");
        const iDataValidations = xml.indexOf("dataValidations");
        expect(iHyperlinks).not.to.equal(-1);
        expect(iDataValidations).not.to.equal(-1);
        expect(iHyperlinks).to.be.greaterThan(iDataValidations);
    });

    it("conditionalFormattings must be before dataValidations", () => {
        const xform = new WorksheetXform();
        const model = require("./data/sheet.4.0.json");
        const xmlStream = new XmlStream();
        const options = {
            styles: new StylesXform(true),
            hyperlinks: [],
        };
        xform.prepare(model, options);
        xform.render(xmlStream, model);

        const { xml } = xmlStream;
        const iConditionalFormatting = xml.indexOf("conditionalFormatting");
        const iDataValidations = xml.indexOf("dataValidations");
        expect(iConditionalFormatting).not.to.equal(-1);
        expect(iDataValidations).not.to.equal(-1);
        expect(iConditionalFormatting).to.be.lessThan(iDataValidations);
    });

    describe("formula conditional formats that Excel writes in the x14 extension", () => {
        const X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
        const XM_NS = "http://schemas.microsoft.com/office/excel/2006/main";

        function renderSheet(conditionalFormattings) {
            const xform = new WorksheetXform();
            const model = JSON.parse(JSON.stringify(require("./data/sheet.4.0.json")));
            model.conditionalFormattings = conditionalFormattings;
            const xmlStream = new XmlStream();
            const options = {
                styles: new StylesXform(true),
                sharedStrings: new SharedStringsXform(),
                hyperlinks: [],
            };
            xform.prepare(model, options);
            xform.render(xmlStream, model);
            return xmlStream.xml;
        }

        const fill = { type: "pattern", pattern: "solid", bgColor: { argb: "FFFFFF00" } };

        it("writes a rule referencing another sheet in the extension only, after the other parts", () => {
            const xml = renderSheet([
                {
                    ref: "A1:A5",
                    rules: [
                        { type: "expression", formulae: ["A1>1"], style: { fill } },
                        { type: "expression", formulae: ["'Other sheet'!$A$1=A1"], style: { fill } },
                    ],
                },
                { ref: "B1:B5", rules: [{ type: "cellIs", operator: "greaterThan", formulae: ["Other!$B$1"], style: { fill } }] },
            ]);

            // the main section keeps the same-sheet rule only
            expect(xml).to.include(
                '<conditionalFormatting sqref="A1:A5"><cfRule type="expression" dxfId="0" priority="1"><formula>A1&gt;1</formula></cfRule></conditionalFormatting>'
            );
            expect(xml).not.to.include('<conditionalFormatting sqref="B1:B5">');
            expect(xml.match(/<cfRule /g)).to.have.length(1);

            const dxf = '<x14:dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFFFF00"/></patternFill></fill></x14:dxf>';
            const id = "\\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\\}";
            const escape = (text) => text.replace(/[$()[\]{}|*+?.^\\]/g, "\\$&");
            const extLst = new RegExp(
                `${escape(`<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="${X14_NS}"><x14:conditionalFormattings>`)}` +
                    `${escape(`<x14:conditionalFormatting xmlns:xm="${XM_NS}">`)}` +
                    `<x14:cfRule type="expression" priority="2" id="${id}">` +
                    `${escape(`<xm:f>&apos;Other sheet&apos;!$A$1=A1</xm:f>${dxf}</x14:cfRule><xm:sqref>A1:A5</xm:sqref></x14:conditionalFormatting>`)}` +
                    `${escape(`<x14:conditionalFormatting xmlns:xm="${XM_NS}">`)}` +
                    `<x14:cfRule type="cellIs" priority="3" operator="greaterThan" id="${id}">` +
                    `${escape(`<xm:f>Other!$B$1</xm:f>${dxf}</x14:cfRule><xm:sqref>B1:B5</xm:sqref></x14:conditionalFormatting>`)}` +
                    `${escape("</x14:conditionalFormattings></ext></extLst></worksheet>")}$`
            );
            expect(xml).to.match(extLst);
        });

        it("writes a rule flagged x14 in the extension, and a same-sheet rule in the main section only", () => {
            const flagged = renderSheet([{ ref: "A1", rules: [{ type: "expression", x14: true, formulae: ["A1=1"] }] }]);
            expect(flagged).not.to.include("<conditionalFormatting");
            expect(flagged).to.match(/<x14:cfRule type="expression" priority="1" id="\{[0-9A-F-]+\}"><xm:f>A1=1<\/xm:f><\/x14:cfRule>/);

            const local = renderSheet([{ ref: "A1", rules: [{ type: "expression", formulae: ['A1="Hi!"'] }] }]);
            expect(local).to.include('<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1">');
            expect(local).not.to.include("extLst");
        });
    });

    it("indexes hyperlinks by the cell references of the file without inheriting Object.prototype members", () => {
        const xform = new WorksheetXform();
        const model = {
            hyperlinks: [
                { address: "__proto__", rId: "rId1" },
                { address: "B2", target: "Sheet2!A1" },
            ],
            relationships: [{ Id: "rId1", Type: RelType.Hyperlink, Target: "http://example.com/", TargetMode: "External" }],
            rows: [],
            conditionalFormattings: [],
        };
        const options = { styles: new StylesXform(true) };
        xform.reconcile(model, options);

        const { hyperlinkMap } = options;
        expect(Object.getOwnPropertyDescriptor(hyperlinkMap, "__proto__").value).to.equal("http://example.com/");
        expect(hyperlinkMap.B2).to.equal("#Sheet2!A1");
        expect(hyperlinkMap.toString).to.equal(undefined);
        expect(hyperlinkMap.constructor).to.equal(undefined);
    });
});
