const testXformHelper = require("../test-xform-helper");

const PageSetupXform = verquire("xlsx/xform/sheet/page-setup-xform");

// -  "blackAndWhite": false
// -  "cellComments": "None"
// -  "draft": false
// -  "errors": "displayed"
// -  "fitToHeight": 1
// -  "fitToWidth": 1
// -  "horizontalDpi": "4294967295"
// +  "horizontalDpi": 4294967295
// "orientation": "portrait"
// -  "pageOrder": "downThenOver"
// "paperSize": 9
// -  "scale": 100
// -  "verticalDpi": "4294967295"
// +  "verticalDpi": 4294967295

const expectations = [
    {
        title: "normal",
        create: () => new PageSetupXform(),
        preparedModel: {
            paperSize: 9,
            orientation: "portrait",
            horizontalDpi: 4294967295,
            verticalDpi: 4294967295,
        },
        xml: '<pageSetup paperSize="9" orientation="portrait" horizontalDpi="4294967295" verticalDpi="4294967295"/>',
        parsedModel: {
            paperSize: 9,
            orientation: "portrait",
            horizontalDpi: 4294967295,
            verticalDpi: 4294967295,
            usePrinterDefaults: false,
            copies: 1,
            blackAndWhite: false,
            cellComments: "None",
            draft: false,
            errors: "displayed",
            fitToHeight: 1,
            fitToWidth: 1,
            pageOrder: "downThenOver",
            scale: 100,
        },
        tests: ["render", "renderIn", "parse"],
    },
    {
        title: "options",
        create: () => new PageSetupXform(),
        preparedModel: {
            paperSize: 119,
            pageOrder: "overThenDown",
            orientation: "portrait",
            blackAndWhite: true,
            draft: true,
            cellComments: "atEnd",
            errors: "dash",
        },
        xml: '<pageSetup paperSize="119" pageOrder="overThenDown" orientation="portrait" blackAndWhite="1" draft="1" cellComments="atEnd" errors="dash"/>',
        parsedModel: {
            paperSize: 119,
            pageOrder: "overThenDown",
            orientation: "portrait",
            blackAndWhite: true,
            draft: true,
            cellComments: "atEnd",
            errors: "dash",
            usePrinterDefaults: false,
            copies: 1,
            horizontalDpi: 4294967295,
            verticalDpi: 4294967295,
            fitToHeight: 1,
            fitToWidth: 1,
            scale: 100,
        },
        tests: ["render", "renderIn", "parse"],
    },
    {
        title: "defaults",
        create: () => new PageSetupXform(),
        preparedModel: {
            pageOrder: "downThenOver",
            orientation: "portrait",
            cellComments: "None",
            errors: "displayed",
        },
        xml: '<pageSetup orientation="portrait"/>',
        parsedModel: {
            pageOrder: "downThenOver",
            orientation: "portrait",
            cellComments: "None",
            errors: "displayed",
            usePrinterDefaults: false,
            copies: 1,
            horizontalDpi: 4294967295,
            verticalDpi: 4294967295,
            blackAndWhite: false,
            draft: false,
            fitToHeight: 1,
            fitToWidth: 1,
            scale: 100,
        },
        tests: ["render", "renderIn", "parse"],
    },
    {
        title: "scale and fit",
        create: () => new PageSetupXform(),
        preparedModel: {
            paperSize: 119,
            scale: 95,
            fitToWidth: 2,
            fitToHeight: 3,
            orientation: "landscape",
        },
        xml: '<pageSetup paperSize="119" scale="95" fitToWidth="2" fitToHeight="3" orientation="landscape"/>',
        parsedModel: {
            paperSize: 119,
            scale: 95,
            fitToWidth: 2,
            fitToHeight: 3,
            orientation: "landscape",
            usePrinterDefaults: false,
            copies: 1,
            horizontalDpi: 4294967295,
            verticalDpi: 4294967295,
            blackAndWhite: false,
            cellComments: "None",
            draft: false,
            errors: "displayed",
            pageOrder: "downThenOver",
        },
        tests: ["render", "renderIn", "parse"],
    },
];

describe("PageSetupXform", () => {
    testXformHelper(expectations);

    describe("cellComments", () => {
        it("writes the ST_CellComments value asDisplayed", () => {
            const xform = new PageSetupXform();
            expect(xform.toXml({ cellComments: "asDisplayed" })).to.equal('<pageSetup cellComments="asDisplayed"/>');
        });

        it("writes the legacy asDisplyed (sic) as asDisplayed", () => {
            const xform = new PageSetupXform();
            expect(xform.toXml({ cellComments: "asDisplyed" })).to.equal('<pageSetup cellComments="asDisplayed"/>');
        });

        it("does not write None, the default", () => {
            const xform = new PageSetupXform();
            expect(xform.toXml({ cellComments: "None", paperSize: 9 })).to.equal('<pageSetup paperSize="9"/>');
        });
    });

    describe("first page number", () => {
        function parse(attributes) {
            const xform = new PageSetupXform();
            xform.parseOpen({ name: "pageSetup", attributes });
            const { firstPageNumber, useFirstPageNumber } = xform.model;
            return { firstPageNumber, useFirstPageNumber, keys: Object.keys(xform.model) };
        }

        it("reads no first page number (automatic numbering) without the attributes", () => {
            for (const attributes of [{}, { useFirstPageNumber: "0" }, { useFirstPageNumber: "false" }]) {
                const { keys } = parse(attributes);
                expect(keys).to.not.include("firstPageNumber");
                expect(keys).to.not.include("useFirstPageNumber");
            }
        });

        it("reads a first page number, used or not", () => {
            expect(parse({ firstPageNumber: "5" })).to.include({ firstPageNumber: 5, useFirstPageNumber: false });
            expect(parse({ firstPageNumber: "5", useFirstPageNumber: "0" })).to.include({ firstPageNumber: 5, useFirstPageNumber: false });
            expect(parse({ firstPageNumber: "5", useFirstPageNumber: "1" })).to.include({ firstPageNumber: 5, useFirstPageNumber: true });
            expect(parse({ firstPageNumber: "5", useFirstPageNumber: "true" })).to.include({
                firstPageNumber: 5,
                useFirstPageNumber: true,
            });
            // the CT_PageSetup default firstPageNumber, used
            expect(parse({ useFirstPageNumber: "1" })).to.include({ firstPageNumber: 1, useFirstPageNumber: true });
        });

        it("writes useFirstPageNumber with a first page number, unless the model says it is not used", () => {
            const xform = new PageSetupXform();
            expect(xform.toXml({ firstPageNumber: 3 })).to.equal('<pageSetup firstPageNumber="3" useFirstPageNumber="1"/>');
            expect(xform.toXml({ firstPageNumber: 3, useFirstPageNumber: true })).to.equal(
                '<pageSetup firstPageNumber="3" useFirstPageNumber="1"/>'
            );
            expect(xform.toXml({ firstPageNumber: 3, useFirstPageNumber: false })).to.equal('<pageSetup firstPageNumber="3"/>');
            expect(xform.toXml({ useFirstPageNumber: true })).to.equal('<pageSetup useFirstPageNumber="1"/>');
            expect(xform.toXml({ useFirstPageNumber: false, paperSize: 9 })).to.equal('<pageSetup paperSize="9"/>');
        });
    });
});
