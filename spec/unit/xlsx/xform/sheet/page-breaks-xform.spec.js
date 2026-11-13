const testXformHelper = require("../test-xform-helper");

const PageBreaksXform = verquire("xlsx/xform/sheet/page-breaks-xform");
const RowBreaksXform = verquire("xlsx/xform/sheet/row-breaks-xform");
const ColBreaksXform = verquire("xlsx/xform/sheet/col-breaks-xform");

const expectations = [
    {
        title: "one page break",
        create() {
            return new PageBreaksXform();
        },
        initialModel: { id: 2, max: 3, min: 1, man: 1 },
        preparedModel: { id: 2, max: 3, min: 1, man: 1 },
        xml: '<brk id="2" max="3" min="1" man="1"/>',
        parsedModel: { id: 2, max: 3, min: 1, man: 1 },
        tests: ["prepare", "render", "renderIn", "parse"],
    },
    {
        title: "a page break across all the columns",
        create() {
            return new PageBreaksXform();
        },
        preparedModel: { id: 5, max: 16383, man: 1 },
        xml: '<brk id="5" max="16383" man="1"/>',
        parsedModel: { id: 5, max: 16383, man: 1 },
        tests: ["render", "parse"],
    },
    {
        title: "the CT_Break defaults (an automatic break; pt is not kept)",
        create() {
            return new PageBreaksXform();
        },
        xml: '<brk pt="1"/>',
        parsedModel: { id: 0, max: 0, man: 0 },
        tests: ["parse"],
    },
    {
        title: "an xsd:boolean man",
        create() {
            return new PageBreaksXform();
        },
        xml: '<brk id="7" max="16383" man="true"/>',
        parsedModel: { id: 7, max: 16383, man: 1 },
        tests: ["parse"],
    },
    {
        title: "row breaks",
        create() {
            return new RowBreaksXform();
        },
        preparedModel: [
            { id: 3, max: 16383, man: 1 },
            { id: 5, max: 3, min: 1, man: 1 },
        ],
        xml: '<rowBreaks count="2" manualBreakCount="2"><brk id="3" max="16383" man="1"/><brk id="5" max="3" min="1" man="1"/></rowBreaks>',
        parsedModel: [
            { id: 3, max: 16383, man: 1 },
            { id: 5, max: 3, min: 1, man: 1 },
        ],
        tests: ["render", "parse"],
    },
    {
        title: "a page break across all the rows (a column break: max 1048575, the last 0-based row)",
        create() {
            // the brk xform of the column breaks
            return new ColBreaksXform().childXform;
        },
        preparedModel: { id: 3, max: 1048575, man: 1 },
        xml: '<brk id="3" max="1048575" man="1"/>',
        parsedModel: { id: 3, max: 1048575, man: 1 },
        tests: ["render", "parse"],
    },
    {
        title: "column breaks",
        create() {
            return new ColBreaksXform();
        },
        preparedModel: [
            { id: 3, max: 1048575, man: 1 },
            { id: 6, max: 9, min: 4, man: 1 },
            { id: 9, max: 1048575, man: 0 },
        ],
        xml:
            '<colBreaks count="3" manualBreakCount="2"><brk id="3" max="1048575" man="1"/>' +
            '<brk id="6" max="9" min="4" man="1"/><brk id="9" max="1048575" man="0"/></colBreaks>',
        parsedModel: [
            { id: 3, max: 1048575, man: 1 },
            { id: 6, max: 9, min: 4, man: 1 },
            { id: 9, max: 1048575, man: 0 },
        ],
        tests: ["render", "parse"],
    },
    {
        title: "column breaks: the values beyond the worksheet are its last column (id) or row (min, max)",
        create() {
            return new ColBreaksXform();
        },
        xml: '<colBreaks><brk id="20000" min="2000000" max="4294967295" man="1"/><brk id="5" max="16383"/></colBreaks>',
        parsedModel: [
            { id: 16384, max: 1048575, min: 1048575, man: 1 },
            { id: 5, max: 16383, man: 0 },
        ],
        tests: ["parse"],
    },
];

describe("PageBreaksXform", () => {
    testXformHelper(expectations);

    describe("ColBreaksXform", () => {
        it("renders nothing without a break", () => {
            expect(new ColBreaksXform().toXml([])).to.equal("");
            expect(new ColBreaksXform().toXml(undefined)).to.equal("");
        });

        it("renders at most 1023 breaks (the limit of Excel)", () => {
            const breaks = (n) => Array.from({ length: n }, (_, i) => ({ id: i + 1, max: 1048575, man: 1 }));
            expect(new ColBreaksXform().toXml(breaks(1023))).to.match(/^<colBreaks count="1023" manualBreakCount="1023">/);
            expect(() => new ColBreaksXform().toXml(breaks(1024))).to.throw(
                Error,
                "Too many column breaks in a worksheet: 1024 (Excel allows at most 1023)"
            );
        });
    });
});
