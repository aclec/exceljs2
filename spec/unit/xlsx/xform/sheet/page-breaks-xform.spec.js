const testXformHelper = require("../test-xform-helper");

const PageBreaksXform = verquire("xlsx/xform/sheet/page-breaks-xform");
const RowBreaksXform = verquire("xlsx/xform/sheet/row-breaks-xform");

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
];

describe("PageBreaksXform", () => {
    testXformHelper(expectations);
});
