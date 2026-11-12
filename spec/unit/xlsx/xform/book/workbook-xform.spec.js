const fs = require("fs");

const testXformHelper = require("../test-xform-helper");

const WorkbookXform = verquire("xlsx/xform/book/workbook-xform");

const expectations = [
    {
        title: "book.1",
        create() {
            return new WorkbookXform();
        },
        preparedModel: require("./data/book.1.1.json"),
        xml: fs.readFileSync(`${__dirname}/data/book.1.2.xml`).toString().replace(/\r\n/g, "\n"),
        parsedModel: require("./data/book.1.3.json"),
        tests: ["render", "renderIn", "parse"],
    },
    {
        title: "book.2 - no properties",
        create() {
            return new WorkbookXform();
        },
        xml: fs.readFileSync(`${__dirname}/data/book.2.2.xml`).toString().replace(/\r\n/g, "\n"),
        parsedModel: require("./data/book.2.3.json"),
        tests: ["parse"],
    },
];

describe("WorkbookXform", () => {
    testXformHelper(expectations);

    it("makes the references of print areas absolute in linear time", () => {
        ["1", "A", "$", "$1", "A$"].forEach((unit) => {
            const printArea = `${unit.repeat(10000)}x:${unit.repeat(10000)}`;
            const model = { worksheets: [{ name: "s", pageSetup: { printArea } }], definedNames: [] };
            const start = Date.now();
            new WorkbookXform().prepare(model);
            expect(Date.now() - start, JSON.stringify(unit)).to.be.below(50);
            expect(model.definedNames[0].ranges[0].length).to.be.above(20000);
        });
    });
});
