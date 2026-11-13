const testXformHelper = require("../test-xform-helper");

const TableColumnXform = verquire("xlsx/xform/table/table-column-xform");

const expectations = [
    {
        title: "label",
        create() {
            return new TableColumnXform();
        },
        preparedModel: { id: 1, name: "Foo", totalsRowLabel: "Bar" },
        xml: '<tableColumn id="1" name="Foo" totalsRowLabel="Bar" />',
        parsedModel: { name: "Foo", totalsRowLabel: "Bar" },
        tests: ["render", "renderIn", "parse"],
    },
    {
        title: "function",
        create() {
            return new TableColumnXform();
        },
        preparedModel: { id: 1, name: "Foo", totalsRowFunction: "sum" },
        xml: '<tableColumn id="1" name="Foo" totalsRowFunction="sum" />',
        parsedModel: { name: "Foo", totalsRowFunction: "sum" },
        tests: ["render", "renderIn", "parse"],
    },
    {
        // not one of ST_TotalsRowFunction: absent (the default, none)
        title: "invalid function",
        create() {
            return new TableColumnXform();
        },
        xml: '<tableColumn id="1" name="Foo" totalsRowFunction="Baz" />',
        parsedModel: { name: "Foo" },
        tests: ["parse"],
    },
];

describe("TableColumnXform", () => {
    testXformHelper(expectations);

    it("throws a clear Error for a column without name (CT_TableColumn requires it)", async () => {
        const parseSax = verquire("utils/parse-sax");
        const error = await new TableColumnXform().parse(parseSax(['<tableColumn id="1"/>'])).then(
            () => undefined,
            (e) => e
        );
        expect(error).to.be.an.instanceOf(Error);
        expect(error.message).to.equal("Invalid table: <tableColumn> without name");
    });
});
