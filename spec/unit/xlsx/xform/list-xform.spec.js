const testXformHelper = require("./test-xform-helper");

const ListXform = verquire("xlsx/xform/list-xform");
const IntegerXform = verquire("xlsx/xform/simple/integer-xform");

const expectations = [
    {
        title: "Tagged",
        create() {
            return new ListXform({
                tag: "ints",
                childXform: new IntegerXform({ tag: "int", attr: "val" }),
            });
        },
        preparedModel: [1, 2, 3],
        get parsedModel() {
            return this.preparedModel;
        },
        xml: '<ints><int val="1"/><int val="2"/><int val="3"/></ints>',
        tests: ["render", "renderIn", "parse"],
    },
    {
        title: "Tagged and Counted",
        create() {
            return new ListXform({
                tag: "ints",
                count: true,
                childXform: new IntegerXform({ tag: "int", attr: "val" }),
            });
        },
        preparedModel: [1, 2, 3],
        get parsedModel() {
            return this.preparedModel;
        },
        xml: '<ints count="3"><int val="1"/><int val="2"/><int val="3"/></ints>',
        tests: ["render", "renderIn", "parse"],
    },
];

describe("ListXform", () => {
    testXformHelper(expectations);

    it("skips an item whose xform gives no model", async () => {
        const parseSax = verquire("utils/parse-sax");
        const MergeCellXform = verquire("xlsx/xform/sheet/merge-cell-xform");
        const list = new ListXform({ tag: "mergeCells", childXform: new MergeCellXform() });
        const model = await list.parse(parseSax(['<mergeCells><mergeCell ref="A1:B2"/><mergeCell/><mergeCell ref="C3:D4"/></mergeCells>']));
        expect(model).to.deep.equal(["A1:B2", "C3:D4"]);
    });
});
