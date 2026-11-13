const { cloneDeep } = require("../../../../utils/under-dash");
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
    {
        // CT_TableColumn has dataDxfId (the style of the data cells), not dxfId
        title: "style",
        create() {
            return new TableColumnXform();
        },
        preparedModel: { id: 1, name: "Foo", dxfId: 0 },
        xml: '<tableColumn id="1" name="Foo" dataDxfId="0" />',
        parsedModel: { name: "Foo", dxfId: "0" },
        tests: ["render", "renderIn", "parse"],
    },
    {
        // exceljs 4.x and 5.0 wrote dxfId: still read
        title: "legacy dxfId",
        create() {
            return new TableColumnXform();
        },
        xml: '<tableColumn id="1" name="Foo" dxfId="3" />',
        parsedModel: { name: "Foo", dxfId: "3" },
        tests: ["parse"],
    },
    {
        title: "custom totals row formula",
        create() {
            return new TableColumnXform();
        },
        preparedModel: { id: 2, name: "B", totalsRowFunction: "custom", totalsRowFormula: "ROW()" },
        xml: '<tableColumn id="2" name="B" totalsRowFunction="custom"><totalsRowFormula>ROW()</totalsRowFormula></tableColumn>',
        parsedModel: { name: "B", totalsRowFunction: "custom", totalsRowFormula: "ROW()" },
        tests: ["render", "renderIn", "parse"],
    },
    {
        title: "calculated column and array totals row formulas",
        create() {
            return new TableColumnXform();
        },
        preparedModel: {
            id: 3,
            name: "C",
            totalsRowFunction: "custom",
            totalsRowFormula: "SUM(T[A]*T[B])",
            totalsRowFormulaArray: true,
            calculatedColumnFormula: 'T[[#This Row],[A]]&"<"',
            dxfId: 1,
        },
        xml:
            '<tableColumn id="3" name="C" totalsRowFunction="custom" dataDxfId="1">' +
            "<calculatedColumnFormula>T[[#This Row],[A]]&amp;&quot;&lt;&quot;</calculatedColumnFormula>" +
            '<totalsRowFormula array="1">SUM(T[A]*T[B])</totalsRowFormula>' +
            "</tableColumn>",
        parsedModel: {
            name: "C",
            totalsRowFunction: "custom",
            totalsRowFormula: "SUM(T[A]*T[B])",
            totalsRowFormulaArray: true,
            calculatedColumnFormula: 'T[[#This Row],[A]]&"<"',
            dxfId: "1",
        },
        tests: ["render", "renderIn", "parse"],
    },
];

const parseSax = verquire("utils/parse-sax");
const XmlStream = verquire("utils/xml-stream");

function render(model) {
    const xmlStream = new XmlStream();
    new TableColumnXform().render(xmlStream, model);
    return xmlStream.xml;
}

// the parsed model, without its undefined properties (as the parse test of test-xform-helper)
async function parse(chunks) {
    return cloneDeep(await new TableColumnXform().parse(parseSax(chunks)), false);
}

describe("TableColumnXform", () => {
    testXformHelper(expectations);

    it("throws a clear Error for a column without name (CT_TableColumn requires it)", async () => {
        const error = await new TableColumnXform().parse(parseSax(['<tableColumn id="1"/>'])).then(
            () => undefined,
            (e) => e
        );
        expect(error).to.be.an.instanceOf(Error);
        expect(error.message).to.equal("Invalid table: <tableColumn> without name");
    });

    it("writes the formulas in the CT_TableColumn order (raw string: chai-xml ignores sibling order)", () => {
        expect(
            render({
                id: 1,
                name: "A",
                totalsRowFunction: "custom",
                totalsRowFormula: "1+1",
                calculatedColumnFormula: "2+2",
                calculatedColumnFormulaArray: true,
            })
        ).to.equal(
            '<tableColumn id="1" name="A" totalsRowFunction="custom">' +
                '<calculatedColumnFormula array="1">2+2</calculatedColumnFormula>' +
                "<totalsRowFormula>1+1</totalsRowFormula>" +
                "</tableColumn>"
        );
    });

    it("writes the totals row formula only for the custom function (the others are SUBTOTALs of the cells)", () => {
        expect(render({ id: 1, name: "A", totalsRowFunction: "sum", totalsRowFormula: "SUBTOTAL(109,T[A])" })).to.equal(
            '<tableColumn id="1" name="A" totalsRowFunction="sum"/>'
        );
    });

    it("writes the style as dataDxfId, never dxfId", () => {
        const xml = render({ id: 1, name: "A", dxfId: 2 });
        expect(xml).to.equal('<tableColumn id="1" name="A" dataDxfId="2"/>');
        expect(xml).not.to.include(" dxfId=");
    });

    it("reads dataDxfId before a legacy dxfId", async () => {
        expect(await parse(['<tableColumn id="1" name="A" dxfId="1" dataDxfId="2"/>'])).to.deep.equal({ name: "A", dxfId: "2" });
    });

    it("reads a formula split across text chunks, and ignores the formulas of nested elements", async () => {
        const model = await parse([
            '<tableColumn id="1" name="A" totalsRowFunction="custom"><totalsRowFormula array="true">SUM(',
            "T[A])</totalsRowFormula>",
            "<extLst><ext><totalsRowFormula>NOT(1)</totalsRowFormula><tableColumn name='x'/></ext></extLst>",
            "</tableColumn>",
        ]);
        expect(model).to.deep.equal({
            name: "A",
            totalsRowFunction: "custom",
            totalsRowFormula: "SUM(T[A])",
            totalsRowFormulaArray: true,
        });
    });

    it('reads array="0" and an empty formula as absent', async () => {
        const model = await parse([
            '<tableColumn id="1" name="A"><calculatedColumnFormula array="0">A</calculatedColumnFormula><totalsRowFormula/></tableColumn>',
        ]);
        expect(model).to.deep.equal({ name: "A", calculatedColumnFormula: "A" });
    });

    it("rejects a formula longer than the Excel limit (8,192 characters) as an invalid table part", async () => {
        const InvalidTablePartError = verquire("xlsx/xform/table/invalid-table-part-error");
        const ok = await parse([`<tableColumn id="1" name="A"><totalsRowFormula>${"1".repeat(8192)}</totalsRowFormula></tableColumn>`]);
        expect(ok.totalsRowFormula).to.have.length(8192);
        const error = await parse([
            '<tableColumn id="1" name="A"><calculatedColumnFormula>',
            "1".repeat(8000),
            "1".repeat(193),
            "</calculatedColumnFormula></tableColumn>",
        ]).then(
            () => undefined,
            (e) => e
        );
        expect(error).to.be.an.instanceOf(InvalidTablePartError);
        expect(error.message).to.equal('Invalid table: the calculatedColumnFormula of column "A" is longer than 8192 characters');
    });
});
