const fs = require("fs");

const testXformHelper = require("../test-xform-helper");

const TableXform = verquire("xlsx/xform/table/table-xform");

const expectations = [
    {
        title: "showing filter",
        create() {
            return new TableXform();
        },
        initialModel: null,
        preparedModel: require("./data/table.1.1"),
        xml: fs.readFileSync(`${__dirname}/data/table.1.2.xml`).toString(),
        parsedModel: require("./data/table.1.3"),
        tests: ["render", "renderIn", "parse"],
    },
    {
        // CT_Table: tableStyleInfo is optional (no table style)
        title: "without table style",
        create() {
            return new TableXform();
        },
        preparedModel: {
            id: 2,
            name: "T",
            tableRef: "A1:A2",
            totalsRow: false,
            headerRow: false,
            columns: [{ id: 1, name: "A" }],
            style: null,
        },
        xml:
            `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"` +
            ` xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr xr3"` +
            ` xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"` +
            ` xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"` +
            ` id="2" name="T" displayName="T" ref="A1:A2" totalsRowShown="1" headerRowCount="0">` +
            `<tableColumns count="1"><tableColumn id="1" name="A"/></tableColumns>` +
            `</table>`,
        parsedModel: {
            name: "T",
            displayName: "T",
            tableRef: "A1:A2",
            totalsRow: false,
            headerRow: false,
            columns: [{ name: "A" }],
            style: null,
        },
        tests: ["render", "parse"],
    },
    {
        // totalsRowCount and headerRowCount are counts (xsd:unsignedInt), not booleans
        title: "row counts",
        create() {
            return new TableXform();
        },
        xml:
            `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" ref="A1:A4"` +
            ` totalsRowCount="2" headerRowCount="00">` +
            `<tableColumns count="1"><tableColumn id="1" name="A"/></tableColumns>` +
            `</table>`,
        parsedModel: {
            name: "T",
            displayName: "T",
            tableRef: "A1:A4",
            totalsRow: true,
            headerRow: false,
            columns: [{ name: "A" }],
            style: null,
        },
        tests: ["parse"],
    },
];

describe("TableXform", () => {
    testXformHelper(expectations);

    const parseSax = verquire("utils/parse-sax");
    const parseError = (xml) =>
        new TableXform().parse(parseSax([xml])).then(
            () => undefined,
            (e) => e
        );

    it("throws a clear Error for a table without name (CT_Table requires it)", async () => {
        const error = await parseError(
            '<table id="1" ref="A1:A2"><tableColumns count="1"><tableColumn id="1" name="A"/></tableColumns></table>'
        );
        expect(error).to.be.an.instanceOf(Error);
        expect(error.message).to.equal("Invalid table part: <table> without name");
    });

    it("throws a clear Error for a table without <tableColumns> (CT_Table requires it)", async () => {
        const error = await parseError('<table id="1" name="T" ref="A1:A2"><tableStyleInfo name="TableStyleLight1"/></table>');
        expect(error).to.be.an.instanceOf(Error);
        expect(error.message).to.equal('Invalid table "T": no <tableColumns>');
    });

    it("keeps the name of a custom table style", async () => {
        const model = await new TableXform().parse(
            parseSax([
                '<table id="1" name="T" ref="A1:A2"><tableColumns count="1"><tableColumn id="1" name="A"/></tableColumns>' +
                    '<tableStyleInfo name="My Style" showRowStripes="1"/></table>',
            ])
        );
        expect(model.style.theme).to.equal("My Style");
    });
});
