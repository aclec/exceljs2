// A table read from a file is edited with the table API exactly as a table created by addTable (it used to throw
// "Table "T" was read from a file: change the cells of its worksheet").
const { readPart } = require("../../utils/xlsx-tamper");

const Excel = verquire("exceljs");

function addTable(ws) {
    return ws.addTable({
        name: "T",
        ref: "B2",
        totalsRow: true,
        columns: [
            { name: "Item", totalsRowLabel: "Sum" },
            { name: "Qty", totalsRowFunction: "sum" },
        ],
        rows: [
            ["a", 1],
            ["b", 2],
        ],
    });
}

// the values of the cells of a range, by row
function values(ws, top, left, bottom, right) {
    const rows = [];
    for (let row = top; row <= bottom; row++) {
        const cells = [];
        for (let col = left; col <= right; col++) {
            cells.push(ws.getCell(row, col).value);
        }
        rows.push(cells);
    }
    return rows;
}

async function load(buffer) {
    const wb = new Excel.Workbook();
    await wb.xlsx.load(buffer);
    return wb;
}

describe("Workbook", () => {
    describe("Tables of a loaded worksheet", () => {
        it("are edited, written and read back as created tables", async () => {
            // the same edits on a created table and on that table read from a file give the same cells and table XML
            const created = new Excel.Workbook();
            const createdTable = addTable(created.addWorksheet("s"));
            const loaded = await load(await created.xlsx.writeBuffer());
            const loadedTable = loaded.getWorksheet("s").getTable("T");

            for (const table of [createdTable, loadedTable]) {
                table.addRow(["c", 3]);
                table.theme = "TableStyleLight9";
                table.showRowStripes = true;
                table.commit();
            }

            const results = [];
            for (const wb of [created, loaded]) {
                const buffer = await wb.xlsx.writeBuffer();
                const ws = (await load(buffer)).getWorksheet("s");
                results.push({ xml: await readPart(buffer, "xl/tables/table1.xml"), cells: values(ws, 2, 2, 6, 3) });
            }

            expect(results[1].cells).to.deep.equal([
                ["Item", "Qty"],
                ["a", 1],
                ["b", 2],
                ["c", 3],
                ["Sum", { formula: "SUBTOTAL(109,T[Qty])" }],
            ]);
            expect(results[1].cells).to.deep.equal(results[0].cells);
            expect(results[1].xml).to.equal(results[0].xml);
            expect(results[1].xml).to.match(/<table [^>]* ref="B2:C6"/);
            expect(results[1].xml).to.match(/<tableStyleInfo name="TableStyleLight9"[^>]* showRowStripes="1"/);
        });

        it("reads its data rows from the cells changed after the load", async () => {
            const source = new Excel.Workbook();
            addTable(source.addWorksheet("s"));
            const wb = await load(await source.xlsx.writeBuffer());
            const ws = wb.getWorksheet("s");
            ws.getCell("C4").value = 20;

            const table = ws.getTable("T");
            table.removeRows(0);
            table.commit();

            const back = (await load(await wb.xlsx.writeBuffer())).getWorksheet("s");
            expect(values(back, 2, 2, 5, 3)).to.deep.equal([
                ["Item", "Qty"],
                ["b", 20],
                ["Sum", { formula: "SUBTOTAL(109,T[Qty])" }],
                [null, null],
            ]);
            expect(back.getTable("T").model.tableRef).to.equal("B2:C4");
        });
    });
});
