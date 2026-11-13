// The <tableColumn> of a table part (CT_TableColumn): the style of the data cells is dataDxfId (exceljs 4.x / 5.0 wrote a
// dxfId attribute that the schema does not have), and the formulas are children (calculatedColumnFormula,
// totalsRowFormula), read and written back.
const { baseWorkbookBuffer, tamper, readPart } = require("../../utils/xlsx-tamper");

const Excel = verquire("exceljs");

const TABLE_PART = "xl/tables/table1.xml";

function addTable(ws) {
    return ws.addTable({
        name: "T",
        ref: "B2",
        totalsRow: true,
        columns: [
            { name: "Item" },
            { name: "Qty", totalsRowFunction: "sum", style: { numFmt: "0.00" } },
            { name: "Rows", totalsRowFunction: "custom", totalsRowFormula: "ROWS(T[Qty])", totalsRowResult: 2 },
        ],
        rows: [
            ["a", 1, 1],
            ["b", 2, 2],
        ],
    });
}

async function load(buffer) {
    const wb = new Excel.Workbook();
    await wb.xlsx.load(buffer);
    return wb;
}

// the <tableColumn> elements of a table part
function tableColumns(xml) {
    const start = xml.indexOf("<tableColumns");
    const end = xml.indexOf("</tableColumns>");
    expect(start).to.be.above(-1);
    return xml.slice(xml.indexOf(">", start) + 1, end);
}

describe("Workbook", () => {
    describe("The columns of a table part", () => {
        it("are written with dataDxfId and the custom totals row formula, read and written again unchanged", async () => {
            const first = await baseWorkbookBuffer((wb, ws) => addTable(ws));
            const firstColumns = tableColumns(await readPart(first, TABLE_PART));
            expect(firstColumns).to.equal(
                '<tableColumn id="1" name="Item" totalsRowLabel="Total"/>' +
                    '<tableColumn id="2" name="Qty" totalsRowFunction="sum" dataDxfId="0"/>' +
                    '<tableColumn id="3" name="Rows" totalsRowFunction="custom">' +
                    "<totalsRowFormula>ROWS(T[Qty])</totalsRowFormula></tableColumn>"
            );

            const wb = await load(first);
            const table = wb.getWorksheet("Sheet1").getTable("T");
            expect(table.getColumn(1).style).to.deep.include({ numFmt: "0.00" });
            expect(table.getColumn(2).totalsRowFormula).to.equal("ROWS(T[Qty])");

            const second = await wb.xlsx.writeBuffer();
            expect(tableColumns(await readPart(second, TABLE_PART))).to.equal(firstColumns);
        });

        it("read with the legacy dxfId of exceljs 4.x / 5.0 are written with dataDxfId", async () => {
            const legacy = await tamper(await baseWorkbookBuffer((wb, ws) => addTable(ws)), {
                [TABLE_PART]: (xml) => {
                    const patched = xml.replace(' dataDxfId="0"', ' dxfId="0"');
                    expect(patched).not.to.equal(xml);
                    return patched;
                },
            });

            const wb = await load(legacy);
            expect(wb.getWorksheet("Sheet1").getTable("T").getColumn(1).style).to.deep.include({ numFmt: "0.00" });

            const xml = await readPart(await wb.xlsx.writeBuffer(), TABLE_PART);
            expect(xml).to.include('<tableColumn id="2" name="Qty" totalsRowFunction="sum" dataDxfId="0"/>');
            expect(xml).not.to.include(" dxfId=");
        });

        it("keep their calculated column formula and the array flag of their formulas", async () => {
            const source = await tamper(await baseWorkbookBuffer((wb, ws) => addTable(ws)), {
                [TABLE_PART]: (xml) => {
                    const patched = xml
                        .replace(
                            '<tableColumn id="1" name="Item" totalsRowLabel="Total"/>',
                            '<tableColumn id="1" name="Item" totalsRowLabel="Total">' +
                                "<calculatedColumnFormula>T[[#This Row],[Qty]]&amp;&quot;x&quot;</calculatedColumnFormula>" +
                                "</tableColumn>"
                        )
                        .replace("<totalsRowFormula>", '<totalsRowFormula array="1">');
                    expect(patched).not.to.equal(xml);
                    return patched;
                },
            });

            const wb = await load(source);
            const xml = await readPart(await wb.xlsx.writeBuffer(), TABLE_PART);
            expect(tableColumns(xml)).to.equal(
                '<tableColumn id="1" name="Item" totalsRowLabel="Total">' +
                    "<calculatedColumnFormula>T[[#This Row],[Qty]]&amp;&quot;x&quot;</calculatedColumnFormula></tableColumn>" +
                    '<tableColumn id="2" name="Qty" totalsRowFunction="sum" dataDxfId="0"/>' +
                    '<tableColumn id="3" name="Rows" totalsRowFunction="custom">' +
                    '<totalsRowFormula array="1">ROWS(T[Qty])</totalsRowFormula></tableColumn>'
            );
        });

        it("lose the array flag of the totals row formula on commit (the totals cell is written with a plain formula)", async () => {
            const source = await tamper(await baseWorkbookBuffer((wb, ws) => addTable(ws)), {
                [TABLE_PART]: (xml) => xml.replace("<totalsRowFormula>", '<totalsRowFormula array="1">'),
            });
            const wb = await load(source);
            const table = wb.getWorksheet("Sheet1").getTable("T");
            table.addRow(["c", 3, 3]);
            table.commit();
            const xml = await readPart(await wb.xlsx.writeBuffer(), TABLE_PART);
            expect(xml).to.include("<totalsRowFormula>ROWS(T[Qty])</totalsRowFormula>");
        });

        it("keep the custom totals row formula of a table read from a file when it is changed", async () => {
            const wb = await load(await baseWorkbookBuffer((book, ws) => addTable(ws)));
            const ws = wb.getWorksheet("Sheet1");
            const table = ws.getTable("T");
            table.addRow(["c", 3, 3]);
            table.commit();
            // the totals row moved down one row: its custom formula is written in its new cell
            expect(ws.getCell("D6").value).to.deep.equal({ formula: "ROWS(T[Qty])" });
        });
    });

    describe("A function in the totals row of the first column", () => {
        it("is written in the totals row, without the default label", async () => {
            const wb = new Excel.Workbook();
            const ws = wb.addWorksheet("s");
            ws.addTable({
                name: "T",
                ref: "A1",
                totalsRow: true,
                columns: [{ name: "N", totalsRowFunction: "count" }, { name: "M" }],
                rows: [[1, 2]],
            });
            expect(ws.getCell("A3").value).to.deep.equal({ formula: "SUBTOTAL(103,T[N])" });
            const xml = await readPart(await wb.xlsx.writeBuffer(), TABLE_PART);
            expect(tableColumns(xml)).to.equal(
                '<tableColumn id="1" name="N" totalsRowFunction="count"/><tableColumn id="2" name="M" totalsRowFunction="none"/>'
            );
        });

        it("of a table read from a file is kept when the table is changed (Excel writes no label with a function)", async () => {
            const wb = new Excel.Workbook();
            const ws = wb.addWorksheet("s");
            ws.addTable({
                name: "T",
                ref: "A1",
                totalsRow: true,
                columns: [{ name: "N", totalsRowFunction: "count" }],
                rows: [[1]],
            });
            const loaded = await load(await wb.xlsx.writeBuffer());
            const sheet = loaded.getWorksheet("s");
            const table = sheet.getTable("T");
            table.addRow([2]);
            table.commit();
            expect(sheet.getCell("A4").value).to.deep.equal({ formula: "SUBTOTAL(103,T[N])" });
            expect(table.getColumn(0).totalsRowLabel).to.equal(undefined);
        });
    });

    describe("The style of a table column (a differential format of its data cells)", () => {
        const FONT = { name: "Arial", size: 14, italic: true };

        // an Excel-like table: the column "Amount" has a dataDxfId (red font, 0.00), its cells have fonts of their own
        async function excelTable() {
            const source = await baseWorkbookBuffer((wb, ws) => {
                ws.addTable({
                    name: "T",
                    ref: "A1",
                    totalsRow: true,
                    columns: [{ name: "Item" }, { name: "Amount", totalsRowFunction: "sum" }],
                    rows: [
                        ["a", 1],
                        ["b", 2],
                    ],
                });
                for (const address of ["A1", "B1", "B2", "B3", "B4"]) {
                    ws.getCell(address).font = { ...FONT };
                }
            });
            const buffer = await tamper(source, {
                "xl/styles.xml": (xml) => {
                    const patched = xml.replace(
                        '<dxfs count="0"/>',
                        '<dxfs count="1"><dxf><font><color rgb="FFFF0000"/></font><numFmt numFmtId="164" formatCode="0.00"/>' +
                            '<fill><patternFill><bgColor rgb="FF00FF00"/></patternFill></fill></dxf></dxfs>'
                    );
                    expect(patched).not.to.equal(xml);
                    return patched;
                },
                [TABLE_PART]: (xml) => xml.replace('name="Amount"', 'name="Amount" dataDxfId="0"'),
            });
            const wb = await load(buffer);
            return wb.getWorksheet("Sheet1");
        }

        it("is merged into the data cells of a table read from a file on commit, without changing its other cells", async () => {
            const ws = await excelTable();
            const table = ws.getTable("T");
            table.addRow(["c", 3]);
            table.commit();

            const styled = { ...FONT, color: { argb: "FFFF0000" } };
            const solid = { type: "pattern", pattern: "solid", fgColor: { argb: "FF00FF00" } };
            // data cells: the font properties of the dxf over those of the cell; a dxf pattern fill without pattern is solid
            for (const address of ["B2", "B3", "B4"]) {
                expect(ws.getCell(address).font, address).to.deep.equal(styled);
                expect(ws.getCell(address).numFmt, address).to.equal("0.00");
                expect(ws.getCell(address).fill, address).to.deep.equal(solid);
            }
            // header and totals cells keep their style (the format of the data cells is not theirs)
            expect(ws.getCell("B1").font).to.deep.equal(FONT);
            expect(ws.getCell("B1").numFmt).to.equal(undefined);
            expect(ws.getCell("B5").font).to.equal(undefined);
            expect(ws.getCell("B5").numFmt).to.equal(undefined);
            // the cells sharing a style object with a data cell (one per style id in a file read) are not changed
            expect(ws.getCell("A1").font).to.deep.equal(FONT);
        });

        it("of a created table is merged into its data cells only, not the header and totals cells", () => {
            const ws = new Excel.Workbook().addWorksheet("s");
            ws.getCell("B2").font = { name: "Arial" };
            ws.addTable({
                name: "T",
                ref: "A1",
                totalsRow: true,
                columns: [{ name: "Item" }, { name: "Amount", totalsRowFunction: "sum", style: { font: { bold: true } } }],
                rows: [["a", 1]],
            });
            expect(ws.getCell("B2").font).to.deep.equal({ name: "Arial", bold: true });
            expect(ws.getCell("B1").font).to.equal(undefined);
            expect(ws.getCell("B3").font).to.equal(undefined);
        });
    });

    describe("A calculated column formula", () => {
        async function loadedTable() {
            const source = await tamper(await baseWorkbookBuffer((wb, ws) => addTable(ws)), {
                [TABLE_PART]: (xml) =>
                    xml.replace(
                        '<tableColumn id="1" name="Item" totalsRowLabel="Total"/>',
                        '<tableColumn id="1" name="Item" totalsRowLabel="Total">' +
                            "<calculatedColumnFormula>T[[#This Row],[Qty]]</calculatedColumnFormula></tableColumn>"
                    ),
            });
            const wb = await load(source);
            return { wb, table: wb.getWorksheet("Sheet1").getTable("T") };
        }

        it("is kept when rows are added", async () => {
            const { wb, table } = await loadedTable();
            table.addRow(["c", 3, 3]);
            table.commit();
            expect(await readPart(await wb.xlsx.writeBuffer(), TABLE_PART)).to.include("<calculatedColumnFormula>");
        });

        for (const [title, edit] of [
            ["a column is renamed", (table) => (table.getColumn(1).name = "Quantity")],
            ["the table is renamed", (table) => (table.name = "U")],
            ["a column is removed", (table) => table.removeColumns(2)],
        ]) {
            it(`is dropped when ${title} (its structured references would be stale)`, async () => {
                const { wb, table } = await loadedTable();
                edit(table);
                table.commit();
                expect(await readPart(await wb.xlsx.writeBuffer(), TABLE_PART)).not.to.include("calculatedColumnFormula");
            });
        }
    });
});
