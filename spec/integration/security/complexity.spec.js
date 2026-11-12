// Regression tests for hostile input that used to cost super-linear time or memory
// (ReDoS, quadratic lookups, per-index traversals of sparse arrays).
const { Writable } = require("stream");
const { baseWorkbookBuffer, tamper, readPart, zipParts } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

const SHEET = "xl/worksheets/sheet1.xml";

async function load(buffer) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    return wb;
}

async function timed(fn) {
    const start = Date.now();
    const result = await fn();
    return { result, ms: Date.now() - start };
}

function timedSync(fn) {
    const start = Date.now();
    const result = fn();
    return { result, ms: Date.now() - start };
}

describe("hostile input: algorithmic complexity", function () {
    this.timeout(20000);

    describe("conditional formatting thresholds", () => {
        const LONG = `${"1".repeat(100000)}x`;

        it("parses a long non-numeric <cfvo val> in linear time", async () => {
            const buffer = await tamper(
                await baseWorkbookBuffer((wb, ws) => {
                    ws.addConditionalFormatting({
                        ref: "A1:A2",
                        rules: [
                            {
                                type: "colorScale",
                                priority: 1,
                                cfvo: [{ type: "num", value: 5 }, { type: "max" }],
                                color: [{ argb: "FFFF0000" }, { argb: "FF00FF00" }],
                            },
                        ],
                    });
                }),
                { [SHEET]: (xml) => xml.replace('<cfvo type="num" val="5"', `<cfvo type="num" val="${LONG}"`) }
            );
            const { result: wb, ms } = await timed(() => load(buffer));
            expect(wb.getWorksheet(1).conditionalFormattings[0].rules[0].cfvo[0].value).to.equal(LONG);
            expect(ms).to.be.below(1000);
        });

        it("parses a long non-numeric x14 <xm:f> threshold in linear time", async () => {
            const buffer = await tamper(
                await baseWorkbookBuffer((wb, ws) => {
                    ws.addConditionalFormatting({
                        ref: "A1:A2",
                        rules: [{ type: "dataBar", cfvo: [{ type: "num", value: 5 }, { type: "max" }], color: { argb: "FF00FF00" } }],
                    });
                }),
                {
                    [SHEET]: (xml) => {
                        expect(xml).to.include("<xm:f>5</xm:f>");
                        return xml.replace("<xm:f>5</xm:f>", `<xm:f>${LONG}</xm:f>`);
                    },
                }
            );
            const { ms } = await timed(() => load(buffer));
            expect(ms).to.be.below(1000);
        });
    });

    it("matches worksheets to many <sheet> elements in linear time", async () => {
        const parts = { "xl/workbook.xml": `<workbook><sheets>${'<sheet name=""/>'.repeat(400000)}</sheets></workbook>` };
        for (let i = 1; i <= 2000; i++) {
            parts[`xl/worksheets/sheet${i}.xml`] = "<worksheet/>";
        }
        const { ms } = await timed(() => load(zipParts(parts)));
        expect(ms).to.be.below(1500);
    });

    it("joins many Print_Area names of a worksheet in linear time", async () => {
        const name = '<definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$B$2</definedName>';
        const buffer = await tamper(await baseWorkbookBuffer(), {
            "xl/workbook.xml": (xml) => xml.replace("</sheets>", `</sheets><definedNames>${name.repeat(100000)}</definedNames>`),
        });
        const { result: wb, ms } = await timed(() => load(buffer));
        const { printArea } = wb.getWorksheet(1).pageSetup;
        expect(printArea.startsWith("A1:B2&&A1:B2&&")).to.equal(true);
        expect(printArea.length).to.equal(100000 * "A1:B2&&".length - 2);
        expect(ms).to.be.below(1500);
    });

    it("writes names of far single cells in time proportional to the cells, not to their indexes", async () => {
        let names = "";
        for (let i = 0; i < 100; i++) {
            names += `<definedName name="n${i}">Sheet1!$XFD$1048576</definedName>`;
        }
        const buffer = await tamper(await baseWorkbookBuffer(), {
            "xl/workbook.xml": (xml) => xml.replace("</sheets>", `</sheets><definedNames>${names}</definedNames>`),
        });
        const wb = await load(buffer);
        const { result: written, ms } = await timed(() => wb.xlsx.writeBuffer());
        expect(ms).to.be.below(1000);
        const names2 = (await load(written)).definedNames.model;
        expect(names2.length).to.equal(100);
        expect(names2[99]).to.deep.equal({ name: "n99", ranges: ["Sheet1!$XFD$1048576"] });
    });

    it("splices names of far single cells in time proportional to the cells", () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet("Sheet1");
        ws.getColumn(1).width = 10;
        for (let i = 0; i < 100; i++) {
            wb.definedNames.add("Sheet1!$XFD$1048576", `n${i}`);
            wb.definedNames.add("Sheet1!$B$2", `n${i}`);
        }
        const { ms } = timedSync(() => {
            ws.spliceRows(1, 1);
            ws.spliceColumns(1, 1);
        });
        expect(ms).to.be.below(500);
        expect(wb.definedNames.model[99]).to.deep.equal({ name: "n99", ranges: ["Sheet1!$A$1", "Sheet1!$XFC$1048575"] });
    });

    it("keeps a <col> range spanning every column as a range instead of one object per column", async () => {
        const SHEETS = 100;
        const patches = {};
        for (let i = 1; i <= SHEETS; i++) {
            patches[`xl/worksheets/sheet${i}.xml`] = (xml) =>
                xml
                    .replace("<sheetData>", '<cols><col min="1" max="16384" width="20" customWidth="1"/></cols><sheetData>')
                    .replace(/<c r="A1"/, '<c r="XFD1"');
        }
        const buffer = await tamper(
            await baseWorkbookBuffer((wb) => {
                for (let i = 2; i <= SHEETS; i++) {
                    wb.addWorksheet(`Sheet${i}`).getCell("A1").value = i;
                }
            }),
            patches
        );
        const heapBefore = process.memoryUsage().heapUsed;
        const wb = await load(buffer);
        const written = await wb.xlsx.writeBuffer();
        // one Column object per column of every worksheet took 270 MB
        expect(process.memoryUsage().heapUsed - heapBefore).to.be.below(100 * 1024 * 1024);

        const ws = (await load(written)).getWorksheet(SHEETS);
        expect(ws.getCell("XFD1").value).to.equal(SHEETS);
        expect(ws.getColumn(100).width).to.equal(20);
        expect(ws.model.cols.map(({ min, max, width }) => [min, max, width])).to.deep.equal([[1, 16384, 20]]);
    });

    describe("far rows", () => {
        // number of rows a worksheet holds, without traversing every index below the last one
        const rowsHeld = (ws) => Object.keys(ws._rows).filter((key) => ws._rows[key]).length;

        it("traverses worksheets whose only row is the last one in time proportional to the rows", async () => {
            const SHEETS = 100;
            const patches = {};
            for (let i = 1; i <= SHEETS; i++) {
                patches[`xl/worksheets/sheet${i}.xml`] = (xml) =>
                    xml.replace(/<sheetData>.*<\/sheetData>/, '<sheetData><row r="1048576"><c r="A1048576"><v>1</v></c></row></sheetData>');
            }
            const buffer = await tamper(
                await baseWorkbookBuffer((wb) => {
                    for (let i = 2; i <= SHEETS; i++) {
                        wb.addWorksheet(`Sheet${i}`).getCell("A1").value = i;
                    }
                }),
                patches
            );
            const wb = await load(buffer);
            // each of these visited the 1,048,576 indexes of the rows of every worksheet: 14 ms per worksheet
            const { result, ms } = await timed(async () => {
                const written = await wb.xlsx.writeBuffer();
                const counts = wb.worksheets.map((ws) => {
                    let eachRow = 0;
                    ws.eachRow(() => eachRow++);
                    return [ws.actualRowCount, ws.rowCount, ws.dimensions.model.bottom, ws.getSheetValues().length, eachRow];
                });
                return { written, counts };
            });
            expect(ms).to.be.below(1000);
            expect(result.counts[SHEETS - 1]).to.deep.equal([1, 1048576, 1048576, 1048577, 1]);
            const ws = (await load(result.written)).getWorksheet(SHEETS);
            expect(ws.getCell("A1048576").value).to.equal(1);
            expect(rowsHeld(ws)).to.equal(1);
        });

        it("splices the rows and columns of a worksheet with a far row without visiting the rows between", () => {
            const FAR = 200000;
            const ws = new ExcelJS.Workbook().addWorksheet("s");
            ws.getCell("A1").value = "first";
            ws.getCell(`A${FAR}`).value = "far";
            const { ms } = timedSync(() => {
                ws.spliceRows(1, 1);
                ws.insertRow(1, ["inserted"]);
                ws.duplicateRow(1, 2, true);
                ws.spliceColumns(1, 0, [undefined, "y"], ["z"]);
                ws.spliceColumns(2, 1);
            });
            expect(ms).to.be.below(500);
            // inserting a column used to create every row up to the last one
            expect(rowsHeld(ws)).to.equal(4);
            expect(ws.getSheetValues().length).to.equal(FAR + 3);
            expect([1, 2, 3, FAR + 2].map((r) => [ws.getCell(r, 1).value, ws.getCell(r, 2).value])).to.deep.equal([
                [null, "inserted"],
                ["y", "inserted"],
                [null, "inserted"],
                [null, "far"],
            ]);
        });
    });

    describe("cells in the last column", () => {
        // number of cells a row holds, without traversing every index below the last one
        const cellsHeld = (row) => Object.keys(row._cells).filter((key) => row._cells[key]).length;

        it("traverses rows whose only cell is in column XFD in time proportional to the cells", async () => {
            const ROWS = 3000;
            let data = "";
            for (let i = 1; i <= ROWS; i++) {
                data += `<row r="${i}"><c r="XFD${i}"><v>${i}</v></c></row>`;
            }
            const buffer = await tamper(await baseWorkbookBuffer(), {
                [SHEET]: (xml) => xml.replace(/<sheetData>.*<\/sheetData>/, `<sheetData>${data}</sheetData>`),
            });
            // each of these visited the 16,384 indexes of the cells of every row: 0.2 ms per row
            const { result, ms } = await timed(async () => {
                const wb = await load(buffer);
                const ws = wb.getWorksheet(1);
                let cells = 0;
                ws.eachRow((row) => row.eachCell(() => cells++));
                const counts = [cells, ws.actualColumnCount, ws.dimensions.model.left, ws.getRow(ROWS).values.length];
                ws.spliceColumns(1, 1);
                return { counts, written: await wb.xlsx.writeBuffer() };
            });
            expect(ms).to.be.below(1500);
            expect(result.counts).to.deep.equal([ROWS, 1, 16384, 16385]);
            const ws = (await load(result.written)).getWorksheet(1);
            expect(ws.getCell(`XFC${ROWS}`).value).to.equal(ROWS);
            expect(cellsHeld(ws.getRow(ROWS))).to.equal(1);
        });

        it("moves rows with a cell in column XFC without creating their empty cells", () => {
            const ws = new ExcelJS.Workbook().addWorksheet("s");
            for (let i = 1; i <= 10; i++) {
                ws.getCell(`XFC${i}`).value = i;
                ws.getCell(`XFC${i}`).font = { name: `f${i}` };
            }
            ws.insertRow(1, ["inserted"]);
            ws.spliceRows(2, 1);
            ws.duplicateRow(2, 2, true);
            ws.spliceColumns(2, 0, ["x"]);
            ws.spliceColumns(3, 1);
            // moving a row created each of its 16,383 cells in both rows
            ws.eachRow((row) => expect(cellsHeld(row), `row ${row.number}`).to.be.at.most(3));
            expect([2, 3, 4, 12].map((r) => ws.getCell(r, 16383).value)).to.deep.equal([2, 2, 2, 10]);
            expect(ws.getRow(1).values).to.deep.equal([undefined, "inserted", "x"]);
            expect(ws.getCell("XFC12").font).to.deep.equal({ name: "f10" });
            expect(ws.getCell("XFC3").font).to.deep.equal({ name: "f2" });
            expect(ws.getRow(12).cellCount).to.equal(16384);
        });
    });

    describe("streaming writer", () => {
        it("commits a far row in time proportional to the rows", async () => {
            const FAR = 10000;
            const chunks = [];
            const stream = new Writable({
                write(chunk, encoding, callback) {
                    chunks.push(chunk);
                    callback();
                },
            });
            const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ stream });
            const ws = wb.addWorksheet("s");
            const { ms } = await timed(async () => {
                ws.getCell("A1").value = "first";
                await ws.getRow(1).commit();
                ws.getCell(`A${FAR}`).value = "far";
                ws.getCell(`XFD${FAR}`).value = "last";
                ws.getCell(`A${FAR + 1}`).value = "next";
                let rows = 0;
                ws.eachRow(() => rows++);
                expect(rows).to.equal(2);
                // committing shifted the rows one index at a time out of an array with a hole of FAR indexes
                await ws.getRow(FAR).commit();
                expect(ws.findRow(FAR + 1).getCell(1).value).to.equal("next");
                ws.getCell(`A${FAR + 2}`).value = "added";
                await ws.commit();
                await wb.commit();
            });
            expect(ms).to.be.below(500);
            const read = (await load(Buffer.concat(chunks))).getWorksheet(1);
            expect([1, FAR, FAR + 1, FAR + 2].map((r) => read.getCell(r, 1).value)).to.deep.equal(["first", "far", "next", "added"]);
            expect(read.getCell(`XFD${FAR}`).value).to.equal("last");
        });
    });

    describe("image anchors", () => {
        const PNG = Buffer.from(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==",
            "base64"
        );

        async function anchoredAt(col, row) {
            const buffer = await tamper(
                await baseWorkbookBuffer((wb, ws) => {
                    ws.addImage(wb.addImage({ buffer: PNG, extension: "png" }), "B2:C3");
                }),
                {
                    "xl/drawings/drawing1.xml": (xml) =>
                        xml
                            .replace("<xdr:col>1</xdr:col>", `<xdr:col>${col}</xdr:col>`)
                            .replace("<xdr:row>1</xdr:row>", `<xdr:row>${row}</xdr:row>`),
                }
            );
            const ws = (await load(buffer)).getWorksheet(1);
            return { ws, tl: ws.getImages()[0].range.tl };
        }

        it("clamps an anchor beyond the grid to the last cell", async () => {
            const { tl } = await anchoredAt(200000, 2000000);
            expect([tl.nativeCol, tl.nativeRow]).to.deep.equal([16383, 1048575]);
            expect([Math.floor(tl.col), Math.floor(tl.row)]).to.deep.equal([16383, 1048575]);
        });

        it("reads the position of an anchor without creating its row and column", async () => {
            const { ws, tl } = await anchoredAt(5000, 70000);
            const { rowCount } = ws;
            expect([tl.col, tl.row]).to.deep.equal([5000, 70000]);
            expect(ws.findRow(70001)).to.equal(undefined);
            expect(ws.rowCount).to.equal(rowCount);
            expect(ws.columns.length).to.be.below(5001);
        });

        it("still uses the width and height of the anchor cell", async () => {
            const { ws, tl } = await anchoredAt(1, 1);
            ws.getColumn(2).width = 20;
            ws.getRow(2).height = 30;
            tl.nativeColOff = 100000;
            tl.nativeRowOff = 150000;
            expect(tl.col).to.equal(1.5);
            expect(tl.row).to.equal(1.5);
        });
    });

    it("writes a long rich text shared by many cells in linear time", async () => {
        const CELLS = 4000;
        let data = "";
        for (let i = 1; i <= CELLS; i++) {
            data += `<row r="${i}"><c r="A${i}" t="s"><v>0</v></c></row>`;
        }
        const buffer = await tamper(await baseWorkbookBuffer(), {
            "xl/sharedStrings.xml": (xml) => xml.replace(/<si>.*<\/si>/, `<si>${"<r><t>a</t></r>".repeat(4000)}</si>`),
            [SHEET]: (xml) => xml.replace(/<sheetData>.*<\/sheetData>/, `<sheetData>${data}</sheetData>`),
        });
        const wb = await load(buffer);
        const { result: written, ms } = await timed(() => wb.xlsx.writeBuffer());
        expect(ms).to.be.below(1000);
        const ws = (await load(written)).getWorksheet(1);
        expect(ws.getCell(`A${CELLS}`).value.richText.length).to.equal(4000);
    });

    it("still writes equal rich texts of different cells once", async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet("s");
        ws.getCell("A1").value = { richText: [{ text: "a", font: { bold: true } }, { text: "b" }] };
        ws.getCell("A2").value = { richText: [{ text: "a", font: { bold: true } }, { text: "b" }] };
        ws.getCell("A3").value = ws.getCell("A1").value;
        ws.getCell("A4").value = { richText: [{ text: "a" }, { text: "b" }] };
        const sst = await readPart(await wb.xlsx.writeBuffer(), "xl/sharedStrings.xml");
        expect(sst).to.include('count="4" uniqueCount="2"');
    });
});
