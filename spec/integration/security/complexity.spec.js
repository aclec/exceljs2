// Regression tests for hostile input that used to cost super-linear time or memory
// (ReDoS, quadratic lookups, per-index traversals of sparse arrays).
const { baseWorkbookBuffer, tamper, zipParts } = require("../../utils/xlsx-tamper");

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
});
