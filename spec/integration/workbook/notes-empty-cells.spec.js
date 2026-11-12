// Notes (comments) on cells without a value: sheetData has no <c> element for them.
const { baseWorkbookBuffer, readPart, tamper } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

async function load(buffer) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    return wb.getWorksheet("Sheet1");
}

describe("Notes on cells without a value", () => {
    it("are read back from a cell that has no <c> element", async () => {
        const buffer = await baseWorkbookBuffer((wb, ws) => {
            ws.getCell("A1").note = "on a value";
            ws.getCell("C1").note = "same row";
            ws.getCell("B3").note = "own row";
        });
        const xml = await readPart(buffer, "xl/worksheets/sheet1.xml");
        expect(xml).not.to.match(/<c r="(C1|B3)"/);

        const ws = await load(buffer);
        expect(ws.getCell("A1").note).to.equal("on a value");
        expect(ws.getCell("C1").note).to.equal("same row");
        expect(ws.getCell("B3").note).to.equal("own row");
        expect(ws.getCell("B3").value).to.equal(null);
    });

    it("are read back when sheetData has no row for them", async () => {
        const buffer = await tamper(
            await baseWorkbookBuffer((wb, ws) => {
                ws.getCell("D7").note = "no row";
            }),
            { "xl/worksheets/sheet1.xml": (xml) => xml.replace(/<row r="7"[^>]*\/>/, "") }
        );
        expect(await readPart(buffer, "xl/worksheets/sheet1.xml")).not.to.match(/<row r="7"/);
        const ws = await load(buffer);
        expect(ws.getCell("D7").note).to.equal("no row");
        // and survive another round trip
        const again = await load(await ws.workbook.xlsx.writeBuffer());
        expect(again.getCell("D7").note).to.equal("no row");
    });

    it("ignores notes whose reference is not a cell of the grid", async () => {
        const buffer = await tamper(
            await baseWorkbookBuffer((wb, ws) => {
                ws.getCell("B2").note = "b2";
                ws.getCell("C2").note = "c2";
            }),
            {
                "xl/comments1.xml": (xml) => xml.replace('ref="B2"', 'ref="B9999999"').replace('ref="C2"', 'ref="__proto__"'),
            }
        );
        const ws = await load(buffer);
        expect(ws.getCell("A1").value).to.equal("hello");
        expect(ws.rowCount).to.equal(2);
    });
});
