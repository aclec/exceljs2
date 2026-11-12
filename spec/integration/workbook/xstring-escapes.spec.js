// ECMA-376 ST_Xstring: "_xHHHH_" in the text of a string encodes the character U+HHHH, and readers decode it. A literal
// "_xHHHH_" is written with its "_" escaped ("_x005F_xHHHH_"), so that it reads back as written.
const { PassThrough } = require("stream");
const { readPart } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

const LITERAL = "x_x0041_y _x000D_";
const RICH = { richText: [{ font: { bold: true }, text: LITERAL }, { text: "_x0042_" }] };

function fill(ws) {
    ws.getCell("A1").value = LITERAL;
    ws.getCell("A2").value = RICH;
    ws.getCell("A3").note = LITERAL;
}

async function readStream(buffer) {
    const values = [];
    for await (const worksheet of new ExcelJS.stream.xlsx.WorkbookReader(PassThrough.from([buffer]))) {
        for await (const row of worksheet) {
            values.push(row.getCell(1).value);
        }
    }
    return values;
}

describe("_xHHHH_ sequences in strings", () => {
    it("round-trip through the document writer and both readers", async () => {
        const wb = new ExcelJS.Workbook();
        fill(wb.addWorksheet("s"));
        const buffer = await wb.xlsx.writeBuffer();
        expect(await readPart(buffer, "xl/sharedStrings.xml")).to.include("x_x005F_x0041_y _x005F_x000D_");

        const loaded = new ExcelJS.Workbook();
        await loaded.xlsx.load(buffer);
        const ws = loaded.getWorksheet("s");
        expect(ws.getCell("A1").value).to.equal(LITERAL);
        expect(ws.getCell("A2").value).to.deep.equal(RICH);
        expect(ws.getCell("A3").note).to.equal(LITERAL);

        expect((await readStream(buffer)).slice(0, 2)).to.deep.equal([LITERAL, RICH]);
    });

    it("round-trip through the document writer without shared strings", async () => {
        const wb = new ExcelJS.Workbook();
        fill(wb.addWorksheet("s"));
        const loaded = new ExcelJS.Workbook();
        await loaded.xlsx.load(await wb.xlsx.writeBuffer({ useSharedStrings: false }));
        const ws = loaded.getWorksheet("s");
        expect(ws.getCell("A1").value).to.equal(LITERAL);
        expect(ws.getCell("A2").value).to.deep.equal(RICH);
    });

    it("round-trip through the streaming writer", async () => {
        const stream = new PassThrough();
        const chunks = [];
        stream.on("data", (chunk) => chunks.push(chunk));
        const writer = new ExcelJS.stream.xlsx.WorkbookWriter({ stream, useSharedStrings: true });
        const ws = writer.addWorksheet("s");
        ws.getCell("A1").value = LITERAL;
        ws.getCell("A2").value = RICH;
        ws.commit();
        await writer.commit();
        const loaded = new ExcelJS.Workbook();
        await loaded.xlsx.load(Buffer.concat(chunks));
        expect(loaded.getWorksheet("s").getCell("A1").value).to.equal(LITERAL);
        expect(loaded.getWorksheet("s").getCell("A2").value).to.deep.equal(RICH);
    });
});
