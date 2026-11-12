// The value types of the text of a hyperlink cell, written and read back.
const { Readable } = require("stream");
const { baseWorkbookBuffer, readPart } = require("../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

async function load(buffer) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    return workbook.getWorksheet("Sheet1");
}

describe("hyperlink values", () => {
    const date = new Date(Date.UTC(2023, 2, 15, 6));

    it("writes a Date text as a date (its serial number), which reads back as a Date", async () => {
        for (const date1904 of [false, true]) {
            const buffer = await baseWorkbookBuffer((wb, ws) => {
                wb.properties.date1904 = date1904;
                ws.getCell("B1").value = { text: date, hyperlink: "http://example.com/" };
                ws.getCell("B1").numFmt = "yyyy-mm-dd";
                // without a number format of its own: the default date format, as for a Date value
                ws.getCell("C1").value = { text: date, hyperlink: "http://example.com/c" };
            });
            const sheet = await readPart(buffer, "xl/worksheets/sheet1.xml");
            const serial = date1904 ? 43538.25 : 45000.25;
            expect(sheet, `date1904 ${date1904}`).to.match(new RegExp(`<c r="B1" s="\\d+"><v>${serial}</v></c>`));
            expect(sheet, `date1904 ${date1904}`).to.match(new RegExp(`<c r="C1" s="\\d+"><v>${serial}</v></c>`));

            const ws = await load(buffer);
            expect(ws.getCell("B1").value, `date1904 ${date1904}`).to.deep.equal({ text: date, hyperlink: "http://example.com/" });
            expect(ws.getCell("B1").numFmt).to.equal("yyyy-mm-dd");
            expect(ws.getCell("C1").value, `date1904 ${date1904}`).to.deep.equal({ text: date, hyperlink: "http://example.com/c" });
            expect(ws.getCell("C1").numFmt).to.equal("mm-dd-yy");
        }
    });

    it("keeps a loaded hyperlinked date cell a date through a second round trip", async () => {
        const buffer = await baseWorkbookBuffer((wb, ws) => {
            ws.getCell("B1").value = { text: date, hyperlink: "http://example.com/" };
            ws.getCell("B1").numFmt = "yyyy-mm-dd";
        });
        const loaded = await load(buffer);
        const again = await load(await loaded.workbook.xlsx.writeBuffer());
        expect(again.getCell("B1").value).to.deep.equal({ text: date, hyperlink: "http://example.com/" });
    });

    it("writes a Date text as a date with the streaming writer", async () => {
        const { PassThrough } = require("stream");
        const stream = new PassThrough();
        const chunks = [];
        stream.on("data", (chunk) => chunks.push(chunk));
        const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ stream, useSharedStrings: true });
        const ws = wb.addWorksheet("Sheet1");
        ws.getCell("A1").value = { text: date, hyperlink: "http://example.com/" };
        ws.getCell("A1").numFmt = "yyyy-mm-dd";
        ws.commit();
        await wb.commit();
        const buffer = Buffer.concat(chunks);
        expect((await load(buffer)).getCell("A1").value).to.deep.equal({ text: date, hyperlink: "http://example.com/" });
        const values = [];
        for await (const sheet of new ExcelJS.stream.xlsx.WorkbookReader(Readable.from([buffer]), {
            hyperlinks: "cache",
            styles: "cache",
        })) {
            for await (const row of sheet) {
                values.push(row.getCell(1).value);
            }
        }
        expect(values).to.deep.equal([date]);
    });
});
