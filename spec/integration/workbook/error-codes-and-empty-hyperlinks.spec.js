// Cells read from a file with the error codes newer than the classic seven (#SPILL!, #CALC!...), and hyperlinks on
// cells without value: both readers keep them, and the loaded worksheet can be edited.
const { Readable } = require("stream");
const { baseWorkbookBuffer, tamper } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

const SHEET = "xl/worksheets/sheet1.xml";

// A1 "h", A2 =SEQUENCE(2) cached #SPILL!, B2 #CALC!, A4 "tail"
async function errorCodesBuffer() {
    const base = await baseWorkbookBuffer((wb, ws) => {
        ws.getCell("A1").value = "h";
        ws.getCell("A2").value = { formula: "SEQUENCE(2)", result: 0 };
        ws.getCell("B2").value = 0;
        ws.getCell("A4").value = "tail";
    });
    return tamper(base, {
        [SHEET]: (xml) =>
            xml
                .replace(/<c r="A2"([^>]*)><f>SEQUENCE\(2\)<\/f><v>0<\/v><\/c>/, '<c r="A2"$1 t="e"><f>SEQUENCE(2)</f><v>#SPILL!</v></c>')
                .replace(/<c r="B2"([^>]*)><v>0<\/v><\/c>/, '<c r="B2"$1 t="e"><v>#CALC!</v></c>'),
    });
}

describe("cells read from a file", () => {
    describe("with recent error codes", () => {
        it("are read by the streaming reader", async () => {
            const values = [];
            for await (const ws of new ExcelJS.stream.xlsx.WorkbookReader(Readable.from([await errorCodesBuffer()]))) {
                for await (const row of ws) {
                    if (row.number === 2) {
                        values.push(row.getCell(1).value, row.getCell(2).value);
                    }
                }
            }
            expect(values).to.deep.equal([{ formula: "SEQUENCE(2)", result: { error: "#SPILL!" } }, { error: "#CALC!" }]);
        });

        it("are kept when the loaded worksheet is edited and its values copied", async () => {
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.load(await errorCodesBuffer());
            const ws = wb.getWorksheet("Sheet1");
            ws.insertRow(1, ["new"]);
            ws.spliceColumns(1, 0, ["c"]);
            ws.duplicateRow(3, 1, true);
            expect(ws.getCell("A1").value).to.equal("c");
            expect(ws.getCell("B1").value).to.equal("new");
            expect(ws.getCell("B3").value).to.deep.equal({ formula: "SEQUENCE(2)", result: { error: "#SPILL!" } });
            expect(ws.getCell("C3").value).to.deep.equal({ error: "#CALC!" });
            expect(ws.getCell("B6").value).to.equal("tail");
            const copy = wb.addWorksheet("copy");
            copy.getCell("A1").value = ws.getCell("C3").value;
            expect(copy.getCell("A1").value).to.deep.equal({ error: "#CALC!" });
        });
    });

    describe("with a hyperlink and no value", () => {
        it("have an empty text", async () => {
            const base = await baseWorkbookBuffer((wb, ws) => {
                ws.getCell("A2").value = { text: "x", hyperlink: "https://example.com" };
                ws.getCell("A2").font = { bold: true };
            });
            // a styled empty cell with a hyperlink: <c r="A2" s="1"/>
            const buffer = await tamper(base, {
                [SHEET]: (xml) => xml.replace(/<c r="A2"([^>]*?)( t="s")?><v>\d+<\/v><\/c>/, '<c r="A2"$1/>'),
            });
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.load(buffer);
            const cell = wb.getWorksheet("Sheet1").getCell("A2");
            expect(cell.hyperlink).to.equal("https://example.com");
            expect([cell.text, cell.html, cell.toString()]).to.deep.equal(["", "", ""]);
        });
    });
});
