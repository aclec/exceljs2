// Regression tests for hostile or malformed workbook-level parts (styles, comments): each case used to crash the load
// or let attributes of the file overwrite the internal fields of a model.
const { Readable } = require("stream");
const { baseWorkbookBuffer, tamper } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

async function load(buffer) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    return wb;
}

describe("hostile workbook-level parts", function () {
    this.timeout(20000);

    describe("styles", () => {
        // a conditional formatting rule (or a table column) referring to a differential format of a styles part
        // that has no <dxfs>
        it("loads a dxfId when the styles part has no differential formats", async () => {
            const buffer = await tamper(
                await baseWorkbookBuffer((wb, ws) => {
                    ws.addConditionalFormatting({
                        ref: "A1:A3",
                        rules: [{ type: "expression", formulae: ["A1>1"], style: { font: { bold: true } } }],
                    });
                }),
                { "xl/styles.xml": (xml) => xml.replace(/<dxfs[^>]*>.*<\/dxfs>|<dxfs[^>]*\/>/s, "") }
            );
            const wb = await load(buffer);
            const [cf] = wb.getWorksheet("Sheet1").conditionalFormattings;
            expect(cf.rules).to.have.length(1);
            expect(cf.rules[0].style).to.equal(undefined);
        });

        it("loads a style id when the styles part has no cell formats", async () => {
            const buffer = await tamper(
                await baseWorkbookBuffer((wb, ws) => {
                    ws.getCell("A2").font = { bold: true };
                }),
                { "xl/styles.xml": (xml) => xml.replace(/<cellXfs[^>]*>.*<\/cellXfs>/s, "") }
            );
            const wb = await load(buffer);
            expect(wb.getWorksheet("Sheet1").getCell("A2").font).to.equal(undefined);
        });
    });

    describe("styles in the streaming reader", () => {
        // with styles: "ignore" (the default), the reader's style manager has no parsed part
        it('reads a row with s="0" when the styles are ignored', async () => {
            const buffer = await tamper(await baseWorkbookBuffer(), {
                "xl/worksheets/sheet1.xml": (xml) => xml.replace('<row r="1"', '<row r="1" s="0" customFormat="1"'),
            });
            const reader = new ExcelJS.stream.xlsx.WorkbookReader(Readable.from([buffer]), {});
            const values = [];
            for await (const worksheet of reader) {
                for await (const row of worksheet) {
                    values.push(row.getCell(1).value);
                }
            }
            expect(values).to.deep.equal(["hello"]);
        });
    });

    describe("comments", () => {
        it("does not let the attributes of <comment> overwrite the type or the text of the note", async () => {
            const buffer = await tamper(
                await baseWorkbookBuffer((wb, ws) => {
                    ws.getCell("A1").note = "hello";
                }),
                { "xl/comments1.xml": (xml) => xml.replace("<comment ", '<comment type="x" note="y" texts="z" ') }
            );
            const wb = await load(buffer);
            const cell = wb.getWorksheet("Sheet1").getCell("A1");
            // a note of one plain text reads back as that text
            expect(cell.note).to.equal("hello");
            expect(cell.model.comment.type).to.equal("note");
        });

        it('does not change the prototype of a note through a "__proto__" attribute', async () => {
            const buffer = await tamper(
                await baseWorkbookBuffer((wb, ws) => {
                    ws.getCell("A1").note = "hello";
                }),
                { "xl/comments1.xml": (xml) => xml.replace("<comment ", '<comment __proto__="x" constructor="y" ') }
            );
            const wb = await load(buffer);
            const cell = wb.getWorksheet("Sheet1").getCell("A1");
            expect(cell.note).to.equal("hello");
            expect(Object.getPrototypeOf(cell.model.comment)).to.equal(Object.prototype);
            expect({}.x).to.equal(undefined);
        });
    });
});
