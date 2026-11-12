// Round trips of worksheet parts whose model was not what its type says (found while removing the type assertions of the
// sheet xforms).
const { baseWorkbookBuffer, tamper, readPart } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

const SHEET = "xl/worksheets/sheet1.xml";

async function load(buffer) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    return wb;
}

async function withConditionalFormatting(xml) {
    return tamper(await baseWorkbookBuffer(), { [SHEET]: (sheet) => sheet.replace("</sheetData>", `</sheetData>${xml}`) });
}

describe("Worksheet parts round trip", () => {
    it("does not write an empty <conditionalFormatting> for a rule of a type the model does not have", async () => {
        // CT_ConditionalFormatting requires a cfRule: the empty element made a corrupt file
        const buffer = await withConditionalFormatting(
            '<conditionalFormatting sqref="A1:A5"><cfRule type="duplicateValues" dxfId="0" priority="1"/></conditionalFormatting>'
        );
        const wb = await load(buffer);
        expect(wb.getWorksheet("Sheet1").conditionalFormattings).to.deep.equal([]);
        expect(await readPart(await wb.xlsx.writeBuffer(), SHEET)).not.to.include("<conditionalFormatting");
    });

    it("writes the error text of a hyperlink as an error value", async () => {
        // a hyperlink on an error cell (as read from a file): it was written as the shared string "[object Object]"
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet("Sheet1");
        ws.getCell("A1").model = { address: "A1", type: ExcelJS.ValueType.Hyperlink, text: { error: "#N/A" }, hyperlink: "http://x/" };
        const buffer = await wb.xlsx.writeBuffer();
        expect(await readPart(buffer, SHEET)).to.include('<c r="A1" t="e"><v>#N/A</v></c>');
        expect((await load(buffer)).getWorksheet("Sheet1").getCell("A1").value).to.deep.equal({
            text: { error: "#N/A" },
            hyperlink: "http://x/",
        });
    });

    it("rejects an image id that is not an index of the workbook media", async () => {
        // "length" found the length of the media array (1 here), written as an image named "undefined"
        const wb = new ExcelJS.Workbook();
        wb.addImage({ buffer: Buffer.from("not checked"), extension: "png" });
        wb.addWorksheet("Sheet1").addBackgroundImage("length");
        let error;
        try {
            await wb.xlsx.writeBuffer();
        } catch (e) {
            error = e;
        }
        expect(error).to.be.an("error");
        expect(error.message).to.equal("Unknown image length");
    });
});
