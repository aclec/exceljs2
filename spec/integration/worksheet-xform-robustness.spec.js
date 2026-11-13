const { baseWorkbookBuffer, editParts, tamper } = require("../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

const SHEET = "xl/worksheets/sheet1.xml";

async function load(buffer) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    return workbook;
}

// a workbook whose sheet1.xml is patched, optionally without its styles part
async function patchedBuffer(patch, { withoutStyles = false } = {}) {
    const buffer = await tamper(await baseWorkbookBuffer(), { [SHEET]: patch });
    if (!withoutStyles) {
        return buffer;
    }
    return editParts(buffer, (parts) => {
        delete parts["xl/styles.xml"];
    });
}

describe("worksheet xform robustness", () => {
    describe("a workbook without a styles part", () => {
        it("reads a styled row", async () => {
            const buffer = await patchedBuffer((xml) => xml.replace('<row r="1"', '<row r="1" s="1" customFormat="1"'), {
                withoutStyles: true,
            });
            const ws = (await load(buffer)).getWorksheet("Sheet1");
            expect(ws.getCell("A1").value).to.equal("hello");
        });

        it("reads a styled column", async () => {
            const buffer = await patchedBuffer(
                (xml) => xml.replace("<sheetData>", '<cols><col min="1" max="1" width="9" style="1"/></cols><sheetData>'),
                { withoutStyles: true }
            );
            const ws = (await load(buffer)).getWorksheet("Sheet1");
            expect(ws.getColumn(1).width).to.equal(9);
        });
    });

    describe("a shared string index outside the shared strings table", () => {
        it("rejects the load with an explicit error", async () => {
            const buffer = await patchedBuffer((xml) => xml.replace(/(<c r="A1"[^>]*t="s"[^>]*>)<v>0<\/v>/, "$1<v>99</v>"));
            let error;
            try {
                await load(buffer);
            } catch (e) {
                error = e;
            }
            expect(error).to.be.an("error");
            expect(error.message).to.match(/shared string 99/);
        });
    });

    describe("a workbook without a shared strings part", () => {
        it('reads a t="s" cell as an empty cell (its string is missing)', async () => {
            const buffer = editParts(await baseWorkbookBuffer(), (parts) => {
                delete parts["xl/sharedStrings.xml"];
            });
            const ws = (await load(buffer)).getWorksheet("Sheet1");
            expect(ws.getCell("A1").model).to.deep.equal({ address: "A1", type: ExcelJS.ValueType.Null, style: {} });
        });
    });

    describe("element names of Object.prototype members", () => {
        it("ignores them in a worksheet and in a row", async () => {
            const buffer = await patchedBuffer((xml) =>
                xml.replace("<sheetData>", "<constructor/><sheetData>").replace('<c r="A1"', '<toString/><c r="A1"')
            );
            const ws = (await load(buffer)).getWorksheet("Sheet1");
            expect(ws.getCell("A1").value).to.equal("hello");
        });
    });
});
