const { baseWorkbookBuffer, tamper } = require("../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

const WORKBOOK = "xl/workbook.xml";

async function load(patch) {
    const buffer = await baseWorkbookBuffer((wb) => {
        wb.definedNames.add("Sheet1!$A$1", "first");
    });
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(await tamper(buffer, { [WORKBOOK]: patch }));
    return wb;
}

describe("xl/workbook.xml values", () => {
    it("throws a clear Error for a <sheet> without name (CT_Sheet requires it)", async () => {
        const error = await load((xml) => xml.replace(/ name="Sheet1"/, "")).then(
            () => undefined,
            (e) => e
        );
        expect(error).to.be.an.instanceOf(Error);
        expect(error.message).to.equal("Invalid xl/workbook.xml: <sheet> without name");
    });

    it("reads a state that is not one of ST_SheetState as visible", async () => {
        const hidden = await load((xml) => xml.replace(/state="visible"/, 'state="hidden"'));
        expect(hidden.getWorksheet("Sheet1").state).to.equal("hidden");
        const invalid = await load((xml) => xml.replace(/state="visible"/, 'state="gone"'));
        expect(invalid.getWorksheet("Sheet1").state).to.equal("visible");
    });

    it("skips a <definedName> without name (CT_DefinedName requires it)", async () => {
        const wb = await load((xml) => xml.replace(/<definedName name="first">/, "<definedName>"));
        expect(wb.definedNames.model).to.deep.equal([]);
    });
});
