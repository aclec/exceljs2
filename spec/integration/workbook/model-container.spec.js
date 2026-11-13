// A ModelContainer holds the model read from a file, without a Workbook: writing it back must not need the members that
// only a Workbook model has (pivotTables).
const ExcelJS = verquire("exceljs");

describe("ModelContainer", () => {
    it("writes back the model of a file it loaded", async () => {
        const workbook = new ExcelJS.Workbook();
        workbook.addWorksheet("Sheet1").getCell("A1").value = 42;
        const buffer = await workbook.xlsx.writeBuffer();

        const container = new ExcelJS.ModelContainer();
        await container.xlsx.load(buffer);
        const written = await container.xlsx.writeBuffer();

        const reread = new ExcelJS.Workbook();
        await reread.xlsx.load(written);
        expect(reread.getWorksheet("Sheet1").getCell("A1").value).to.equal(42);
    });
});
