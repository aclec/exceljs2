const { Readable } = require("stream");
const { tamper } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

// <workbookPr date1904> is an xsd:boolean: "true" is as valid as "1"
async function workbook1904(attribute) {
    const wb = new ExcelJS.Workbook();
    wb.properties.date1904 = true;
    wb.addWorksheet("Sheet1").getCell("B4").value = new Date("1904-01-01T00:00:00.000Z");
    const buffer = await wb.xlsx.writeBuffer();
    return tamper(buffer, { "xl/workbook.xml": (xml) => xml.replace('date1904="1"', `date1904="${attribute}"`) });
}

describe("github issues", () => {
    describe('issue 219 - date1904="true" (xsd:boolean)', () => {
        it('reads date1904="true" as the 1904 date system', async () => {
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.load(await workbook1904("true"));
            expect(wb.properties.date1904).to.equal(true);
            expect(wb.getWorksheet("Sheet1").getCell("B4").value.toISOString()).to.equal("1904-01-01T00:00:00.000Z");
        });

        it('reads date1904="false" as the 1900 date system', async () => {
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.load(await workbook1904("false"));
            expect(wb.properties.date1904).to.equal(false);
            // the serial number written for 1904-01-01 in the 1904 system (0) is a date of 1899 in the 1900 system
            expect(wb.getWorksheet("Sheet1").getCell("B4").value.toISOString()).to.equal("1899-12-30T00:00:00.000Z");
        });

        it('streams date1904="true" as the 1904 date system', async () => {
            const reader = new ExcelJS.stream.xlsx.WorkbookReader(Readable.from([await workbook1904("true")]), { styles: "cache" });
            const values = [];
            for await (const worksheet of reader) {
                for await (const row of worksheet) {
                    values.push(row.getCell(2).value);
                }
            }
            expect(reader.properties.model.date1904).to.equal(true);
            expect(values.map((value) => value.toISOString())).to.deep.equal(["1904-01-01T00:00:00.000Z"]);
        });
    });
});
