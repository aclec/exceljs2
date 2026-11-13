// The dates of the data validations are serial dates of the date system of the workbook: they were written and read in
// the 1900 date system, 4 years and a day off in a workbook using the 1904 date system.
const { baseWorkbookBuffer, tamper, readPart } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

const JAN_1_2016 = new Date(Date.UTC(2016, 0, 1));
// 2016-01-01: 42370 in the 1900 date system, 42370 - 1462 in the 1904 one
const SERIAL_1900 = 42370;
const SERIAL_1904 = 40908;

async function writeDateValidation(date1904) {
    return baseWorkbookBuffer((wb, ws) => {
        wb.properties.date1904 = date1904;
        ws.getCell("B2").dataValidation = { type: "date", operator: "lessThan", formulae: [JAN_1_2016] };
    });
}

async function load(buffer) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    return wb;
}

describe("Data validations in a 1904 workbook", () => {
    it("writes a date as a serial date of the 1904 date system", async () => {
        const sheet = await readPart(await writeDateValidation(true), "xl/worksheets/sheet1.xml");
        expect(sheet).to.include(`<formula1>${SERIAL_1904}</formula1>`);
    });

    it("still writes a date of a 1900 workbook in the 1900 date system", async () => {
        const sheet = await readPart(await writeDateValidation(false), "xl/worksheets/sheet1.xml");
        expect(sheet).to.include(`<formula1>${SERIAL_1900}</formula1>`);
    });

    it("reads back the date written", async () => {
        const wb = await load(await writeDateValidation(true));
        expect(wb.properties.date1904).to.equal(true);
        expect(wb.getWorksheet("Sheet1").getCell("B2").dataValidation.formulae).to.deep.equal([JAN_1_2016]);
    });

    it("reads the serial date of a 1904 file in the 1904 date system", async () => {
        // a file written by another application: a 1904 workbook with a date rule
        const buffer = await tamper(await baseWorkbookBuffer(), {
            "xl/workbook.xml": (xml) => xml.replace(/<workbookPr[^>]*\/>/, '<workbookPr date1904="1"/>'),
            "xl/worksheets/sheet1.xml": (xml) =>
                xml.replace(
                    "</sheetData>",
                    `</sheetData><dataValidations count="1"><dataValidation type="date" operator="between" sqref="C3:C4">` +
                        `<formula1>${SERIAL_1904}</formula1><formula2>${SERIAL_1904 + 1}</formula2></dataValidation></dataValidations>`
                ),
        });
        const ws = (await load(buffer)).getWorksheet("Sheet1");
        expect(ws.getCell("C3").dataValidation.formulae).to.deep.equal([JAN_1_2016, new Date(Date.UTC(2016, 0, 2))]);
        expect(ws.getCell("C4").dataValidation.formulae).to.deep.equal([JAN_1_2016, new Date(Date.UTC(2016, 0, 2))]);
    });

    it("keeps the other formulae of a date rule as text", async () => {
        const buffer = await tamper(await baseWorkbookBuffer(), {
            "xl/workbook.xml": (xml) => xml.replace(/<workbookPr[^>]*\/>/, '<workbookPr date1904="1"/>'),
            "xl/worksheets/sheet1.xml": (xml) =>
                xml.replace(
                    "</sheetData>",
                    '</sheetData><dataValidations count="1"><dataValidation type="date" operator="greaterThan" sqref="C3">' +
                        "<formula1>TODAY()</formula1></dataValidation></dataValidations>"
                ),
        });
        const ws = (await load(buffer)).getWorksheet("Sheet1");
        expect(ws.getCell("C3").dataValidation.formulae).to.deep.equal(["TODAY()"]);
    });
});
