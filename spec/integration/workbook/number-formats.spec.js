// Number formats written and read back exactly, escapes included.
const { baseWorkbookBuffer, readPart } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

// format codes as Excel writes them for various locales
const FORMATS = [
    '#,##0.00\\ "€"',
    "#,##0\\ [$€-40C]",
    "#,##0.00\\ [$€-1];[Red]\\-#,##0.00\\ [$€-1]",
    '_-* #,##0.00\\ _€_-;\\-* #,##0.00\\ _€_-;_-* "-"??\\ _€_-;_-@_-',
    '"Total: "0;\\-0;"none"',
    "[Red]\\-#,##0.00",
    "\\#0",
    "0.0\\ \\m",
    "[$-40C]d\\ mmmm\\ yyyy",
    "@",
];

describe("number formats", () => {
    it("keep their escaped characters when read and written back", async () => {
        const source = await baseWorkbookBuffer((wb, ws) => {
            FORMATS.forEach((numFmt, index) => {
                const cell = ws.getCell(index + 2, 1);
                cell.value = 1;
                cell.numFmt = numFmt;
            });
        });
        const styles = await readPart(source, "xl/styles.xml");
        expect(styles).to.include('formatCode="#,##0.00\\ &quot;€&quot;"');

        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(source);
        const read = (workbook) => FORMATS.map((numFmt, index) => workbook.getWorksheet("Sheet1").getCell(index + 2, 1).numFmt);
        expect(read(wb)).to.deep.equal(FORMATS);
        // literal text, not a date format
        expect(wb.getWorksheet("Sheet1").getCell(9, 1).value).to.equal(1);
        expect(wb.getWorksheet("Sheet1").getCell(10, 1).value).to.be.instanceOf(Date);

        const buffer = await wb.xlsx.writeBuffer();
        expect(await readPart(buffer, "xl/styles.xml")).to.include('formatCode="#,##0.00\\ &quot;€&quot;"');
        const back = new ExcelJS.Workbook();
        await back.xlsx.load(buffer);
        expect(read(back)).to.deep.equal(FORMATS);
    });
});
