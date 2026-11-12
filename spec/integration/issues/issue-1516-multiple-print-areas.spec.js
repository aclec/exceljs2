const { PassThrough } = require("stream");
const { baseWorkbookBuffer, tamper, readPart } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

const definedNames = (xml) => xml.match(/<definedName [^>]*>[^<]*<\/definedName>/g) || [];

async function load(buffer) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    return wb;
}

describe("github issues", () => {
    describe("issue 1516 - multiple print areas", () => {
        it("writes one _xlnm.Print_Area per sheet and reads every range back", async () => {
            const buffer = await baseWorkbookBuffer((wb, ws) => {
                ws.pageSetup.printArea = "A1:B2&&D4:E5";
            });
            const names = definedNames(await readPart(buffer, "xl/workbook.xml"));
            expect(names).to.deep.equal([
                '<definedName name="_xlnm.Print_Area" localSheetId="0">' +
                    "&apos;Sheet1&apos;!$A$1:$B$2,&apos;Sheet1&apos;!$D$4:$E$5</definedName>",
            ]);

            const wb = await load(buffer);
            expect(wb.getWorksheet("Sheet1").pageSetup.printArea).to.equal("A1:B2&&D4:E5");
        });

        it("reads a defined name with several ranges and legacy duplicate names", async () => {
            const buffer = await tamper(await baseWorkbookBuffer(), {
                "xl/workbook.xml": (xml) =>
                    xml.replace(
                        "</sheets>",
                        "</sheets><definedNames>" +
                            '<definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$B$2,Sheet1!$D$4:$E$5</definedName>' +
                            '<definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$G$1:$H$2</definedName>' +
                            "</definedNames>"
                    ),
            });
            const wb = await load(buffer);
            expect(wb.getWorksheet("Sheet1").pageSetup.printArea).to.equal("A1:B2&&D4:E5&&G1:H2");

            const names = definedNames(await readPart(await wb.xlsx.writeBuffer(), "xl/workbook.xml"));
            expect(names).to.have.length(1);
        });

        it("round-trips whole-column and whole-row print areas", async () => {
            const buffer = await tamper(await baseWorkbookBuffer(), {
                "xl/workbook.xml": (xml) =>
                    xml.replace(
                        "</sheets>",
                        "</sheets><definedNames>" +
                            '<definedName name="_xlnm.Print_Area" localSheetId="0">' +
                            "'Sheet1'!$A:$C,'Sheet1'!$1:$3,'Sheet1'!$E$5</definedName>" +
                            "</definedNames>"
                    ),
            });
            const wb = await load(buffer);
            expect(wb.getWorksheet("Sheet1").pageSetup.printArea).to.equal("A:C&&1:3&&E5");

            const names = definedNames(await readPart(await wb.xlsx.writeBuffer(), "xl/workbook.xml"));
            expect(names).to.deep.equal([
                '<definedName name="_xlnm.Print_Area" localSheetId="0">' +
                    "&apos;Sheet1&apos;!$A:$C,&apos;Sheet1&apos;!$1:$3,&apos;Sheet1&apos;!$E$5</definedName>",
            ]);
        });

        it("writes absolute references, whether the areas are given with $ or not, with both writers", async () => {
            const setup = (wb) => {
                wb.addWorksheet("s").pageSetup.printArea = "A1:G20";
                const t = wb.addWorksheet("t");
                t.pageSetup.printArea = "$A$1:$B$2&&C:$D&&$3:4";
                t.pageSetup.printTitlesRow = "$1:$2";
                t.pageSetup.printTitlesColumn = "A:$B";
            };
            const expected = [
                '<definedName name="_xlnm.Print_Area" localSheetId="0">&apos;s&apos;!$A$1:$G$20</definedName>',
                '<definedName name="_xlnm.Print_Area" localSheetId="1">' +
                    "&apos;t&apos;!$A$1:$B$2,&apos;t&apos;!$C:$D,&apos;t&apos;!$3:$4</definedName>",
                '<definedName name="_xlnm.Print_Titles" localSheetId="1">&apos;t&apos;!$A:$B,&apos;t&apos;!$1:$2</definedName>',
            ];

            const wb = new ExcelJS.Workbook();
            setup(wb);
            const buffer = await wb.xlsx.writeBuffer();
            expect(definedNames(await readPart(buffer, "xl/workbook.xml"))).to.deep.equal(expected);
            const { pageSetup } = (await load(buffer)).getWorksheet("t");
            expect(pageSetup.printArea).to.equal("A1:B2&&C:D&&3:4");
            expect(pageSetup.printTitlesRow).to.equal("1:2");
            expect(pageSetup.printTitlesColumn).to.equal("A:B");

            const stream = new PassThrough();
            const chunks = [];
            stream.on("data", (chunk) => chunks.push(chunk));
            const writer = new ExcelJS.stream.xlsx.WorkbookWriter({ stream });
            setup(writer);
            await writer.commit();
            expect(definedNames(await readPart(Buffer.concat(chunks), "xl/workbook.xml"))).to.deep.equal(expected);
        });
    });
});
