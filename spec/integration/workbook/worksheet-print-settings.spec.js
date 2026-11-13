// Round trips of the print settings and views of a worksheet: page breaks, first page number, views, header and footer.
const { baseWorkbookBuffer, tamper, readPart } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");
const ModelContainer = verquire("doc/modelcontainer");

const SHEET = "xl/worksheets/sheet1.xml";

async function load(buffer) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    return wb;
}

// a workbook whose worksheet has a page break (patched by `patch`), loaded
async function withRowBreaks(rowBreaksXml) {
    const base = await baseWorkbookBuffer((wb, ws) => ws.getRow(3).addPageBreak());
    return tamper(base, { [SHEET]: (xml) => xml.replace(/<rowBreaks.*<\/rowBreaks>/, rowBreaksXml) });
}

// a workbook whose <pageSetup> is replaced
async function withPageSetup(pageSetupXml) {
    const base = await baseWorkbookBuffer();
    return tamper(base, { [SHEET]: (xml) => xml.replace(/<pageSetup[^>]*\/>/, pageSetupXml) });
}

async function pageSetupXml(wb) {
    const xml = await readPart(await wb.xlsx.writeBuffer(), SHEET);
    return xml.match(/<pageSetup[^>]*\/>/)[0];
}

describe("Worksheet print settings round trip", () => {
    describe("row breaks", () => {
        it("reads the row breaks back", async () => {
            const wb = new ExcelJS.Workbook();
            const ws = wb.addWorksheet("s");
            ws.getCell("A1").value = 1;
            ws.getRow(3).addPageBreak();
            ws.getRow(5).addPageBreak(2, 4);
            const buffer = await wb.xlsx.writeBuffer();

            const loaded = (await load(buffer)).getWorksheet("s");
            expect(loaded.rowBreaks).to.deep.equal([
                { id: 3, max: 16383, man: 1 },
                { id: 5, max: 3, min: 1, man: 1 },
            ]);

            const rewritten = await readPart(await loaded.workbook.xlsx.writeBuffer(), SHEET);
            expect(rewritten).to.contain(
                '<rowBreaks count="2" manualBreakCount="2"><brk id="3" max="16383" man="1"/><brk id="5" max="3" man="1" min="1"/></rowBreaks>'
            );
        });

        it("reads the row breaks into a ModelContainer", async () => {
            const buffer = await withRowBreaks('<rowBreaks count="1" manualBreakCount="1"><brk id="9" max="16383" man="1"/></rowBreaks>');
            const container = new ModelContainer();
            await container.xlsx.load(buffer);
            expect(container.model.worksheets[0].rowBreaks).to.deep.equal([{ id: 9, max: 16383, man: 1 }]);
        });

        it("reads a worksheet without row breaks as no break", async () => {
            const wb = await load(await baseWorkbookBuffer());
            expect(wb.getWorksheet("Sheet1").rowBreaks).to.deep.equal([]);
        });

        it("reads more than 1023 breaks (the limit of Excel), and writes at most 1023", async () => {
            const breaks = (n) => Array.from({ length: n }, (_, i) => `<brk id="${i + 1}" max="16383" man="1"/>`).join("");
            const wb = await load(await withRowBreaks(`<rowBreaks>${breaks(1100)}</rowBreaks>`));
            const ws = wb.getWorksheet("Sheet1");
            expect(ws.rowBreaks).to.have.length(1100);

            let error;
            await wb.xlsx.writeBuffer().catch((e) => {
                error = e;
            });
            expect(error).to.be.an("error");
            expect(error.message).to.equal("Too many row breaks in a worksheet: 1100 (Excel allows at most 1023)");

            ws.rowBreaks.length = 1023;
            const reloaded = await load(await wb.xlsx.writeBuffer());
            expect(reloaded.getWorksheet("Sheet1").rowBreaks).to.have.length(1023);
        });

        it("reads the breaks written by 4.x (max 16838) and other values beyond the worksheet as its last row or column", async () => {
            const wb = await load(
                await withRowBreaks(
                    '<rowBreaks><brk id="10" max="16838" man="1"/><brk id="2000000" min="20000" max="4294967295"/></rowBreaks>'
                )
            );
            expect(wb.getWorksheet("Sheet1").rowBreaks).to.deep.equal([
                { id: 10, max: 16383, man: 1 },
                { id: 1048576, max: 16383, man: 0, min: 16383 },
            ]);
        });

        it("rejects a break whose attributes are not unsigned integers", async () => {
            for (const brk of [
                '<brk id="-1"/>',
                '<brk id="x"/>',
                '<brk id="1" max="1.5"/>',
                '<brk id="1" min="1e3"/>',
                '<brk id="99999999999"/>',
            ]) {
                let error;
                await load(await withRowBreaks(`<rowBreaks>${brk}</rowBreaks>`)).catch((e) => {
                    error = e;
                });
                expect(error, brk).to.be.an("error");
                expect(error.message, brk).to.match(/^Invalid page break <brk /);
            }
        });
    });

    describe("first page number", () => {
        it("keeps the first page number absent (automatic numbering) when the file has none", async () => {
            const ws = (await load(await baseWorkbookBuffer())).getWorksheet("Sheet1");
            expect(ws.pageSetup).to.not.have.property("firstPageNumber");
            expect(ws.pageSetup).to.not.have.property("useFirstPageNumber");
            const xml = await pageSetupXml(ws.workbook);
            expect(xml).to.not.contain("firstPageNumber");
            expect(xml).to.not.contain("useFirstPageNumber");
        });

        it("keeps a first page number that the file does not use (useFirstPageNumber absent)", async () => {
            const wb = await load(await withPageSetup('<pageSetup paperSize="9" firstPageNumber="5"/>'));
            const ws = wb.getWorksheet("Sheet1");
            expect(ws.pageSetup.firstPageNumber).to.equal(5);
            expect(ws.pageSetup.useFirstPageNumber).to.equal(false);
            const xml = await pageSetupXml(wb);
            expect(xml).to.contain('firstPageNumber="5"');
            expect(xml).to.not.contain("useFirstPageNumber");
        });

        it("keeps a first page number that the file uses", async () => {
            for (const flag of ["1", "true"]) {
                const wb = await load(await withPageSetup(`<pageSetup firstPageNumber="5" useFirstPageNumber="${flag}"/>`));
                const ws = wb.getWorksheet("Sheet1");
                expect(ws.pageSetup.firstPageNumber).to.equal(5);
                expect(ws.pageSetup.useFirstPageNumber).to.equal(true);
                expect(await pageSetupXml(wb)).to.contain('firstPageNumber="5" useFirstPageNumber="1"');
            }
        });

        it("reads useFirstPageNumber without firstPageNumber as the first page number 1 (the default it uses)", async () => {
            const wb = await load(await withPageSetup('<pageSetup useFirstPageNumber="1"/>'));
            const ws = wb.getWorksheet("Sheet1");
            expect(ws.pageSetup.firstPageNumber).to.equal(1);
            expect(ws.pageSetup.useFirstPageNumber).to.equal(true);
            expect(await pageSetupXml(wb)).to.contain('firstPageNumber="1" useFirstPageNumber="1"');
        });

        it("writes the first page number set on a new worksheet with useFirstPageNumber", async () => {
            const wb = new ExcelJS.Workbook();
            wb.addWorksheet("s", { pageSetup: { firstPageNumber: 3 } });
            expect(await pageSetupXml(wb)).to.contain('firstPageNumber="3" useFirstPageNumber="1"');
        });
    });

    describe("views, header and footer", () => {
        it("reads a file without views nor header / footer as a new worksheet / workbook has them", async () => {
            const buffer = await baseWorkbookBuffer();
            const wb = await load(buffer);
            const ws = wb.getWorksheet("Sheet1");
            expect(ws.views).to.deep.equal([]);
            expect(ws.headerFooter).to.deep.equal(new ExcelJS.Workbook().addWorksheet("x").headerFooter);
            expect(wb.views).to.deep.equal([]);
            expect(ws.model.views).to.deep.equal([]);
            expect(wb.model.views).to.deep.equal([]);

            // and writes them back as it read them: no element
            const rewritten = await wb.xlsx.writeBuffer();
            const sheetXml = await readPart(rewritten, SHEET);
            const workbookXml = await readPart(rewritten, "xl/workbook.xml");
            expect(sheetXml).to.not.contain("<sheetViews");
            expect(sheetXml).to.not.contain("<headerFooter");
            expect(workbookXml).to.not.contain("<bookViews");
            const again = await (await load(rewritten)).xlsx.writeBuffer();
            expect(await readPart(again, SHEET)).to.equal(sheetXml);
            expect(await readPart(again, "xl/workbook.xml")).to.equal(workbookXml);

            // the lists can be added to
            ws.views.push({ state: "frozen", ySplit: 1 });
            wb.views.push({ x: 0, y: 0, width: 10000, height: 20000, visibility: "visible" });
            ws.headerFooter.oddFooter = "&P";
            const written = await wb.xlsx.writeBuffer();
            expect(await readPart(written, SHEET))
                .to.contain("<sheetViews>")
                .and.to.contain("<oddFooter>&amp;P</oddFooter>");
            expect(await readPart(written, "xl/workbook.xml")).to.contain("<bookViews>");
        });

        it("reads a file without views nor header / footer into a ModelContainer as empty views and the default header / footer", async () => {
            const container = new ModelContainer();
            await container.xlsx.load(await baseWorkbookBuffer());
            const [worksheet] = container.model.worksheets;
            expect(worksheet.views).to.deep.equal([]);
            expect(worksheet.headerFooter).to.deep.equal(new ExcelJS.Workbook().addWorksheet("x").headerFooter);
            expect(container.model.views).to.deep.equal([]);
        });
    });
});
