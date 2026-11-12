// Merged-cell caps: every cell of a merge is materialised, so merges are capped per worksheet (one full column) and per
// workbook (workbook.maxMergedCells, the maxMergedCells load option, one full column by default). Reading and writing apply
// the same defaults, so a workbook ExcelJS writes with the defaults can be read back with the defaults.
const { PassThrough } = require("stream");
const { baseWorkbookBuffer, tamper } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

const MAX_PER_SHEET = 1048576;
const DEFAULT_MAX_PER_WORKBOOK = MAX_PER_SHEET;

// a workbook with sheets Sheet1..SheetN, the merges of mergeRefs(i) on sheet i
async function mergedSheetsBuffer(count, mergeRefs) {
    const buffer = await baseWorkbookBuffer((wb) => {
        for (let i = 2; i <= count; i++) {
            wb.addWorksheet(`Sheet${i}`).getCell("A1").value = i;
        }
    });
    const patches = {};
    for (let i = 1; i <= count; i++) {
        const refs = mergeRefs(i);
        const xml = refs.map((ref) => `<mergeCell ref="${ref}"/>`).join("");
        patches[`xl/worksheets/sheet${i}.xml`] = (sheet) =>
            sheet.replace("</sheetData>", `</sheetData><mergeCells count="${refs.length}">${xml}</mergeCells>`);
    }
    return tamper(buffer, patches);
}

async function loadError(buffer, options) {
    const start = Date.now();
    const error = await new ExcelJS.Workbook().xlsx.load(buffer, options).then(
        () => undefined,
        (e) => e
    );
    return { error, ms: Date.now() - start };
}

describe("Merged-cell limits", function () {
    this.timeout(30000);

    it("has a default budget of one full worksheet, for a workbook, a load and the streaming writer", () => {
        expect(new ExcelJS.Workbook().maxMergedCells).to.equal(1048576);
        const stream = new PassThrough();
        stream.resume();
        expect(new ExcelJS.stream.xlsx.WorkbookWriter({ stream }).maxMergedCells).to.equal(1048576);
    });

    it("reads back a workbook written with the defaults, whose sheets together hold merged cells up to the budget", async () => {
        const wb = new ExcelJS.Workbook();
        for (const name of ["a", "b"]) {
            const ws = wb.addWorksheet(name);
            // 500 merges of 1,000 cells: 500,000 merged cells per sheet, 1,000,000 in the workbook
            for (let i = 0; i < 500; i++) {
                ws.mergeCells(i * 100 + 1, 1, i * 100 + 100, 10);
            }
        }
        // a workbook built with the defaults cannot exceed the budget of a load with the defaults
        expect(() => wb.getWorksheet("b").mergeCells("L1:L48577")).to.throw(/workbook would exceed 1048576 merged cells/);
        const reloaded = new ExcelJS.Workbook();
        await reloaded.xlsx.load(await wb.xlsx.writeBuffer());
        expect(reloaded.getWorksheet("b").getCell("J50000").master.address).to.equal("A49901");
    });

    it("reads back a workbook built with a raised budget when the load is given the same budget", async () => {
        const wb = new ExcelJS.Workbook();
        wb.maxMergedCells = 1100000;
        for (const name of ["a", "b"]) {
            const ws = wb.addWorksheet(name);
            // 550 merges of 1,000 cells: 550,000 merged cells per sheet, 1,100,000 in the workbook
            for (let i = 0; i < 550; i++) {
                ws.mergeCells(i * 100 + 1, 1, i * 100 + 100, 10);
            }
        }
        const buffer = await wb.xlsx.writeBuffer();
        const { error } = await loadError(buffer);
        expect(error).to.be.an("error");
        expect(error.message).to.match(/workbook would exceed 1048576 merged cells/);
        const reloaded = new ExcelJS.Workbook();
        await reloaded.xlsx.load(buffer, { maxMergedCells: 1100000 });
        expect(reloaded.getWorksheet("b").getCell("J55000").master.address).to.equal("A54901");
    });

    it("caps the merged cells of one worksheet on read, whatever the workbook budget", async () => {
        const buffer = await mergedSheetsBuffer(1, () => ["B1:B1048576", "C1:C2"]);
        const { error, ms } = await loadError(buffer, { maxMergedCells: Infinity });
        expect(error).to.be.an("error");
        expect(error.message).to.match(/worksheet would exceed 1048576 merged cells/);
        expect(ms).to.be.below(3000); // rejected before the cells are created
    });

    it("caps the merged cells of the workbook on read with a default budget of one full worksheet", async () => {
        // a few KB of XML that would create 2 x 1048576 merged cells (about 1.2 GB of heap)
        const buffer = await mergedSheetsBuffer(2, () => ["B1:B1048576"]);
        const { error, ms } = await loadError(buffer);
        expect(error).to.be.an("error");
        expect(error.message).to.match(new RegExp(`workbook would exceed ${DEFAULT_MAX_PER_WORKBOOK} merged cells`));
        expect(error.message).to.match(/maxMergedCells/);
        expect(ms).to.be.below(3000); // rejected before the cells are created
    });

    it("applies the maxMergedCells load option to the whole workbook", async () => {
        // 600 merged cells per sheet
        const buffer = await mergedSheetsBuffer(2, () => ["B1:K60"]);
        const { error } = await loadError(buffer, { maxMergedCells: 1000 });
        expect(error).to.be.an("error");
        expect(error.message).to.match(/workbook would exceed 1000 merged cells/);

        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buffer, { maxMergedCells: 1200 });
        expect(wb.getWorksheet("Sheet2").getCell("K60").master.address).to.equal("B1");
        // the loaded workbook keeps the budget it was loaded with
        expect(wb.maxMergedCells).to.equal(1200);
        expect(() => wb.getWorksheet("Sheet1").mergeCells("M1:M2")).to.throw(/workbook would exceed 1200 merged cells/);
    });

    it("rejects an invalid maxMergedCells load option", async () => {
        const buffer = await baseWorkbookBuffer();
        for (const value of [-1, NaN, "100", null]) {
            const { error } = await loadError(buffer, { maxMergedCells: value });
            expect(error).to.be.an.instanceOf(TypeError);
            expect(error.message).to.match(/maxMergedCells/);
        }
    });

    it("caps the merged cells of a workbook built in memory with the same budget, which can be raised", () => {
        const wb = new ExcelJS.Workbook();
        expect(wb.maxMergedCells).to.equal(DEFAULT_MAX_PER_WORKBOOK);
        wb.maxMergedCells = 10;
        wb.addWorksheet("a").mergeCells("A1:A6");
        const ws = wb.addWorksheet("b");
        expect(() => ws.mergeCells("A1:A5")).to.throw(/workbook would exceed 10 merged cells/);
        expect(ws.getCell("A2").isMerged).to.equal(false);
        ws.mergeCells("A1:A4");
        // a splice that widens a merge is checked against the budget too
        expect(() => ws.spliceRows(2, 0, [])).to.throw(/would exceed 10 merged cells/);
        wb.maxMergedCells = 11;
        ws.spliceRows(2, 0, []);
        expect(ws.getCell("A5").master.address).to.equal("A1");
        // removing a worksheet frees its share
        wb.removeWorksheet(wb.getWorksheet("a").id);
        ws.mergeCells("B1:B6");
    });

    it("checks the workbook budget in constant time per merge, whatever the sheet ids", async () => {
        // 8000 two-cell merges on a worksheet whose sheetId is 65535: summing the merged cells of every worksheet id on each
        // merge took 7 s
        const merges = [];
        for (let row = 1; row <= 8000; row++) {
            merges.push(`B${row}:C${row}`);
        }
        const buffer = await tamper(await mergedSheetsBuffer(1, () => merges), {
            "xl/workbook.xml": (xml) => xml.replace('sheetId="1"', 'sheetId="65535"'),
        });
        const start = Date.now();
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buffer);
        const ms = Date.now() - start;
        expect(wb.getWorksheet(65535).getCell("C8000").master.address).to.equal("B8000");
        expect(ms).to.be.below(2000);
    });

    it("keeps the merged cells of the workbook up to date as merges and worksheets change", () => {
        const wb = new ExcelJS.Workbook();
        wb.maxMergedCells = 10;
        const a = wb.addWorksheet("a");
        const b = wb.addWorksheet("b");
        a.mergeCells("A1:A4"); // 4
        b.mergeCells("A1:A4"); // 8
        expect(() => b.mergeCells("C1:C3")).to.throw(/workbook would exceed 10 merged cells/);
        // unmerging frees cells
        a.unMergeCells("A1");
        b.mergeCells("C1:C3"); // 7
        b.mergeCells("E1:E3"); // 10
        expect(() => a.mergeCells("A1:A2")).to.throw(/workbook would exceed 10 merged cells/);
        // a splice that removes a merge frees its cells
        b.spliceRows(1, 4); // the merges of rows 1-4 are removed: 0
        a.mergeCells("A1:J1"); // 10
        expect(() => b.mergeCells("A1:A2")).to.throw(/workbook would exceed 10 merged cells/);
        // replacing the model of a worksheet replaces its merges
        a.model = { ...a.model, mergeCells: ["B1:B2"] }; // 2
        b.mergeCells("A1:H1"); // 10
        expect(() => b.mergeCells("A3:A4")).to.throw(/workbook would exceed 10 merged cells/);
        // removing a worksheet frees its share, once
        wb.removeWorksheet(a.id); // 8
        a.destroy();
        b.mergeCells("A3:A4"); // 10
        expect(() => b.mergeCells("A6:A7")).to.throw(/workbook would exceed 10 merged cells/);
        // a removed worksheet no longer counts for the workbook
        a.mergeCells("D1:D5");
        expect(() => b.mergeCells("A6:A7")).to.throw(/workbook would exceed 10 merged cells/);
        // replacing the model of the workbook replaces all its merges
        wb.model = { ...wb.model, worksheets: [] };
        const c = wb.addWorksheet("c");
        c.mergeCells("A1:J1");
    });

    describe("streaming writer", () => {
        function writer(options) {
            const stream = new PassThrough();
            stream.resume();
            return new ExcelJS.stream.xlsx.WorkbookWriter({ stream, ...options });
        }

        it("caps the merged cells of one worksheet", () => {
            // rejected before the cells are created
            const ws = writer({ maxMergedCells: Infinity }).addWorksheet("a");
            expect(() => ws.mergeCells("A1:B1048576")).to.throw(/worksheet would exceed 1048576 merged cells/);
            expect(ws.findCell("A2")).to.equal(undefined);
        });

        it("caps the merged cells of the workbook with the maxMergedCells option", () => {
            const wb = writer({ maxMergedCells: 10 });
            wb.addWorksheet("a").mergeCells("A1:A6");
            const ws = wb.addWorksheet("b");
            expect(() => ws.mergeCells("A1:A5")).to.throw(/workbook would exceed 10 merged cells/);
            ws.mergeCells("A1:A4");
        });
    });
});
