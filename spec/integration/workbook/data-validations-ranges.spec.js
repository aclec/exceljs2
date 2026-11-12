// Data validations on whole columns / whole rows ('A:A', 'A:C', '1:1'): kept as one rectangle, never expanded per cell.
const { PassThrough } = require("stream");
const { baseWorkbookBuffer, readPart, tamper } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

const LIST = { type: "list", allowBlank: true, formulae: ['"a,b"'] };

async function streamBuffer(build) {
    const stream = new PassThrough();
    const chunks = [];
    stream.on("data", (chunk) => chunks.push(chunk));
    const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ stream });
    const ws = wb.addWorksheet("Sheet1");
    build(ws);
    ws.getCell("A1").value = "x";
    await wb.commit();
    return Buffer.concat(chunks);
}

async function load(buffer) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    return wb.getWorksheet("Sheet1");
}

describe("Data validations on whole columns and rows", () => {
    const cases = [
        { ref: "A:A", sqref: "A1:A1048576", inside: ["A1", "A1048576"], outside: ["B1"] },
        { ref: "C:A", sqref: "A1:C1048576", inside: ["B500000"], outside: ["D1"] },
        { ref: "$B:$B", sqref: "B1:B1048576", inside: ["B7"], outside: ["A7"] },
        { ref: "1:1", sqref: "A1:XFD1", inside: ["A1", "XFD1"], outside: ["A2"] },
        { ref: "3:2", sqref: "A2:XFD3", inside: ["Z3"], outside: ["Z4"] },
    ];

    cases.forEach(({ ref, sqref, inside, outside }) => {
        it(`find() applies '${ref}' to its cells`, () => {
            const ws = new ExcelJS.Workbook().addWorksheet("Sheet1");
            ws.dataValidations.add(ref, LIST);
            inside.forEach((address) => expect(ws.dataValidations.find(address), address).to.equal(LIST));
            outside.forEach((address) => expect(ws.dataValidations.find(address), address).to.equal(undefined));
            expect(ws.getCell(inside[0]).dataValidation).to.equal(LIST);
        });

        it(`writes '${ref}' as sqref ${sqref} and reads it back`, async () => {
            const buffer = await baseWorkbookBuffer((wb, ws) => {
                ws.dataValidations.add(ref, LIST);
            });
            const xml = await readPart(buffer, "xl/worksheets/sheet1.xml");
            expect(xml).to.include(`sqref="${sqref}"`);
            const ws = await load(buffer);
            inside.forEach((address) => expect(ws.dataValidations.find(address), address).to.deep.equal(LIST));
            outside.forEach((address) => expect(ws.dataValidations.find(address), address).to.equal(undefined));
        });

        it(`streams '${ref}' as sqref ${sqref}`, async () => {
            const buffer = await streamBuffer((ws) => ws.dataValidations.add(ref, LIST));
            const xml = await readPart(buffer, "xl/worksheets/sheet1.xml");
            expect(xml).to.include(`sqref="${sqref}"`);
        });
    });

    it("removes a whole-column validation added the same way", () => {
        const ws = new ExcelJS.Workbook().addWorksheet("Sheet1");
        ws.dataValidations.add("A:A", LIST);
        ws.dataValidations.remove("A:A");
        expect(ws.dataValidations.find("A9")).to.equal(undefined);
    });

    it("ignores whole-column and whole-row sqrefs outside the grid", async () => {
        const buffer = await tamper(await baseWorkbookBuffer(), {
            "xl/worksheets/sheet1.xml": (xml) =>
                xml.replace(
                    "<pageMargins",
                    '<dataValidations count="1"><dataValidation type="list" sqref="0:0 9999999:9999999 A:XFE B:B">' +
                        '<formula1>"a,b"</formula1></dataValidation></dataValidations><pageMargins'
                ),
        });
        const ws = await load(buffer);
        expect(Object.keys(ws.dataValidations.model)).to.deep.equal(["B1:B1048576"]);
    });

    it("rejects whole-column and whole-row references outside the grid", () => {
        const ws = new ExcelJS.Workbook().addWorksheet("Sheet1");
        ["0:0", "1048577:1048577", "A:XFE"].forEach((ref) => {
            expect(() => ws.dataValidations.add(ref, LIST), ref).to.throw(/Out of bounds/);
        });
    });

    it("reads whole-column and whole-row sqrefs written by other tools", async () => {
        const buffer = await tamper(await baseWorkbookBuffer(), {
            "xl/worksheets/sheet1.xml": (xml) =>
                xml.replace(
                    "<pageMargins",
                    '<dataValidations count="1"><dataValidation type="list" allowBlank="1" sqref="$D:$E 4:10">' +
                        '<formula1>"a,b"</formula1></dataValidation></dataValidations><pageMargins'
                ),
        });
        const ws = await load(buffer);
        expect(ws.dataValidations.find("E1048576")).to.deep.equal(LIST);
        expect(ws.dataValidations.find("XFD4")).to.deep.equal(LIST);
        expect(ws.dataValidations.find("F11")).to.equal(undefined);
        // kept as rectangles: too large to expand per cell
        expect(Object.keys(ws.dataValidations.model).sort()).to.deep.equal(["A4:XFD10", "D1:E1048576"]);
    });
});
