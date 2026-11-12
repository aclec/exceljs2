const { baseWorkbookBuffer, tamper, readPart } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

async function roundTrip(buffer) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    return { wb, buffer: await wb.xlsx.writeBuffer() };
}

describe("Notes VML", () => {
    async function noteWorkbook() {
        return baseWorkbookBuffer((wb, ws) => {
            ws.getCell("A1").note = {
                // two runs: a single plain run is read back as a bare string
                texts: [{ text: "a " }, { text: "note", font: { bold: true } }],
                margins: { insetmode: "custom", inset: [0.25, 0.25, 0.35, 0.35] },
            };
        });
    }

    it("reads a textbox without inset (as Excel writes it) without inventing NaN margins", async () => {
        const buffer = await tamper(await noteWorkbook(), {
            "xl/drawings/vmlDrawing1.vml": (xml) => xml.replace(/ inset="[^"]*"/, ""),
        });
        expect(await readPart(buffer, "xl/drawings/vmlDrawing1.vml")).not.to.match(/ inset=/);

        const { wb, buffer: written } = await roundTrip(buffer);
        const { note } = wb.getWorksheet("Sheet1").getCell("A1");
        expect(note.margins.inset).to.equal(undefined);

        const vml = await readPart(written, "xl/drawings/vmlDrawing1.vml");
        expect(vml).not.to.match(/NaN/);
        expect(vml).to.match(/<v:textbox [^>]*inset="1.3mm,1.3mm,2.5mm,2.5mm"/);
    });

    it("keeps an explicit inset across a round trip", async () => {
        const { wb, buffer } = await roundTrip(await noteWorkbook());
        expect(wb.getWorksheet("Sheet1").getCell("A1").note.margins.inset).to.deep.equal([0.25, 0.25, 0.35, 0.35]);
        const vml = await readPart(buffer, "xl/drawings/vmlDrawing1.vml");
        expect(vml).to.match(/<v:textbox [^>]*inset="2.5mm,2.5mm,3.5mm,3.5mm"/);
    });

    it("never writes non-finite inset values", async () => {
        const buffer = await tamper(await noteWorkbook(), {
            "xl/drawings/vmlDrawing1.vml": (xml) => xml.replace(/ inset="[^"]*"/, ' inset="abc,1mm"'),
        });
        const { wb, buffer: written } = await roundTrip(buffer);
        expect(wb.getWorksheet("Sheet1").getCell("A1").note.margins.inset).to.equal(undefined);
        expect(await readPart(written, "xl/drawings/vmlDrawing1.vml")).not.to.match(/NaN/);

        const wb2 = new ExcelJS.Workbook();
        wb2.addWorksheet("s").getCell("A1").note = { texts: [{ text: "x" }, { text: "y" }], margins: { inset: [NaN, 1, 1, 1] } };
        const vml = await readPart(await wb2.xlsx.writeBuffer(), "xl/drawings/vmlDrawing1.vml");
        expect(vml).not.to.match(/NaN/);
        expect(vml).not.to.match(/ inset=/);
    });
});
