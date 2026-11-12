const { PassThrough } = require("stream");
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

describe("Notes VML shape ids", function () {
    this.timeout(20000);

    // idmap blocks and shape ids of the notes VML of each worksheet
    async function shapeIds(buffer, count) {
        const drawings = [];
        for (let id = 1; id <= count; id++) {
            const vml = await readPart(buffer, `xl/drawings/vmlDrawing${id}.vml`);
            const blocks = /<o:idmap [^>]*data="([^"]*)"/.exec(vml)[1].split(",").map(Number);
            const ids = [...vml.matchAll(/<v:shape [^>]*id="_x0000_s(\d+)"/g)].map((match) => Number(match[1]));
            drawings.push({ blocks, ids });
        }
        return drawings;
    }

    function expectUniqueIdsInClaimedBlocks(drawings) {
        const all = new Set();
        const blocks = new Set();
        drawings.forEach((drawing) => {
            drawing.blocks.forEach((block) => {
                expect(blocks.has(block), `block ${block} claimed twice`).to.equal(false);
                blocks.add(block);
            });
            drawing.ids.forEach((id) => {
                expect(drawing.blocks, `shape ${id}`).to.include(Math.floor(id / 1024));
                // n * 1024 is not a shape id (Excel numbers the shapes of block n from n * 1024 + 1)
                expect(id % 1024, `shape ${id}`).to.not.equal(0);
                expect(all.has(id), `shape ${id} used twice`).to.equal(false);
                all.add(id);
            });
        });
    }

    function addNotes(ws, count) {
        for (let row = 1; row <= count; row++) {
            ws.getCell(`A${row}`).value = row;
            ws.getCell(`A${row}`).note = `note ${row}`;
        }
    }

    it("gives a worksheet with more than 1023 notes several blocks, not those of the next worksheet", async () => {
        const wb = new ExcelJS.Workbook();
        addNotes(wb.addWorksheet("one"), 2100);
        addNotes(wb.addWorksheet("two"), 3);
        const buffer = await wb.xlsx.writeBuffer();
        const drawings = await shapeIds(buffer, 2);
        expect(drawings.map((drawing) => drawing.blocks)).to.deep.equal([[1, 2, 3], [4]]);
        expect(drawings[0].ids).to.have.length(2100);
        expectUniqueIdsInClaimedBlocks(drawings);

        const reloaded = new ExcelJS.Workbook();
        await reloaded.xlsx.load(buffer);
        expect(reloaded.getWorksheet("one").getCell("A2100").note).to.equal("note 2100");
        expect(reloaded.getWorksheet("two").getCell("A3").note).to.equal("note 3");
    });

    it("does the same in the streaming writer", async () => {
        const stream = new PassThrough();
        const chunks = [];
        stream.on("data", (chunk) => chunks.push(chunk));
        const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ stream });
        const one = wb.addWorksheet("one");
        const two = wb.addWorksheet("two");
        addNotes(two, 3); // a later worksheet with notes before the first one is committed
        addNotes(one, 2100);
        one.commit();
        two.commit();
        await wb.commit();
        const buffer = Buffer.concat(chunks);
        const drawings = await shapeIds(buffer, 2);
        expect(drawings[0].ids).to.have.length(2100);
        expect(drawings[0].blocks).to.have.length(3);
        expect(drawings[1].blocks).to.have.length(1);
        expectUniqueIdsInClaimedBlocks(drawings);

        const reloaded = new ExcelJS.Workbook();
        await reloaded.xlsx.load(buffer);
        expect(reloaded.getWorksheet("one").getCell("A2100").note).to.equal("note 2100");
        expect(reloaded.getWorksheet("two").getCell("A3").note).to.equal("note 3");
    });
});
