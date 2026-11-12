const { readPart } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

describe("github issues", () => {
    it("issue 3070 - comments on several sheets get distinct VML shape ids", async () => {
        const wb = new ExcelJS.Workbook();
        ["One", "Two", "Three"].forEach((name) => {
            const ws = wb.addWorksheet(name);
            ["A1", "B2"].forEach((address) => {
                ws.getCell(address).value = address;
                ws.getCell(address).note = `${name} ${address}`;
            });
        });
        const buffer = await wb.xlsx.writeBuffer();

        const idmaps = [];
        const shapeIds = [];
        for (const id of [1, 2, 3]) {
            const vml = await readPart(buffer, `xl/drawings/vmlDrawing${id}.vml`);
            idmaps.push(vml.match(/<o:idmap [^>]*data="([^"]*)"/)[1]);
            const ids = vml.match(/<v:shape id="_x0000_s\d+"/g).map((m) => Number(m.match(/s(\d+)"/)[1]));
            // shape ids of a drawing must fall inside the 1024-id blocks its idmap claims
            ids.forEach((shapeId) => expect(Math.floor(shapeId / 1024)).to.equal(id));
            shapeIds.push(...ids);
        }
        expect(idmaps).to.deep.equal(["1", "2", "3"]);
        expect(new Set(shapeIds).size).to.equal(6);

        const wb2 = new ExcelJS.Workbook();
        await wb2.xlsx.load(buffer);
        expect(wb2.getWorksheet("One").getCell("A1").note).to.equal("One A1");
        expect(wb2.getWorksheet("Three").getCell("B2").note).to.equal("Three B2");
    });
});
