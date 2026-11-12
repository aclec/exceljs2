// xlsx.load keys the parts it reads by their zip entry names (worksheets, media, drawings, comments, tables, VML,
// themes): names from the file, so the hashes must have no prototype (CLAUDE.md invariant), whatever the names.
const { baseWorkbookBuffer, editParts } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");
const XLSX = verquire("xlsx/xlsx");

const PART_HASHES = ["worksheetHash", "mediaIndex", "drawings", "drawingRels", "comments", "tables", "vmlDrawings", "themes"];

// the read model of xlsx.load, as handed to reconcile
async function loadModel(buffer) {
    const reconcileSteps = XLSX.prototype._reconcileSteps;
    let captured;
    XLSX.prototype._reconcileSteps = function (model, options) {
        captured = { ...model };
        return reconcileSteps.call(this, model, options);
    };
    try {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buffer);
        return { wb, model: captured };
    } finally {
        XLSX.prototype._reconcileSteps = reconcileSteps;
    }
}

describe("xlsx part hashes keyed by zip entry names", () => {
    let buffer;
    before(async () => {
        const base = await baseWorkbookBuffer((wb, ws) => {
            ws.getCell("B2").note = "a note";
            ws.addTable({ name: "T", ref: "D1", columns: [{ name: "x" }], rows: [[1]] });
            const imageId = wb.addImage({ filename: "./spec/integration/data/image.png", extension: "png" });
            ws.addImage(imageId, "F2:G4");
        });
        // the same parts again under the names of Object.prototype members
        buffer = editParts(base, (parts) => {
            parts["xl/theme/constructor.xml"] = parts["xl/theme/theme1.xml"];
            parts["xl/theme/toString.xml"] = parts["xl/theme/theme1.xml"];
            parts["xl/drawings/constructor.xml"] = parts["xl/drawings/drawing1.xml"];
            parts["xl/drawings/_rels/constructor.xml.rels"] = parts["xl/drawings/_rels/drawing1.xml.rels"];
            const media = Object.keys(parts).find((name) => name.startsWith("xl/media/"));
            parts["xl/media/constructor.png"] = parts[media];
        });
    });

    it("creates them without a prototype", async () => {
        const { model } = await loadModel(buffer);
        for (const name of PART_HASHES) {
            expect(Object.getPrototypeOf(model[name]), name).to.equal(null);
        }
        expect(model.themes.constructor).to.be.a("string");
        expect(model.themes.toString).to.be.a("string");
        expect(model.drawings.constructor.anchors).to.be.an("array");
        expect(model.mediaIndex.constructor).to.be.a("number");
        expect(model.mediaIndex["constructor.png"]).to.be.a("number");
    });

    it("loads and writes back a workbook with such parts", async () => {
        const { wb } = await loadModel(buffer);
        const ws = wb.getWorksheet("Sheet1");
        expect(ws.getCell("A1").value).to.equal("hello");
        expect(ws.getCell("B2").note).to.equal("a note");
        expect(ws.getImages()).to.have.length(1);

        const back = new ExcelJS.Workbook();
        await back.xlsx.load(await wb.xlsx.writeBuffer());
        expect(back.getWorksheet("Sheet1").getCell("A1").value).to.equal("hello");
        expect({}.constructor).to.equal(Object);
    });
});
