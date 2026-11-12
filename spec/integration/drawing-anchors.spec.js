const fs = require("fs");
const path = require("path");
const { tamper } = require("../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

const DRAWING = "xl/drawings/drawing1.xml";
const png = fs.readFileSync(path.join(__dirname, "data/image.png"));

// a two-cell anchor (B2:D4), then a one-cell anchor (B7, 50 x 40 px)
async function drawingBuffer(patch) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("s");
    const imageId = wb.addImage({ buffer: png, extension: "png" });
    ws.addImage(imageId, "B2:D4");
    ws.addImage(imageId, { tl: { col: 1, row: 6 }, ext: { width: 50, height: 40 } });
    const buffer = await wb.xlsx.writeBuffer();
    return patch ? tamper(buffer, { [DRAWING]: patch }) : buffer;
}

async function images(buffer) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    // the workbook loaded can be written
    await wb.xlsx.writeBuffer();
    return wb.getWorksheet("s").getImages();
}

describe("Drawing anchors read from a file", () => {
    it("reads both anchors of a valid drawing", async () => {
        const [twoCell, oneCell] = await images(await drawingBuffer());
        expect(twoCell.range.br.nativeCol).to.equal(4);
        expect(oneCell.range.ext).to.deep.equal({ width: 50, height: 40 });
    });

    it("skips an anchor without one of its required elements (<xdr:from>, <xdr:to>, <xdr:ext>)", async () => {
        const withoutFrom = await images(await drawingBuffer((xml) => xml.replace(/<xdr:from>[\s\S]*?<\/xdr:from>/, "")));
        expect(withoutFrom.map((image) => image.range.tl.nativeRow)).to.deep.equal([6]);

        const withoutTo = await images(await drawingBuffer((xml) => xml.replace(/<xdr:to>[\s\S]*?<\/xdr:to>/, "")));
        expect(withoutTo.map((image) => image.range.tl.nativeRow)).to.deep.equal([6]);

        const withoutExt = await images(await drawingBuffer((xml) => xml.replace(/<xdr:ext [^>]*\/>/, "")));
        expect(withoutExt.map((image) => image.range.tl.nativeRow)).to.deep.equal([1]);
    });

    it("reads a missing or invalid value of a corner as 0", async () => {
        const [withoutColOff] = await images(await drawingBuffer((xml) => xml.replace(/<xdr:colOff>[^<]*<\/xdr:colOff>/, "")));
        expect(withoutColOff.range.tl.nativeColOff).to.equal(0);

        const [invalidCol] = await images(await drawingBuffer((xml) => xml.replace(/<xdr:col>[^<]*<\/xdr:col>/, "<xdr:col>x</xdr:col>")));
        expect(invalidCol.range.tl.nativeCol).to.equal(0);
    });

    it("reads an invalid editAs as oneCell", async () => {
        const [image] = await images(await drawingBuffer((xml) => xml.replace(/editAs="oneCell"/, 'editAs="bogus"')));
        expect(image.range.editAs).to.equal("oneCell");
    });

    it("throws a clear Error when writing an image anchored by one cell without its size", async () => {
        const wb = new ExcelJS.Workbook();
        const imageId = wb.addImage({ buffer: png, extension: "png" });
        wb.addWorksheet("s").addImage(imageId, { tl: { col: 1, row: 1 } });
        const error = await wb.xlsx.writeBuffer().then(
            () => undefined,
            (e) => e
        );
        expect(error).to.be.an.instanceOf(Error);
        expect(error.message).to.equal("An image anchored by one cell needs its size (range.ext: width and height) and its picture");
    });
});
