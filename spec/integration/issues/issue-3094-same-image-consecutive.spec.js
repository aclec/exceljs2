const { readPart } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

// two distinct 1x1 PNGs (red, blue)
const RED = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==", "base64");
const BLUE = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPj/HwADBwIAMCbHYQAAAABJRU5ErkJggg==", "base64");

// media file name of every picture of a drawing, in anchor order
async function pictureTargets(buffer, drawing) {
    const xml = await readPart(buffer, `xl/drawings/${drawing}.xml`);
    const rels = await readPart(buffer, `xl/drawings/_rels/${drawing}.xml.rels`);
    const targets = {};
    rels.replace(/<Relationship ([^>]*)\/>/g, (m, attrs) => {
        targets[attrs.match(/Id="([^"]+)"/)[1]] = attrs.match(/Target="([^"]+)"/)[1];
        return m;
    });
    return xml.match(/r:embed="[^"]+"/g).map((embed) => targets[embed.match(/"([^"]+)"/)[1]]);
}

describe("github issues", () => {
    it("issue 3094 - each picture references its own image, across sheets", async () => {
        const wb = new ExcelJS.Workbook();
        const red = wb.addImage({ buffer: RED, extension: "png" });
        const blue = wb.addImage({ buffer: BLUE, extension: "png" });
        const ws1 = wb.addWorksheet("One");
        [red, red, blue, red, blue, blue].forEach((imageId, i) => {
            ws1.addImage(imageId, { tl: { col: i, row: 0 }, ext: { width: 10, height: 10 } });
        });
        const ws2 = wb.addWorksheet("Two");
        [blue, red, red].forEach((imageId, i) => {
            ws2.addImage(imageId, { tl: { col: i, row: 0 }, ext: { width: 10, height: 10 } });
        });

        const buffer = await wb.xlsx.writeBuffer();
        const R = "../media/image1.png";
        const B = "../media/image2.png";
        expect(await pictureTargets(buffer, "drawing1")).to.deep.equal([R, R, B, R, B, B]);
        expect(await pictureTargets(buffer, "drawing2")).to.deep.equal([B, R, R]);

        const wb2 = new ExcelJS.Workbook();
        await wb2.xlsx.load(buffer);
        const images = (ws) => ws.getImages().map((image) => wb2.getImage(image.imageId).buffer.equals(RED));
        expect(images(wb2.getWorksheet("One"))).to.deep.equal([true, true, false, true, false, false]);
        expect(images(wb2.getWorksheet("Two"))).to.deep.equal([false, true, true]);
    });
});
