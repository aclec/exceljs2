// An image `extension` ends up in a zip entry name (xl/media/image1.<extension>) and in the XML of the
// relationships: it must not be able to escape the media folder (zip slip) or the XML attribute.
const { PassThrough } = require("stream");

const SheetRelsWriter = verquire("stream/xlsx/sheet-rels-writer");
const RelType = verquire("xlsx/rel-type");
const { unzipParts } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

const PNG = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==", "base64");

const HOSTILE = [
    'png"/><Relationship Id="rIdX" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://evil.example/" TargetMode="External"/><x a="',
    "png/../../../../../tmp/zipslip",
    "",
    "jpeg2000",
    undefined,
];

describe("image extension", () => {
    it("is checked by the document writer before anything is written", async () => {
        for (const extension of HOSTILE) {
            const wb = new ExcelJS.Workbook();
            const imageId = wb.addImage({ buffer: PNG, extension });
            wb.addWorksheet("Sheet1").addImage(imageId, "A1:B2");
            const output = new PassThrough();
            const written = [];
            output.on("data", (chunk) => written.push(chunk));
            let error;
            try {
                await wb.xlsx.write(output);
            } catch (err) {
                error = err;
            }
            expect(error && error.message, String(extension)).to.equal(`Invalid image extension: ${JSON.stringify(extension)}`);
            expect(written).to.have.length(0);
        }
    });

    it("is checked by the streaming writer", () => {
        const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ stream: new PassThrough() });
        for (const extension of HOSTILE) {
            expect(() => wb.addImage({ buffer: PNG, extension }), String(extension)).to.throw(
                `Invalid image extension: ${JSON.stringify(extension)}`
            );
        }
    });

    it("accepts the usual extensions", async () => {
        for (const extension of ["png", "jpeg", "JPG", "gif", "emf", "tiff"]) {
            const wb = new ExcelJS.Workbook();
            wb.addWorksheet("Sheet1").addImage(wb.addImage({ buffer: PNG, extension }), "A1:B2");
            const parts = unzipParts(await wb.xlsx.writeBuffer());
            expect(Object.keys(parts)).to.include(`xl/media/image1.${extension}`);
        }
    });

    it("XML-encodes the relationship targets of the streaming writer", () => {
        const written = [];
        const stream = {
            write: (data) => written.push(String(data)),
            end() {},
        };
        const writer = new SheetRelsWriter({ id: 1, workbook: { _openStream: () => stream } });
        writer.addMedia({ Type: RelType.Image, Target: '../media/a"b<c&d.png' });
        writer.commit();
        expect(written.join("")).to.include('Target="../media/a&quot;b&lt;c&amp;d.png"');
    });
});
