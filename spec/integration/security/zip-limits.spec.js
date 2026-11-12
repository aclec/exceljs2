// Decompression limits of both readers (maxEntrySize, maxUncompressedSize, maxCompressionRatio):
// a zip bomb must be rejected while it is inflated, before its content is materialised.
const fs = require("fs");
const { Readable } = require("stream");

const { baseWorkbookBuffer, editParts, unzipParts } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");
const { ZipWriter } = verquire("utils/zip-stream");

const MiB = 1024 * 1024;
const isWorksheet = (name) => /^xl\/worksheets\/sheet\d+\.xml$/.test(name);

async function load(buffer, options) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer, options);
    return wb;
}

async function streamRead(buffer, options) {
    const reader = new ExcelJS.stream.xlsx.WorkbookReader(Readable.from([buffer]), options);
    let rows = 0;
    for await (const ws of reader) {
        // eslint-disable-next-line no-unused-vars
        for await (const row of ws) {
            rows++;
        }
    }
    return rows;
}

async function rejection(promise) {
    try {
        await promise;
    } catch (error) {
        return error;
    }
    throw new Error("expected a rejection");
}

// a workbook whose first worksheet ends with `padding` bytes of a single repeated character, which
// DEFLATE compresses about 1000:1
async function bombBuffer(padding) {
    return editParts(await baseWorkbookBuffer(), (parts) => {
        const xml = Buffer.from(parts["xl/worksheets/sheet1.xml"]).toString();
        parts["xl/worksheets/sheet1.xml"] = xml.replace("</worksheet>", `${" ".repeat(padding)}</worksheet>`);
    });
}

// the parts as streamed entries: sizes unknown in the local headers (data descriptors), so the
// readers can only count the bytes they inflate
async function streamedZip(parts) {
    const zip = new ZipWriter();
    const chunks = [];
    zip.on("data", (chunk) => chunks.push(chunk));
    for (const [name, data] of Object.entries(parts)) {
        zip.openStream(name).end(data);
    }
    await zip.finalize();
    return Buffer.concat(chunks);
}

// sets the uncompressed size recorded in the central directory for `name`
function patchDeclaredSize(buffer, name, size) {
    const patched = Buffer.from(buffer);
    for (let pos = patched.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06])); pos >= 0; pos--) {
        if (patched.readUInt32LE(pos) === 0x02014b50) {
            const nameLength = patched.readUInt16LE(pos + 28);
            if (patched.toString("utf8", pos + 46, pos + 46 + nameLength) === name) {
                patched.writeUInt32LE(size, pos + 24);
                return patched;
            }
        }
    }
    throw new Error(`no central directory entry ${name}`);
}

describe("zip decompression limits", function () {
    this.timeout(20000);

    describe("document reader (xlsx.load)", () => {
        it("reads a normal workbook with the default limits", async () => {
            const wb = await load(await baseWorkbookBuffer());
            expect(wb.getWorksheet("Sheet1").getCell("A1").value).to.equal("hello");
        });

        it("rejects an entry larger than maxEntrySize", async () => {
            const error = await rejection(load(await baseWorkbookBuffer(), { maxEntrySize: 500 }));
            expect(error.message).to.match(/^Zip entry ".+" is larger than maxEntrySize \(500 bytes\)$/);
        });

        it("rejects content larger than maxUncompressedSize", async () => {
            const error = await rejection(load(await baseWorkbookBuffer(), { maxUncompressedSize: 5000 }));
            expect(error.message).to.equal("Zip content is larger than maxUncompressedSize (5000 bytes)");
        });

        it("rejects a highly compressed part before inflating it (maxCompressionRatio)", async () => {
            const buffer = await bombBuffer(64 * MiB);
            expect(buffer.length).to.be.below(MiB);
            const error = await rejection(load(buffer));
            expect(error.message).to.equal('Zip entry "xl/worksheets/sheet1.xml" exceeds maxCompressionRatio (1000)');
        });

        it("stops inflating an entry that is larger than its declared size", async () => {
            const buffer = patchDeclaredSize(await bombBuffer(64 * MiB), "xl/worksheets/sheet1.xml", 1000);
            const error = await rejection(load(buffer));
            expect(error.message).to.equal('Invalid or unsupported zip file: "xl/worksheets/sheet1.xml" is larger than its declared size');
        });

        it("validates the limit options", async () => {
            const error = await rejection(load(await baseWorkbookBuffer(), { maxEntrySize: -1 }));
            expect(error).to.be.an.instanceOf(TypeError);
        });
    });

    describe("streaming reader (WorkbookReader)", () => {
        it("reads a normal workbook with the default limits", async () => {
            expect(await streamRead(await baseWorkbookBuffer())).to.equal(1);
        });

        it("rejects an entry larger than maxEntrySize", async () => {
            const error = await rejection(streamRead(await baseWorkbookBuffer(), { maxEntrySize: 500 }));
            expect(error.message).to.match(/^Zip entry ".+" is larger than maxEntrySize \(500 bytes\)$/);
        });

        it("rejects a zip bomb from its declared size", async () => {
            const error = await rejection(streamRead(await bombBuffer(64 * MiB), { maxEntrySize: 16 * MiB, maxCompressionRatio: 1e9 }));
            expect(error.message).to.equal('Zip entry "xl/worksheets/sheet1.xml" is larger than maxEntrySize (16777216 bytes)');
        });

        it("rejects a zip bomb while inflating it", async () => {
            const buffer = await streamedZip(unzipParts(await bombBuffer(64 * MiB)));
            // sizes unknown in the local header
            expect(buffer.readUInt16LE(6) & 8).to.equal(8);
            const error = await rejection(streamRead(buffer, { maxEntrySize: 16 * MiB, maxCompressionRatio: 1e9 }));
            expect(error.message).to.equal('Zip entry "xl/worksheets/sheet1.xml" is larger than maxEntrySize (16777216 bytes)');
            const ratioError = await rejection(streamRead(buffer));
            expect(ratioError.message).to.equal('Zip entry "xl/worksheets/sheet1.xml" exceeds maxCompressionRatio (1000)');
        });

        it("bounds the worksheets spooled to temp files, and removes them", async () => {
            // worksheet first: it is spooled to a temp file until the parts it needs are read
            const parts = unzipParts(await bombBuffer(64 * MiB));
            const ordered = {};
            Object.keys(parts)
                .sort((a, b) => isWorksheet(b) - isWorksheet(a))
                .forEach((name) => (ordered[name] = parts[name]));
            const created = [];
            const original = fs.promises.mkdtemp;
            fs.promises.mkdtemp = async (...args) => {
                const dir = await original.apply(fs.promises, args);
                created.push(dir);
                return dir;
            };
            let error;
            try {
                error = await rejection(streamRead(await streamedZip(ordered), { maxUncompressedSize: 8 * MiB, maxCompressionRatio: 1e9 }));
            } finally {
                fs.promises.mkdtemp = original;
            }
            expect(error.message).to.equal("Zip content is larger than maxUncompressedSize (8388608 bytes)");
            expect(created).to.have.length(1);
            expect(fs.existsSync(created[0])).to.equal(false);
        });
    });
});
