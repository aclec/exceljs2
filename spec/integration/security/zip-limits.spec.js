// Decompression limits of both readers (maxEntrySize, maxUncompressedSize, maxCompressionRatio):
// a zip bomb must be rejected while it is inflated, before its content is materialised.
const fs = require("fs");
const { Readable } = require("stream");
const zlib = require("zlib");

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

// sets a 16-bit field of the central directory entry of `name` (e.g. 10: compression method)
function patchCentralU16(buffer, name, offset, value) {
    const patched = Buffer.from(buffer);
    for (let pos = patched.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06])); pos >= 0; pos--) {
        if (patched.readUInt32LE(pos) === 0x02014b50) {
            const nameLength = patched.readUInt16LE(pos + 28);
            if (patched.toString("utf8", pos + 46, pos + 46 + nameLength) === name) {
                patched.writeUInt16LE(value, pos + offset);
                return patched;
            }
        }
    }
    throw new Error(`no central directory entry ${name}`);
}

function localHeader({ name, data, method, size }) {
    const header = Buffer.alloc(30 + Buffer.byteLength(name));
    header.writeUInt32LE(0x04034b50, 0);
    header.writeUInt16LE(20, 4);
    header.writeUInt16LE(method, 8);
    header.writeUInt32LE(data.length, 18);
    header.writeUInt32LE(size, 22);
    header.writeUInt16LE(Buffer.byteLength(name), 26);
    header.write(name, 30);
    return Buffer.concat([header, data]);
}

// A hand-made archive: the `locals` entries ({name, data (as stored), method, size}) are written in order, then a central
// directory that lists `central` entries ({name, entry, offset}): the local header of `entry` is expected at `offset`,
// so that several entries can share the same data (overlapping zip bombs).
function rawZip(locals, central) {
    const data = Buffer.concat(locals.map(localHeader));
    const headers = central.map(({ name, entry, offset }) => {
        const header = Buffer.alloc(46 + Buffer.byteLength(name));
        header.writeUInt32LE(0x02014b50, 0);
        header.writeUInt16LE(20, 4);
        header.writeUInt16LE(20, 6);
        header.writeUInt16LE(entry.method, 10);
        header.writeUInt32LE(entry.data.length, 20);
        header.writeUInt32LE(entry.size, 24);
        header.writeUInt16LE(Buffer.byteLength(name), 28);
        header.writeUInt32LE(offset, 42);
        header.write(name, 46);
        return header;
    });
    const directory = Buffer.concat(headers);
    const end = Buffer.alloc(22);
    end.writeUInt32LE(0x06054b50, 0);
    end.writeUInt16LE(central.length, 8);
    end.writeUInt16LE(central.length, 10);
    end.writeUInt32LE(directory.length, 12);
    end.writeUInt32LE(data.length, 16);
    return Buffer.concat([data, directory, end]);
}

// complete entries, one after the other
function simpleZip(entries) {
    let offset = 0;
    const central = entries.map((entry) => {
        const item = { name: entry.name, entry, offset };
        offset += localHeader(entry).length;
        return item;
    });
    return rawZip(entries, central);
}

const BACKENDS = ["zlib", "fflate"];

describe("zip decompression limits", function () {
    this.timeout(20000);

    describe("overlapping entries and unread parts (xlsx.load)", () => {
        // 1 MiB that DEFLATE compresses about 190:1
        const content = Buffer.alloc(MiB, " ");
        for (let i = 0; i < content.length; i += 256) {
            content[i] = 97 + ((i * 7919) % 26);
        }
        const image = (name) => ({ name, data: zlib.deflateRawSync(content), method: 8, size: content.length });

        BACKENDS.forEach((backend) => {
            it(`rejects entries that share their data (${backend})`, async () => {
                const shared = image("xl/media/image1.png");
                const central = Array.from({ length: 50 }, (_, i) => ({ name: `xl/media/image${i + 1}.png`, entry: shared, offset: 0 }));
                const error = await rejection(load(rawZip([shared], central), { _backend: backend }));
                expect(error.message).to.equal(
                    'Invalid or unsupported zip file: entries "xl/media/image1.png" and "xl/media/image2.png" overlap'
                );
            });

            it(`rejects an entry whose data contains the local header of another one (${backend})`, async () => {
                // the second local header and its data are stored as the data of the first entry
                const inner = image("xl/media/image2.png");
                const nested = localHeader(inner);
                const outer = { name: "xl/media/image1.png", data: nested, method: 0, size: nested.length };
                const central = [
                    { name: outer.name, entry: outer, offset: 0 },
                    { name: inner.name, entry: inner, offset: 30 + outer.name.length },
                ];
                const error = await rejection(load(rawZip([outer], central), { _backend: backend }));
                expect(error.message).to.equal(
                    'Invalid or unsupported zip file: entries "xl/media/image1.png" and "xl/media/image2.png" overlap'
                );
            });

            it(`bounds the content of all the parts together by maxCompressionRatio (${backend})`, async () => {
                // 12 parts of 1 MiB: each is under the 10 MiB above which a single part is checked
                const buffer = simpleZip(Array.from({ length: 12 }, (_, i) => image(`xl/media/image${i + 1}.png`)));
                expect((await load(buffer, { _backend: backend })).model.media).to.have.length(12);
                const error = await rejection(load(buffer, { _backend: backend, maxCompressionRatio: 100 }));
                expect(error.message).to.equal("Zip content exceeds maxCompressionRatio (100)");
            });

            it(`does not inflate the parts it does not read (${backend})`, async () => {
                const buffer = editParts(await baseWorkbookBuffer(), (parts) => {
                    parts["customXml/item1.xml"] = "<a/>";
                    parts["xl/printerSettings/printerSettings1.bin"] = new Uint8Array(100);
                });
                // an unknown compression method fails when the entry is inflated
                let patched = patchCentralU16(buffer, "customXml/item1.xml", 10, 99);
                patched = patchCentralU16(patched, "xl/printerSettings/printerSettings1.bin", 10, 99);
                const wb = await load(patched, { _backend: backend });
                expect(wb.getWorksheet("Sheet1").getCell("A1").value).to.equal("hello");
                // the parts it reads are inflated
                const error = await rejection(load(patchCentralU16(buffer, "xl/styles.xml", 10, 99), { _backend: backend }));
                expect(error.message).to.equal('Invalid or unsupported zip file: unknown compression method 99 for "xl/styles.xml"');
            });
        });
    });

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

        it("bounds the content of all the parts together by maxCompressionRatio", async () => {
            // 12 worksheets of 1 MiB (each under the 10 MiB above which a single part is checked), compressed ~1000:1
            const parts = unzipParts(await baseWorkbookBuffer());
            const xml = Buffer.from(parts["xl/worksheets/sheet1.xml"]).toString();
            for (let i = 1; i <= 12; i++) {
                parts[`xl/worksheets/sheet${i}.xml`] = Buffer.from(xml.replace("</worksheet>", `${" ".repeat(MiB)}</worksheet>`));
            }
            const buffer = await streamedZip(parts);
            expect(await streamRead(buffer, { maxCompressionRatio: 2000 })).to.equal(12);
            const error = await rejection(streamRead(buffer, { maxCompressionRatio: 100 }));
            expect(error.message).to.equal("Zip content exceeds maxCompressionRatio (100)");
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
