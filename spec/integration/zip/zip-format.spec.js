// Zip container details of the writers and of the streaming reader.
const { PassThrough, Readable, Stream, Writable } = require("stream");

const { ZipWriter } = verquire("utils/zip-stream");
const { parseZipStream, readZip } = verquire("utils/zip-reader");
const { readPart, unzipParts, zipParts } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

// compression method of every local header, by entry name
function methods(buffer) {
    const result = {};
    for (let pos = 0; pos + 30 <= buffer.length && buffer.readUInt32LE(pos) === 0x04034b50;) {
        const nameLength = buffer.readUInt16LE(pos + 26);
        const name = buffer.toString("utf8", pos + 30, pos + 30 + nameLength);
        result[name] = buffer.readUInt16LE(pos + 8);
        const flags = buffer.readUInt16LE(pos + 6);
        expect(flags & 8, `${name} has its sizes in the local header`).to.equal(0);
        pos += 30 + nameLength + buffer.readUInt16LE(pos + 28) + buffer.readUInt32LE(pos + 18);
    }
    return result;
}

async function streamToBuffer(build, options) {
    const stream = new PassThrough();
    const chunks = [];
    stream.on("data", (chunk) => chunks.push(chunk));
    const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ stream, ...options });
    build(wb);
    await wb.commit();
    return Buffer.concat(chunks);
}

function documentWorkbook() {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Sheet1");
    for (let i = 1; i <= 100; i++) {
        ws.addRow([i, `row ${i}`, "the same text on every row"]);
    }
    return wb;
}

async function rejection(promise) {
    try {
        await promise;
    } catch (error) {
        return error;
    }
    throw new Error("expected a rejection");
}

describe("Zip format", () => {
    describe("zip options", () => {
        const sheet = "xl/worksheets/sheet1.xml";

        it("deflates at level 6 by default", async () => {
            const buffer = await documentWorkbook().xlsx.writeBuffer();
            expect(methods(buffer)[sheet]).to.equal(8);
        });

        it("accepts {level}, and the archiver and JSZip option shapes", async () => {
            const size = async (zip) => (await documentWorkbook().xlsx.writeBuffer({ zip })).length;
            const stored = [{ level: 0 }, { store: true }, { zlib: { level: 0 } }, { compression: "STORE" }];
            for (const zip of stored) {
                const buffer = await documentWorkbook().xlsx.writeBuffer({ zip });
                expect(methods(buffer)[sheet], JSON.stringify(zip)).to.equal(0);
            }
            expect(await size({ level: 9 })).to.equal(await size({ zlib: { level: 9 } }));
            expect(await size({ level: 9 })).to.equal(await size({ compression: "DEFLATE", compressionOptions: { level: 9 } }));
            expect(await size({ level: 1 })).to.be.above(await size({ level: 9 }));
            expect(await size({ level: 1 })).to.be.below(await size({ level: 0 }));
        });

        it("applies the zip options to the streaming writer", async () => {
            const build = (wb) => wb.addWorksheet("Sheet1").addRow(["a", 1]).commit();
            const stored = await streamToBuffer(build, { zip: { store: true } });
            const local = unzipParts(stored);
            expect(Object.keys(local)).to.include(sheet);
            // the streamed worksheet is written stored (method 0)
            const sheetHeader = stored.indexOf(Buffer.from(sheet)) - 30;
            expect(stored.readUInt16LE(sheetHeader + 8)).to.equal(0);
        });

        it("rejects an invalid level", async () => {
            const error = await rejection(documentWorkbook().xlsx.writeBuffer({ zip: { level: 11 } }));
            expect(error).to.be.an.instanceOf(TypeError);
            expect(() => new ExcelJS.stream.xlsx.WorkbookWriter({ zip: { zlib: { level: -2 } } })).to.throw(TypeError);
        });

        it("writes the archive comment", async () => {
            const buffer = await documentWorkbook().xlsx.writeBuffer({ zip: { comment: "made by ExcelJS" } });
            expect(buffer.subarray(-15).toString()).to.equal("made by ExcelJS");
            expect(buffer.readUInt16LE(buffer.length - 15 - 2)).to.equal(15);
        });
    });

    it("keeps characters whole at any position of a large part", async () => {
        // JSZip encoded strings in 16 KiB slices and broke a surrogate pair cut by a slice boundary;
        // the two paddings put the boundary inside a pair for one of them
        for (const padding of ["", "a"]) {
            const text = `${padding}${"\u{1F600}".repeat(12000)}`;
            const wb = new ExcelJS.Workbook();
            wb.addWorksheet("Sheet1").getCell("A1").value = text;
            const buffer = await wb.xlsx.writeBuffer();
            expect(await readPart(buffer, "xl/sharedStrings.xml")).not.to.include("�");
            const wb2 = new ExcelJS.Workbook();
            await wb2.xlsx.load(buffer);
            expect(wb2.getWorksheet("Sheet1").getCell("A1").value).to.equal(text);
        }
    });

    describe("zip writer", () => {
        async function collect(zip) {
            const chunks = [];
            zip.on("data", (chunk) => chunks.push(chunk));
            return async () => {
                await zip.finalize();
                return Buffer.concat(chunks);
            };
        }
        const contents = async (buffer) => {
            const result = {};
            for (const entry of readZip(buffer)) {
                result[entry.name] = Buffer.from(await entry.data()).toString();
            }
            return result;
        };

        it("copies the caller's buffers when it stores them uncompressed", async () => {
            const zip = new ZipWriter({ level: 0 });
            const done = await collect(zip);
            const first = zip.openStream("first.txt");
            const second = zip.openStream("second.txt");
            const reused = Buffer.from("AAAA");
            // second.txt waits behind first.txt, which is still open
            second.write(reused);
            reused.fill("B");
            second.end(reused);
            const appended = Buffer.from("DDDD");
            zip.append(appended, { name: "third.txt" });
            appended.fill("E");
            first.end("first");
            expect(await contents(await done())).to.deep.equal({ "first.txt": "first", "second.txt": "AAAABBBB", "third.txt": "DDDD" });
        });

        it("still finalizes after an append with invalid data", async () => {
            const zip = new ZipWriter();
            const done = await collect(zip);
            zip.append("ok", { name: "ok.txt" });
            expect(() => zip.append(null, { name: "bad.txt" })).to.throw(TypeError);
            expect(await contents(await done())).to.deep.equal({ "ok.txt": "ok" });
        });

        it("rejects finalize when the destination throws", async () => {
            const zip = new ZipWriter();
            let writes = 0;
            zip.pipe({
                write() {
                    writes++;
                    throw new Error("cannot write");
                },
                end() {},
            });
            const entry = zip.openStream("streamed.txt");
            entry.write("data");
            entry.end("more data");
            // the failure is reported to whoever adds an entry next, and by finalize
            expect(() => zip.append("other", { name: "other.txt" })).to.throw("cannot write");
            const error = await rejection(zip.finalize());
            expect(error.message).to.equal("cannot write");
            // nothing more is written to a failed destination
            expect(writes).to.equal(1);
        });

        it("writes the central directory once when an entry is ended from a 'zipped' listener", async () => {
            const zip = new ZipWriter();
            let ends = 0;
            const chunks = [];
            zip.pipe({ write: (chunk) => chunks.push(chunk), end: () => ends++ });
            const first = zip.openStream("first.txt");
            const second = zip.openStream("second.txt");
            first.on("zipped", () => second.end("second"));
            const finalized = zip.finalize();
            first.end("first");
            await finalized;
            const buffer = Buffer.concat(chunks);
            expect(ends).to.equal(1);
            expect(buffer.indexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]))).to.equal(buffer.length - 22);
            expect(await contents(buffer)).to.deep.equal({ "first.txt": "first", "second.txt": "second" });
        });
    });

    it("returns stored data that does not share the input buffer", async () => {
        const input = await documentWorkbook().xlsx.writeBuffer({ zip: { level: 0 } });
        const [entry] = readZip(input);
        const data = await entry.data();
        const copy = Buffer.from(data);
        input.fill(0);
        expect(Buffer.from(data).equals(copy)).to.equal(true);
    });

    describe("streaming reader", () => {
        async function readEntries(buffer, chunkSize) {
            const chunks = [];
            for (let i = 0; i < buffer.length; i += chunkSize) {
                chunks.push(buffer.subarray(i, i + chunkSize));
            }
            const result = {};
            for await (const entry of parseZipStream(Readable.from(chunks))) {
                const parts = [];
                for await (const chunk of entry) {
                    parts.push(chunk);
                }
                result[entry.path] = Buffer.concat(parts);
            }
            return result;
        }

        it("finds the end of streamed entries whose data contains zip signatures", async () => {
            // stored data with a data descriptor: the signatures below are in the entry data itself
            const tricky = Buffer.concat([
                Buffer.from("before "),
                Buffer.from([0x50, 0x4b, 0x07, 0x08, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]),
                Buffer.from([0x50, 0x4b, 0x03, 0x04]),
                Buffer.from([0x50, 0x4b, 0x01, 0x02]),
                Buffer.from(" after"),
            ]);
            for (const level of [0, 6]) {
                const zip = new ZipWriter({ level });
                const out = [];
                zip.on("data", (chunk) => out.push(chunk));
                const entry = zip.openStream("data.bin");
                for (let i = 0; i < 50; i++) {
                    entry.write(tricky);
                }
                const second = zip.openStream("second.txt");
                second.end("second");
                entry.end();
                await zip.finalize();
                const buffer = Buffer.concat(out);
                for (const chunkSize of [3, 17, 4096]) {
                    const entries = await readEntries(buffer, chunkSize);
                    expect(entries["data.bin"].equals(Buffer.concat(Array(50).fill(tricky))), `level ${level}`).to.equal(true);
                    expect(entries["second.txt"].toString()).to.equal("second");
                }
            }
        });

        it("skips the rest of an entry that the consumer stopped reading", async () => {
            const big = "x".repeat(200000);
            for (const streamed of [false, true]) {
                for (const level of [0, 6]) {
                    const zip = new ZipWriter({ level });
                    const out = [];
                    zip.on("data", (chunk) => out.push(chunk));
                    if (streamed) {
                        zip.openStream("big.txt").end(big);
                        zip.openStream("next.txt").end("next");
                    } else {
                        zip.append(big, { name: "big.txt" });
                        zip.append("next", { name: "next.txt" });
                    }
                    await zip.finalize();
                    const read = {};
                    for await (const entry of parseZipStream(Readable.from([Buffer.concat(out)]))) {
                        const parts = [];
                        for await (const chunk of entry) {
                            parts.push(chunk);
                            break;
                        }
                        read[entry.path] = Buffer.concat(parts).toString();
                    }
                    expect(read["next.txt"], `streamed: ${streamed}, level ${level}`).to.equal("next");
                    expect(big.startsWith(read["big.txt"])).to.equal(true);
                }
            }
        });

        it("reads the ZIP64 sizes of a local header where only one size is 0xFFFFFFFF", async () => {
            const zip = new ZipWriter({ forceZip64: true });
            const out = [];
            zip.on("data", (chunk) => out.push(chunk));
            zip.append("x".repeat(5000), { name: "a.txt" });
            zip.append("b", { name: "b.txt" });
            await zip.finalize();
            const buffer = Buffer.concat(out);
            // the field still holds both sizes (uncompressed first), but only the compressed one is flagged
            expect(buffer.readUInt32LE(22)).to.equal(0xffffffff);
            buffer.writeUInt32LE(5000, 22);
            const read = {};
            for await (const entry of parseZipStream(Readable.from([buffer]))) {
                const parts = [];
                for await (const chunk of entry) {
                    parts.push(chunk);
                }
                read[entry.path] = Buffer.concat(parts).toString();
            }
            expect(read).to.deep.equal({ "a.txt": "x".repeat(5000), "b.txt": "b" });
        });

        it("reads the worksheets after a worksheet the consumer stopped reading (Excel part order)", async () => {
            const wb = new ExcelJS.Workbook();
            for (const name of ["One", "Two"]) {
                const ws = wb.addWorksheet(name);
                for (let i = 1; i <= 3000; i++) {
                    ws.addRow([name, i]);
                }
            }
            // worksheets last, as Excel writes them: they are not spooled
            const parts = unzipParts(await wb.xlsx.writeBuffer());
            const isSheet = (name) => name.startsWith("xl/worksheets/");
            const ordered = {};
            Object.keys(parts)
                .sort((a, b) => isSheet(a) - isSheet(b))
                .forEach((name) => (ordered[name] = parts[name]));
            const counts = {};
            for await (const ws of new ExcelJS.stream.xlsx.WorkbookReader(Readable.from([zipParts(ordered)]))) {
                counts[ws.name] = 0;
                for await (const row of ws) {
                    counts[ws.name]++;
                    if (ws.name === "One" && row.number === 1) break;
                }
            }
            expect(counts).to.deep.equal({ One: 1, Two: 3000 });
        });

        it("skips the entries that are not read", async () => {
            const buffer = await documentWorkbook().xlsx.writeBuffer();
            const names = [];
            for await (const entry of parseZipStream(Readable.from([buffer]))) {
                names.push(entry.path);
            }
            expect(names).to.include("xl/worksheets/sheet1.xml");
            expect(names).to.include("xl/workbook.xml");
        });

        it("rejects a truncated archive", async () => {
            const buffer = await documentWorkbook().xlsx.writeBuffer();
            const error = await rejection(readEntries(buffer.subarray(0, buffer.length - 2000), 1000));
            expect(error.message).to.equal("Unexpected end of zip data");
        });
    });

    describe("WorkbookReader input", () => {
        let buffer;
        before(async () => {
            buffer = await documentWorkbook().xlsx.writeBuffer();
        });

        async function rowCount(input) {
            let count = 0;
            for await (const ws of new ExcelJS.stream.xlsx.WorkbookReader(input)) {
                // eslint-disable-next-line no-unused-vars
                for await (const row of ws) {
                    count++;
                }
            }
            return count;
        }

        it("reads a stream that is not a stream.Readable (older stream implementations)", async () => {
            // a legacy stream: pipe() and 'data' / 'end' events only, no async iterator
            const legacy = new Stream();
            legacy.readable = true;
            setImmediate(() => {
                legacy.emit("data", buffer.subarray(0, 1000));
                legacy.emit("data", buffer.subarray(1000));
                legacy.emit("end");
            });
            expect(legacy instanceof Readable).to.equal(false);
            expect(await rowCount(legacy)).to.equal(100);
        });

        it("reads a web ReadableStream", async () => {
            const web = new ReadableStream({
                start(controller) {
                    controller.enqueue(new Uint8Array(buffer.subarray(0, 500)));
                    controller.enqueue(new Uint8Array(buffer.subarray(500)));
                    controller.close();
                },
            });
            expect(await rowCount(web)).to.equal(100);
        });

        it("reads an async iterable of chunks", async () => {
            async function* chunks() {
                for (let i = 0; i < buffer.length; i += 333) {
                    yield buffer.subarray(i, i + 333);
                }
            }
            expect(await rowCount(chunks())).to.equal(100);
        });

        it("still refuses what it cannot read", async () => {
            const error = await rejection(rowCount({}));
            expect(error.message).to.match(/Could not recognise input/);
        });
    });

    describe("write errors", () => {
        let unhandled;
        const onUnhandled = (reason) => unhandled.push(reason);
        beforeEach(() => {
            unhandled = [];
            process.on("unhandledRejection", onUnhandled);
        });
        afterEach(async () => {
            await new Promise((resolve) => setImmediate(resolve));
            process.removeListener("unhandledRejection", onUnhandled);
            expect(unhandled).to.deep.equal([]);
        });

        const failing = () =>
            new Writable({
                write(chunk, encoding, callback) {
                    callback(new Error("disk full"));
                },
            });

        it("rejects xlsx.write when the output stream fails", async () => {
            const error = await rejection(documentWorkbook().xlsx.write(failing()));
            expect(error.message).to.equal("disk full");
        });

        it("rejects xlsx.write when the output reports an error while it is written", async () => {
            const { EventEmitter } = require("events");
            const output = new EventEmitter();
            output.write = () => output.emit("error", new Error("connection reset"));
            output.end = () => {};
            const error = await rejection(documentWorkbook().xlsx.write(output));
            expect(error.message).to.equal("connection reset");
            // the listener is removed once the archive has failed or ended
            expect(output.listenerCount("error")).to.be.at.most(1);
        });

        it("rejects WorkbookWriter.commit when the output stream fails", async () => {
            const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ stream: failing() });
            wb.addWorksheet("Sheet1").addRow([1]).commit();
            const error = await rejection(wb.commit());
            expect(error.message).to.equal("disk full");
        });

        it("rejects WorkbookWriter.commit when an image file cannot be read", async () => {
            const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ stream: new PassThrough() });
            const imageId = wb.addImage({ filename: "./spec/out/no-such-image.png", extension: "png" });
            wb.addWorksheet("Sheet1").addBackgroundImage(imageId);
            const error = await rejection(wb.commit());
            expect(error.code).to.equal("ENOENT");
        });
    });
});
