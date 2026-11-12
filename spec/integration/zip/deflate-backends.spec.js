// The two DEFLATE backends of the zip code (lib/utils/deflate.js): node:zlib, used on Node.js and Bun,
// and fflate, used in the browser bundle (spec/browser runs it there). The internal `_backend` option
// selects one; each must read what the other writes.
const { Readable } = require("stream");
const { unzipSync } = require("fflate");

const deflate = verquire("utils/deflate");
const { ZipWriter } = verquire("utils/zip-stream");
const { readZip, parseZipStream } = verquire("utils/zip-reader");

const ExcelJS = verquire("exceljs");

const BACKENDS = ["zlib", "fflate"];

// deterministic, somewhat compressible data
function sample(size, seed) {
    const data = Buffer.alloc(size);
    let x = seed;
    for (let i = 0; i < size; i++) {
        // xorshift32
        x ^= x << 13;
        x ^= x >>> 17;
        x ^= x << 5;
        data[i] = 97 + ((x >>> 0) % 16);
    }
    return data;
}

const FILES = {
    "small.txt": Buffer.from("hello"),
    "empty.txt": Buffer.alloc(0),
    // more than the 1 MiB segments of the zlib compressor
    "large.xml": sample(2.5 * 1024 * 1024, 1),
    "streamed.xml": sample(1.5 * 1024 * 1024, 2),
};

async function write(backend, level) {
    const zip = new ZipWriter({ level, _backend: backend });
    const chunks = [];
    zip.on("data", (chunk) => chunks.push(chunk));
    // a streamed entry, written in small pieces while complete entries queue behind it
    const streamed = zip.openStream("streamed.xml");
    for (const name of ["small.txt", "empty.txt", "large.xml"]) {
        zip.append(FILES[name], { name });
    }
    const data = FILES["streamed.xml"];
    for (let i = 0; i < data.length; i += 1000) {
        streamed.write(data.subarray(i, i + 1000));
    }
    streamed.end();
    await zip.finalize();
    return Buffer.concat(chunks);
}

async function readDocument(buffer, backend) {
    const result = {};
    for (const entry of readZip(buffer, { _backend: backend })) {
        result[entry.name] = Buffer.from(await entry.data());
    }
    return result;
}

async function readStream(buffer, backend) {
    const result = {};
    const chunks = [];
    for (let i = 0; i < buffer.length; i += 50000) {
        chunks.push(buffer.subarray(i, i + 50000));
    }
    for await (const entry of parseZipStream(Readable.from(chunks), { _backend: backend })) {
        const parts = [];
        for await (const chunk of entry) {
            parts.push(chunk);
        }
        result[entry.path] = Buffer.concat(parts);
    }
    return result;
}

function expectFiles(actual, label) {
    expect(Object.keys(actual).sort(), label).to.deep.equal(Object.keys(FILES).sort());
    for (const [name, data] of Object.entries(FILES)) {
        expect(actual[name].equals(data), `${label}: ${name}`).to.equal(true);
    }
}

describe("Deflate backends", function () {
    this.timeout(60000);

    it("uses node:zlib on Node.js and Bun", () => {
        expect(deflate.NATIVE).to.equal(true);
        expect(new ZipWriter().backend).to.equal("zlib");
        expect(() => new ZipWriter({ _backend: "other" })).to.throw(TypeError);
    });

    for (const level of [0, 1, 6, 9]) {
        it(`reads with each backend what the other writes (level ${level})`, async () => {
            for (const writer of BACKENDS) {
                const buffer = await write(writer, level);
                const parts = unzipSync(new Uint8Array(buffer));
                const reference = {};
                Object.entries(parts).forEach(([name, data]) => (reference[name] = Buffer.from(data)));
                expectFiles(reference, `${writer} -> fflate.unzipSync`);
                for (const reader of BACKENDS) {
                    expectFiles(await readDocument(buffer, reader), `${writer} -> readZip(${reader})`);
                    expectFiles(await readStream(buffer, reader), `${writer} -> parseZipStream(${reader})`);
                }
            }
        });
    }

    it("round-trips workbooks across backends, both writers", async () => {
        const build = (wb) => {
            const ws = wb.addWorksheet("data");
            for (let i = 1; i <= 3000; i++) {
                ws.addRow([i, `text ${i}`, i / 7]);
            }
        };
        for (const writer of BACKENDS) {
            const wb = new ExcelJS.Workbook();
            build(wb);
            const document = await wb.xlsx.writeBuffer({ zip: { _backend: writer } });

            const chunks = [];
            const stream = new (require("stream").PassThrough)();
            stream.on("data", (chunk) => chunks.push(chunk));
            const swb = new ExcelJS.stream.xlsx.WorkbookWriter({ stream, zip: { _backend: writer } });
            const sws = swb.addWorksheet("data");
            for (let i = 1; i <= 3000; i++) {
                sws.addRow([i, `text ${i}`, i / 7]).commit();
            }
            await swb.commit();
            const streamed = Buffer.concat(chunks);

            for (const reader of BACKENDS) {
                for (const buffer of [document, streamed]) {
                    const loaded = new ExcelJS.Workbook();
                    await loaded.xlsx.load(buffer, { _backend: reader });
                    expect(loaded.getWorksheet("data").getRow(3000).values).to.deep.equal([undefined, 3000, "text 3000", 3000 / 7]);

                    let rows = 0;
                    for await (const ws of new ExcelJS.stream.xlsx.WorkbookReader(Readable.from([buffer]), { _backend: reader })) {
                        for await (const row of ws) {
                            rows = row.values[1];
                        }
                    }
                    expect(rows, `${writer} -> ${reader}`).to.equal(3000);
                }
            }
        }
    });

    describe("limits and errors with each backend", () => {
        // 64 MiB of a single byte: about 64 KiB compressed
        let bomb;
        before(async () => {
            const zip = new ZipWriter();
            const chunks = [];
            zip.on("data", (chunk) => chunks.push(chunk));
            const entry = zip.openStream("bomb.xml");
            const block = Buffer.alloc(1024 * 1024, 32);
            for (let i = 0; i < 64; i++) {
                entry.write(block);
            }
            entry.end();
            await zip.finalize();
            bomb = Buffer.concat(chunks);
        });

        for (const backend of BACKENDS) {
            it(`stops a zip bomb while inflating it (${backend})`, async () => {
                const options = { _backend: backend, maxEntrySize: 16 * 1024 * 1024, maxCompressionRatio: 1e9 };
                let error;
                try {
                    for await (const entry of parseZipStream(Readable.from([bomb]), options)) {
                        // eslint-disable-next-line no-unused-vars
                        for await (const chunk of entry) {
                            // inflate
                        }
                    }
                } catch (err) {
                    error = err;
                }
                expect(error && error.message).to.equal('Zip entry "bomb.xml" is larger than maxEntrySize (16777216 bytes)');

                error = undefined;
                try {
                    for await (const entry of parseZipStream(Readable.from([bomb]), { _backend: backend })) {
                        // eslint-disable-next-line no-unused-vars
                        for await (const chunk of entry) {
                            // inflate
                        }
                    }
                } catch (err) {
                    error = err;
                }
                expect(error && error.message).to.equal('Zip entry "bomb.xml" exceeds maxCompressionRatio (1000)');
            });

            it(`rejects corrupt and truncated deflate data (${backend})`, async () => {
                const buffer = await write("zlib", 6);
                const corrupt = Buffer.from(buffer);
                // inside the compressed data of the first (streamed) entry
                for (let i = 2000; i < 2400; i++) {
                    corrupt[i] ^= 0x55;
                }
                for (const input of [corrupt, buffer.subarray(0, Math.floor(buffer.length / 2))]) {
                    let error;
                    try {
                        await readStream(input, backend);
                    } catch (err) {
                        error = err;
                    }
                    expect(error).to.be.an.instanceOf(Error);
                }
                let error;
                try {
                    await readDocument(corrupt, backend);
                } catch (err) {
                    error = err;
                }
                expect(error).to.be.an.instanceOf(Error);
            });
        }
    });
});
