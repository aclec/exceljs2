// Streaming writer backpressure: `await row.commit()` / `await worksheet.commit()` wait while the output stream is full or
// too much data is being compressed, so that memory stays bounded however many rows are written.
const fs = require("fs");
const os = require("os");
const path = require("path");
const { Writable } = require("stream");
const v8 = require("v8");
const vm = require("vm");

const ExcelJS = verquire("exceljs");
const { ZipWriter } = verquire("utils/zip-stream");

const ROW = (i) => [i, `name ${i}`, i * 1.5, new Date(2020, 0, 1 + (i % 365)), `some longer text value for row ${i}`];

// a function that runs a full garbage collection (Bun, or Node.js with V8's gc exposed at run time)
function exposeGc() {
    if (typeof Bun !== "undefined") {
        return () => Bun.gc(true); // eslint-disable-line no-undef
    }
    if (typeof global.gc === "function") {
        return global.gc;
    }
    v8.setFlagsFromString("--expose-gc");
    return vm.runInNewContext("gc");
}

// a destination that accepts one chunk per event-loop turn
function slowStream(highWaterMark) {
    const chunks = [];
    const stream = new Writable({
        highWaterMark,
        write(chunk, encoding, callback) {
            chunks.push(chunk);
            setImmediate(callback);
        },
    });
    stream.chunks = chunks;
    return stream;
}

describe("WorkbookWriter (stream) backpressure", () => {
    it("row.commit() and worksheet.commit() wait for a full output stream to drain", async () => {
        const stream = slowStream(16 * 1024);
        const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ stream, zip: { level: 0 } });
        const ws = wb.addWorksheet("data");
        let maxBuffered = 0;
        for (let i = 1; i <= 20000; i++) {
            const pending = ws.addRow(ROW(i)).commit();
            expect(pending).to.be.an.instanceof(Promise);
            await pending;
            maxBuffered = Math.max(maxBuffered, stream.writableLength);
        }
        await ws.commit();
        await wb.commit();
        // settled once the archive is complete
        await ws.commit();

        const written = Buffer.concat(stream.chunks);
        expect(written.length).to.be.above(2 * 1024 * 1024);
        // without waiting, the whole sheet would wait in the stream's buffer
        expect(maxBuffered).to.be.below(256 * 1024);

        const reloaded = new ExcelJS.Workbook();
        await reloaded.xlsx.load(written);
        const sheet = reloaded.getWorksheet("data");
        expect(sheet.rowCount).to.equal(20000);
        expect(sheet.getCell("B20000").value).to.equal("name 20000");
    });

    it("keeps the synchronous API: rows committed without waiting are all written", async () => {
        const stream = slowStream(16 * 1024);
        const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ stream });
        const ws = wb.addWorksheet("data");
        for (let i = 1; i <= 5000; i++) {
            ws.addRow(ROW(i)).commit();
        }
        ws.commit();
        await wb.commit();
        const reloaded = new ExcelJS.Workbook();
        await reloaded.xlsx.load(Buffer.concat(stream.chunks));
        expect(reloaded.getWorksheet("data").rowCount).to.equal(5000);
    });

    it("rejects the pending commits when the output stream fails, without unhandled rejections", async () => {
        const stream = new Writable({
            highWaterMark: 1024,
            write(chunk, encoding, callback) {
                setImmediate(() => callback(new Error("disk full")));
            },
        });
        const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ stream, zip: { level: 0 } });
        const ws = wb.addWorksheet("data");
        let error;
        try {
            for (let i = 1; i <= 20000; i++) {
                ws.addRow(ROW(i)).commit(); // not awaited: must not be an unhandled rejection
                await ws.addRow(ROW(i)).commit();
            }
        } catch (err) {
            error = err;
        }
        expect(error && error.message).to.equal("disk full");
        let commitError;
        await wb.commit().catch((err) => {
            commitError = err;
        });
        expect(commitError && commitError.message).to.equal("disk full");
    });

    it("rejects the pending commits when the output stream is destroyed without an error", async () => {
        const stream = slowStream(1024);
        const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ stream, zip: { level: 0 } });
        const ws = wb.addWorksheet("data");
        let error;
        try {
            for (let i = 1; i <= 20000; i++) {
                if (i === 100) {
                    stream.destroy(); // e.g. an HTTP response whose client went away
                }
                await ws.addRow(ROW(i)).commit();
            }
        } catch (err) {
            error = err;
        }
        expect(error && error.message).to.match(/closed/);
        let commitError;
        await wb.commit().catch((err) => {
            commitError = err;
        });
        expect(commitError).to.equal(error);
    });

    it("settles a pending drained() when the archive completes", async () => {
        const zip = new ZipWriter({ level: 0 });
        zip.pipe(slowStream(16));
        zip.append("x".repeat(1000), { name: "a.txt" });
        const pending = zip.drained();
        await zip.finalize();
        const settled = await Promise.race([pending.then(() => true), new Promise((resolve) => setTimeout(resolve, 500, false))]);
        expect(settled).to.equal(true);
    });

    it("keeps memory bounded when rows are awaited (300,000 rows)", async function () {
        this.timeout(60000);
        const gc = exposeGc();
        // live memory: RSS also holds what the allocator keeps after a GC (V8's heap high-water mark, freed zlib
        // buffers), which makes it a noisy signal
        const liveMemory = () => {
            gc();
            const { heapUsed, external } = process.memoryUsage();
            return heapUsed + external;
        };
        const dir = fs.mkdtempSync(path.join(os.tmpdir(), "exceljs-backpressure-"));
        try {
            const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: path.join(dir, "big.xlsx") });
            const ws = wb.addWorksheet("data");
            let start;
            let peak = 0;
            for (let i = 1; i <= 300000; i++) {
                await ws.addRow(ROW(i)).commit();
                if (i % 25000 === 0) {
                    const live = liveMemory();
                    start = start === undefined ? live : start;
                    peak = Math.max(peak, live);
                }
            }
            await ws.commit();
            await wb.commit();
            // the sheet is ~170 MB of XML, which all stayed in memory before (rows committed faster than compressed);
            // now at most a few MB are being compressed at a time
            expect((peak - start) / 1e6).to.be.below(40);
        } finally {
            fs.rmSync(dir, { recursive: true, force: true });
        }
    });
});
