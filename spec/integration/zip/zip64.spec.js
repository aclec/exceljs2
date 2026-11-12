// ZIP64 output of the zip writer. Real ZIP64 archives need more than 4 GiB or 65535 entries, so the
// limits are lowered through the internal `_zip64Limits` zip option (tests only), or every record is
// forced to ZIP64 with the public `forceZip64` option. The archives are checked with fflate, with both
// ExcelJS readers and, when they are installed, with `unzip -t` and Python's zipfile.
const { execFileSync } = require("child_process");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { PassThrough, Readable } = require("stream");
const { unzipSync } = require("fflate");

const { ZipWriter } = verquire("utils/zip-stream");
const { readZip, parseZipStream } = verquire("utils/zip-reader");

const ExcelJS = verquire("exceljs");

const SIG_LOCAL = 0x04034b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_ZIP64_EOCD = 0x06064b50;
const SIG_ZIP64_LOCATOR = 0x07064b50;
const SIG_DATA_DESCRIPTOR = 0x08074b50;

function hasCommand(command, args) {
    try {
        execFileSync(command, args, { stdio: "ignore" });
        return true;
    } catch (error) {
        return false;
    }
}
const HAS_UNZIP = hasCommand("unzip", ["-v"]);
const HAS_PYTHON = hasCommand("python3", ["-I", "-c", "import zipfile"]);

// {name: sha256} of every file, read by Python's zipfile (which checks the CRC-32 of each one)
const PYTHON_LIST = [
    "import hashlib, json, sys, zipfile",
    "with zipfile.ZipFile(sys.argv[1]) as z:",
    "    assert z.testzip() is None",
    "    print(json.dumps({i.filename: hashlib.sha256(z.read(i)).hexdigest() for i in z.infolist()}))",
].join("\n");

function sha256(data) {
    return require("crypto").createHash("sha256").update(data).digest("hex");
}

async function withTempFile(buffer, fn) {
    const dir = await fs.promises.mkdtemp(path.join(os.tmpdir(), "exceljs-zip64-"));
    try {
        const file = path.join(dir, "archive.zip");
        await fs.promises.writeFile(file, buffer);
        return fn(file);
    } finally {
        await fs.promises.rm(dir, { recursive: true, force: true });
    }
}

// the archive must read the same everywhere: returns {name: content}
async function checkArchive(buffer, expected) {
    const fflateFiles = unzipSync(new Uint8Array(buffer));
    expect(Object.keys(fflateFiles)).to.deep.equal(Object.keys(expected));
    for (const [name, data] of Object.entries(expected)) {
        expect(Buffer.from(fflateFiles[name]).equals(data), `fflate: ${name}`).to.equal(true);
    }

    const documentEntries = readZip(buffer);
    expect(documentEntries.map((entry) => entry.name)).to.deep.equal(Object.keys(expected));
    for (const entry of documentEntries) {
        expect(Buffer.from(await entry.data()).equals(expected[entry.name]), `readZip: ${entry.name}`).to.equal(true);
    }

    const streamed = {};
    for await (const entry of parseZipStream(Readable.from([buffer.subarray(0, 100), buffer.subarray(100)]))) {
        const chunks = [];
        for await (const chunk of entry) {
            chunks.push(chunk);
        }
        streamed[entry.path] = Buffer.concat(chunks);
    }
    expect(Object.keys(streamed)).to.deep.equal(Object.keys(expected));
    for (const [name, data] of Object.entries(expected)) {
        expect(streamed[name].equals(data), `parseZipStream: ${name}`).to.equal(true);
    }

    await withTempFile(buffer, (file) => {
        if (HAS_UNZIP) {
            expect(execFileSync("unzip", ["-t", file]).toString()).to.include("No errors detected");
        }
        if (HAS_PYTHON) {
            const hashes = JSON.parse(execFileSync("python3", ["-I", "-c", PYTHON_LIST, file]).toString());
            const expectedHashes = {};
            Object.entries(expected).forEach(([name, data]) => (expectedHashes[name] = sha256(data)));
            expect(hashes).to.deep.equal(expectedHashes);
        }
    });
}

// the records of an archive, walked from the start
function records(buffer) {
    const result = { local: [], descriptors: [], central: [], zip64End: 0, zip64Locator: 0 };
    for (let pos = 0; pos + 4 <= buffer.length; pos++) {
        switch (buffer.readUInt32LE(pos)) {
            case SIG_LOCAL: {
                const nameLength = buffer.readUInt16LE(pos + 26);
                const extraLength = buffer.readUInt16LE(pos + 28);
                result.local.push({
                    name: buffer.toString("utf8", pos + 30, pos + 30 + nameLength),
                    version: buffer.readUInt16LE(pos + 4),
                    flags: buffer.readUInt16LE(pos + 6),
                    zip64Extra: extraLength >= 4 && buffer.readUInt16LE(pos + 30 + nameLength) === 1,
                });
                break;
            }
            case SIG_DATA_DESCRIPTOR:
                result.descriptors.push(pos);
                break;
            case SIG_CENTRAL: {
                const nameLength = buffer.readUInt16LE(pos + 28);
                const extraLength = buffer.readUInt16LE(pos + 30);
                result.central.push({
                    name: buffer.toString("utf8", pos + 46, pos + 46 + nameLength),
                    offset: buffer.readUInt32LE(pos + 42),
                    zip64Extra: extraLength >= 4 && buffer.readUInt16LE(pos + 46 + nameLength) === 1,
                });
                break;
            }
            case SIG_ZIP64_EOCD:
                result.zip64End++;
                break;
            case SIG_ZIP64_LOCATOR:
                result.zip64Locator++;
                break;
            default:
                break;
        }
    }
    return result;
}

async function writeZip(options, build) {
    const zip = new ZipWriter(options);
    const chunks = [];
    zip.on("data", (chunk) => chunks.push(chunk));
    const expected = build(zip);
    await zip.finalize();
    return { buffer: Buffer.concat(chunks), expected };
}

// a streamed entry (sizes unknown in its local header), then complete entries, one larger than 1000 bytes
function mixedEntries(zip) {
    const expected = {};
    const sheet = zip.openStream("xl/worksheets/sheet1.xml");
    const rows = [];
    for (let i = 0; i < 300; i++) {
        rows.push(`<row r="${i + 1}"><c><v>${(i * 7919) % 10007}</v></c></row>`);
    }
    rows.forEach((row) => sheet.write(row));
    zip.append("small", { name: "a.txt" });
    const big = Buffer.alloc(3000, "x");
    zip.append(big, { name: "xl/media/image1.bin" });
    sheet.end();
    expected["xl/worksheets/sheet1.xml"] = Buffer.from(rows.join(""));
    expected["a.txt"] = Buffer.from("small");
    expected["xl/media/image1.bin"] = big;
    for (let i = 0; i < 5; i++) {
        zip.append(`entry ${i}`, { name: `docProps/e${i}.txt` });
        expected[`docProps/e${i}.txt`] = Buffer.from(`entry ${i}`);
    }
    return expected;
}

describe("Zip writer ZIP64", function () {
    this.timeout(20000);

    it("writes no ZIP64 record below the limits", async () => {
        const { buffer, expected } = await writeZip({}, mixedEntries);
        const found = records(buffer);
        expect(found.zip64End + found.zip64Locator).to.equal(0);
        expect(found.local.some((entry) => entry.zip64Extra)).to.equal(false);
        expect(found.central.some((entry) => entry.zip64Extra)).to.equal(false);
        // the streamed entry has a 16 byte data descriptor, the others their sizes in the local header
        expect(found.local.map((entry) => entry.flags & 8)).to.deep.equal([8, 0, 0, 0, 0, 0, 0, 0]);
        expect(found.descriptors).to.have.length(1);
        await checkArchive(buffer, expected);
    });

    it("adds the ZIP64 records past the size and entry count limits", async () => {
        const { buffer, expected } = await writeZip({ _zip64Limits: { size: 1000, entries: 8 } }, mixedEntries);
        const found = records(buffer);
        // 8 entries: ZIP64 end of central directory record and locator
        expect(found.zip64End).to.equal(1);
        expect(found.zip64Locator).to.equal(1);
        expect(buffer.readUInt16LE(buffer.length - 22 + 10)).to.equal(0xffff);
        const local = Object.fromEntries(found.local.map((entry) => [entry.name, entry]));
        // a complete entry larger than the limit announces 64-bit sizes in its local header
        expect(local["xl/media/image1.bin"].zip64Extra).to.equal(true);
        expect(local["xl/media/image1.bin"].version).to.equal(45);
        expect(local["a.txt"].zip64Extra).to.equal(false);
        // the streamed entry (larger than the limit) gets a 24 byte data descriptor with 64-bit sizes
        const descriptor = found.descriptors[0];
        expect(buffer.readUInt32LE(descriptor + 24)).to.equal(SIG_LOCAL);
        // in the central directory, every size or offset past the limit moves to the ZIP64 field
        const central = Object.fromEntries(found.central.map((entry) => [entry.name, entry]));
        expect(central["xl/worksheets/sheet1.xml"].zip64Extra).to.equal(true);
        expect(central["xl/media/image1.bin"].zip64Extra).to.equal(true);
        // a.txt is small, but written after the streamed entry: its offset is past the limit
        expect(central["a.txt"].zip64Extra).to.equal(true);
        expect(central["a.txt"].offset).to.equal(0xffffffff);
        await checkArchive(buffer, expected);
    });

    it("writes every record as ZIP64 with forceZip64", async () => {
        const { buffer, expected } = await writeZip({ forceZip64: true }, mixedEntries);
        const found = records(buffer);
        expect(found.zip64End).to.equal(1);
        expect(found.zip64Locator).to.equal(1);
        expect(found.local.every((entry) => entry.zip64Extra)).to.equal(true);
        expect(found.central.every((entry) => entry.zip64Extra)).to.equal(true);
        await checkArchive(buffer, expected);
    });

    it("reads back a workbook streamed into a ZIP64 archive", async () => {
        const write = async (zip) => {
            const stream = new PassThrough();
            const chunks = [];
            stream.on("data", (chunk) => chunks.push(chunk));
            const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ stream, useSharedStrings: true, zip });
            for (const name of ["One", "Two"]) {
                const ws = wb.addWorksheet(name);
                for (let i = 1; i <= 200; i++) {
                    ws.addRow([name, i, `text ${i}`]).commit();
                }
                ws.commit();
            }
            await wb.commit();
            return Buffer.concat(chunks);
        };
        for (const zip of [{ forceZip64: true }, { _zip64Limits: { size: 2000, entries: 4 } }]) {
            const buffer = await write(zip);
            expect(records(buffer).zip64End).to.equal(1);

            const wb = new ExcelJS.Workbook();
            await wb.xlsx.load(buffer);
            expect(wb.worksheets.map((ws) => ws.name)).to.deep.equal(["One", "Two"]);
            expect(wb.getWorksheet("Two").getRow(200).values).to.deep.equal([undefined, "Two", 200, "text 200"]);

            const rows = [];
            for await (const ws of new ExcelJS.stream.xlsx.WorkbookReader(Readable.from([buffer]))) {
                for await (const row of ws) {
                    rows.push(row.values.slice(1));
                }
            }
            expect(rows).to.have.length(400);
            expect(rows[399]).to.deep.equal(["Two", 200, "text 200"]);

            const parts = {};
            Object.entries(unzipSync(new Uint8Array(buffer))).forEach(([name, data]) => (parts[name] = Buffer.from(data)));
            await checkArchive(buffer, parts);
        }
    });
});
