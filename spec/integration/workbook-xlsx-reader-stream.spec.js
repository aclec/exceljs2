// Streaming reader (ExcelJS.stream.xlsx.WorkbookReader) robustness tests.
// Fixtures are built in-test with ExcelJS + fflate so they stay reviewable.
const fs = require("fs");
const v8 = require("v8");
const vm = require("vm");
const { Readable } = require("stream");
const { baseWorkbookBuffer, tamper, unzipParts, zipParts } = require("../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

// feed the buffer in small chunks so that chunk boundaries fall everywhere
function chunkedStream(buffer, size = 1024) {
    return Readable.from(
        (function* () {
            for (let i = 0; i < buffer.length; i += size) {
                yield buffer.subarray(i, i + size);
            }
        })(),
        { objectMode: false }
    );
}

async function readAll(input, options) {
    const reader = new ExcelJS.stream.xlsx.WorkbookReader(input, options);
    const sheets = [];
    for await (const worksheet of reader) {
        const rows = [];
        for await (const row of worksheet) {
            rows.push({ number: row.number, values: row.values });
        }
        sheets.push({ id: worksheet.id, name: worksheet.name, worksheet, rows });
    }
    return sheets;
}

// rebuild the zip with its parts sorted by rank(path) (stable: equal ranks keep their order)
async function reorderZip(buffer, rank) {
    const source = unzipParts(buffer);
    const ordered = Object.keys(source)
        .map((name, index) => ({ name, index, rank: rank(name) }))
        .sort((a, b) => a.rank - b.rank || a.index - b.index)
        .map(({ name }) => name);
    const target = {};
    for (const name of ordered) {
        target[name] = source[name];
    }
    return zipParts(target);
}

async function readError(input) {
    try {
        await readAll(input);
    } catch (err) {
        return err;
    }
    return undefined;
}

const isWorksheet = (name) => /^xl\/worksheets\/sheet\d+\.xml$/.test(name);

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

// records every temp dir the reader creates (spy on the lookup done at call time)
function trackTempDirs() {
    const created = [];
    const original = fs.promises.mkdtemp;
    fs.promises.mkdtemp = async (...args) => {
        const dir = await original.apply(fs.promises, args);
        created.push(dir);
        return dir;
    };
    return {
        created,
        restore() {
            fs.promises.mkdtemp = original;
        },
    };
}

async function bigSheetFirstBuffer(rows = 3000) {
    const buffer = await baseWorkbookBuffer((wb, ws) => {
        for (let i = 2; i <= rows; i++) {
            ws.getCell(`A${i}`).value = `row ${i} ${(i * 7919) % 10007}`;
            ws.getCell(`B${i}`).value = i;
        }
        wb.addWorksheet("Second").getCell("A1").value = "second";
    });
    return reorderZip(buffer, (name) => (isWorksheet(name) ? 0 : 1));
}

// records the files the reader opens: file handles (fs.promises.open) and read streams
function trackOpenedFiles() {
    const handles = [];
    const streams = [];
    const { open } = fs.promises;
    const { createReadStream } = fs;
    fs.promises.open = async (...args) => {
        const handle = await open.apply(fs.promises, args);
        handles.push(handle);
        return handle;
    };
    fs.createReadStream = (...args) => {
        const stream = createReadStream.apply(fs, args);
        streams.push(stream);
        return stream;
    };
    return {
        handles,
        streams,
        restore() {
            fs.promises.open = open;
            fs.createReadStream = createReadStream;
        },
    };
}

// for each removal (fs.promises.rm), the number of files in `files` still open at that time
function trackRemovals(files) {
    const openFiles = [];
    const { rm } = fs.promises;
    fs.promises.rm = (...args) => {
        openFiles.push(files.handles.filter((handle) => handle.fd !== -1).length);
        return rm.apply(fs.promises, args);
    };
    return {
        openFiles,
        restore() {
            fs.promises.rm = rm;
        },
    };
}

// a workbook of `count` worksheets stored before the parts they depend on
async function manySheetsFirstBuffer(count) {
    const buffer = await baseWorkbookBuffer((wb) => {
        for (let i = 2; i <= count; i++) {
            wb.addWorksheet(`Sheet${i}`).getCell("A1").value = i;
        }
    });
    return reorderZip(buffer, (name) => (isWorksheet(name) ? 0 : 1));
}

describe("WorkbookReader (stream)", () => {
    describe("temp files for worksheets read before shared strings", () => {
        let tracker;
        beforeEach(() => {
            tracker = trackTempDirs();
        });
        afterEach(() => {
            tracker.restore();
        });

        it("spools the worksheets to one file, and reads a worksheet only when the consumer does", async () => {
            const buffer = await manySheetsFirstBuffer(300);
            const files = trackOpenedFiles();
            const names = [];
            let rows = 0;
            try {
                for await (const worksheet of new ExcelJS.stream.xlsx.WorkbookReader(chunkedStream(buffer))) {
                    names.push(worksheet.name);
                    if (worksheet.name !== "Sheet300") continue;
                    // eslint-disable-next-line no-unused-vars
                    for await (const row of worksheet) {
                        rows++;
                    }
                }
            } finally {
                files.restore();
            }
            expect(names).to.have.length(300);
            expect(rows).to.equal(1);
            // one file descriptor whatever the number of worksheets (thousands used to fail with EMFILE)
            expect(files.handles.length + files.streams.length).to.equal(1);
            expect(tracker.created).to.have.length(1);
            expect(fs.existsSync(tracker.created[0])).to.equal(false);
        });

        it("does not spool the worksheets when they are ignored", async () => {
            const buffer = await manySheetsFirstBuffer(3);
            const reader = new ExcelJS.stream.xlsx.WorkbookReader(chunkedStream(buffer), { worksheets: "ignore", entries: "emit" });
            const entries = [];
            reader.on("entry", (entry) => entries.push(entry));
            await new Promise((resolve, reject) => {
                reader.on("end", resolve);
                reader.on("error", reject);
                reader.read();
            });
            expect(tracker.created).to.have.length(0);
            expect(entries.filter((entry) => entry.type === "worksheet")).to.have.length(3);
        });

        it("removes the temp dir when the consumer stops early", async () => {
            const buffer = await bigSheetFirstBuffer();
            const reader = new ExcelJS.stream.xlsx.WorkbookReader(chunkedStream(buffer));
            let firstRow;

            for await (const worksheet of reader) {
                for await (const row of worksheet) {
                    firstRow = row.values;
                    break;
                }
                break;
            }
            expect(firstRow[1]).to.equal("hello");
            expect(tracker.created).to.have.length(1);
            expect(fs.existsSync(tracker.created[0])).to.equal(false);
        });

        it("closes the spooled worksheet file when the consumer stops early", async () => {
            const files = trackOpenedFiles();
            const removed = trackRemovals(files);
            try {
                const buffer = await bigSheetFirstBuffer();
                const reader = new ExcelJS.stream.xlsx.WorkbookReader(chunkedStream(buffer));
                // skip the first worksheet without reading it, stop at the second
                for await (const worksheet of reader) {
                    if (worksheet.name === "Second") break;
                }
            } finally {
                files.restore();
                removed.restore();
            }
            expect(files.handles).to.have.length(1);
            // closed before the directory is removed: Windows cannot remove open files
            expect(removed.openFiles).to.deep.equal([0]);
            expect(fs.existsSync(tracker.created[0])).to.equal(false);
        });

        it("reads spooled worksheets collected during the iteration", async () => {
            const buffer = await bigSheetFirstBuffer();
            const worksheets = [];
            for await (const worksheet of new ExcelJS.stream.xlsx.WorkbookReader(chunkedStream(buffer))) {
                worksheets.push(worksheet);
            }
            const counts = [];
            for (const worksheet of worksheets) {
                let count = 0;
                // eslint-disable-next-line no-unused-vars
                for await (const row of worksheet) {
                    count++;
                }
                counts.push(count);
            }
            expect(counts).to.deep.equal([3000, 1]);
            // removed once the collected worksheets have been read
            await new Promise((resolve) => setTimeout(resolve, 50));
            expect(fs.existsSync(tracker.created[0])).to.equal(false);
        });

        it("removes the temp dir when spooled worksheets are skipped", async () => {
            const buffer = await bigSheetFirstBuffer();
            for await (const worksheet of new ExcelJS.stream.xlsx.WorkbookReader(chunkedStream(buffer))) {
                if (worksheet.name !== "Second") continue;
                // eslint-disable-next-line no-unused-vars
                for await (const row of worksheet) {
                    // read
                }
            }
            expect(tracker.created).to.have.length(1);
            expect(fs.existsSync(tracker.created[0])).to.equal(false);
        });

        it("closes the spooled worksheet file once the reader of a skipped worksheet is collected", async function () {
            this.timeout(10000);
            const gc = exposeGc();
            const files = trackOpenedFiles();
            try {
                const buffer = await bigSheetFirstBuffer();
                for await (const worksheet of new ExcelJS.stream.xlsx.WorkbookReader(chunkedStream(buffer))) {
                    if (worksheet.name !== "Second") continue;
                    // eslint-disable-next-line no-unused-vars
                    for await (const row of worksheet) {
                        // read
                    }
                }
            } finally {
                files.restore();
            }
            expect(files.handles).to.have.length(1);
            // the skipped worksheet could still be read through its reader
            expect(files.handles[0].fd).to.not.equal(-1);
            // it can never be read once its reader is gone: the file must not stay open for the life of the process
            for (let i = 0; i < 50 && files.handles[0].fd !== -1; i++) {
                gc();
                await new Promise((resolve) => setTimeout(resolve, 20));
            }
            expect(files.handles[0].fd).to.equal(-1);
        });

        it("does not leave an unhandled rejection when closing the file of a collected worksheet fails", async function () {
            this.timeout(10000);
            const gc = exposeGc();
            const files = trackOpenedFiles();
            const { open } = fs.promises;
            fs.promises.open = async (...args) => {
                const handle = await open.apply(fs.promises, args);
                const close = handle.close.bind(handle);
                handle.close = async () => {
                    await close();
                    throw new Error("simulated close failure");
                };
                return handle;
            };
            // the rejection would be reported to mocha's own listener: record it instead
            const listeners = process.listeners("unhandledRejection");
            process.removeAllListeners("unhandledRejection");
            const unhandled = [];
            process.on("unhandledRejection", (reason) => unhandled.push(reason));
            try {
                const buffer = await bigSheetFirstBuffer();
                for await (const worksheet of new ExcelJS.stream.xlsx.WorkbookReader(chunkedStream(buffer))) {
                    if (worksheet.name !== "Second") continue;
                    // eslint-disable-next-line no-unused-vars
                    for await (const row of worksheet) {
                        // read
                    }
                }
                for (let i = 0; i < 50 && files.handles[0].fd !== -1; i++) {
                    gc();
                    await new Promise((resolve) => setTimeout(resolve, 20));
                }
                expect(files.handles[0].fd).to.equal(-1);
                await new Promise((resolve) => setTimeout(resolve, 20));
            } finally {
                fs.promises.open = open;
                files.restore();
                process.removeAllListeners("unhandledRejection");
                listeners.forEach((listener) => process.on("unhandledRejection", listener));
            }
            expect(unhandled.map((reason) => reason.message)).to.deep.equal([]);
        });

        it("closes the spooled worksheet file before removing the temp dir when every worksheet is read", async () => {
            const files = trackOpenedFiles();
            const removed = trackRemovals(files);
            let sheets;
            try {
                sheets = await readAll(chunkedStream(await bigSheetFirstBuffer()));
            } finally {
                files.restore();
                removed.restore();
            }
            expect(sheets.map((sheet) => sheet.rows.length)).to.deep.equal([3000, 1]);
            expect(files.handles).to.have.length(1);
            expect(removed.openFiles).to.deep.equal([0]);
            expect(fs.existsSync(tracker.created[0])).to.equal(false);
        });

        it("rejects the read of a collected worksheet when the iteration stopped early", async () => {
            const buffer = await bigSheetFirstBuffer();
            let first;
            for await (const worksheet of new ExcelJS.stream.xlsx.WorkbookReader(chunkedStream(buffer))) {
                first = worksheet;
                break;
            }
            let error;
            try {
                // eslint-disable-next-line no-unused-vars
                for await (const row of first) {
                    // read
                }
            } catch (err) {
                error = err;
            }
            expect(error).to.be.an.instanceOf(Error);
            expect(error.code).to.equal("EBADF");
        });

        it("rejects and removes the temp dir on a truncated file", async function () {
            this.timeout(10000);
            // large enough for the input to end inside the worksheet entry being spooled
            const buffer = await bigSheetFirstBuffer(20000);
            const truncated = buffer.subarray(0, Math.floor(buffer.length * 0.5));
            const error = await readError(chunkedStream(truncated));
            expect(error).to.be.an.instanceOf(Error);
            expect(tracker.created).to.have.length(1);
            expect(fs.existsSync(tracker.created[0])).to.equal(false);
        });

        it("rejects and removes the temp dir on corrupt compressed data", async function () {
            this.timeout(5000);
            const buffer = Buffer.from(await bigSheetFirstBuffer());
            // flip bytes in the middle of the first (deflated) worksheet entry
            for (let i = 2000; i < 2400; i++) {
                buffer[i] ^= 0x55;
            }
            let error;
            try {
                await readAll(chunkedStream(buffer));
            } catch (err) {
                error = err;
            }
            expect(error).to.be.an.instanceOf(Error);
            expect(tracker.created).to.have.length(1);
            expect(fs.existsSync(tracker.created[0])).to.equal(false);
        });
    });

    describe("input ending or failing inside a worksheet read directly", () => {
        let buffer;
        before(async function () {
            this.timeout(10000);
            // shared strings, rels, workbook and styles first: the worksheet is not spooled
            buffer = await reorderZip(await bigSheetFirstBuffer(20000), (name) => (isWorksheet(name) ? 1 : 0));
        });

        it("rejects when the input is truncated", async function () {
            this.timeout(5000);
            const error = await readError(chunkedStream(buffer.subarray(0, Math.floor(buffer.length * 0.8))));
            expect(error && error.message).to.equal("Unexpected end of zip data");
        });

        it("rejects with the error of the input stream", async function () {
            this.timeout(5000);
            const cut = Math.floor(buffer.length * 0.8);
            const input = Readable.from(
                (function* () {
                    for (let i = 0; i < cut; i += 1024) {
                        yield buffer.subarray(i, Math.min(i + 1024, cut));
                    }
                    throw new Error("connection reset");
                })(),
                { objectMode: false }
            );
            const error = await readError(input);
            expect(error && error.message).to.equal("connection reset");
        });
    });

    it("rejects when the input file does not exist", async () => {
        let error;
        try {
            await readAll("./spec/out/does-not-exist.xlsx");
        } catch (err) {
            error = err;
        }
        expect(error).to.be.an.instanceOf(Error);
        expect(error.code).to.equal("ENOENT");
    });

    it("reads many worksheets stored before sharedStrings and workbook rels", async function () {
        this.timeout(10000);
        const count = 50;
        const buffer = await baseWorkbookBuffer((wb, ws) => {
            ws.getCell("B1").value = 1;
            for (let i = 2; i <= count; i++) {
                const sheet = wb.addWorksheet(`Sheet${i}`);
                sheet.getCell("A1").value = `text ${i}`;
                sheet.getCell("B1").value = i;
            }
        });
        // worksheets first, then everything they depend on
        const reordered = await reorderZip(buffer, (name) => (isWorksheet(name) ? 0 : 1));

        const sheets = await readAll(chunkedStream(reordered));
        expect(sheets).to.have.length(count);
        sheets.forEach((sheet, index) => {
            const n = index + 1;
            expect(sheet.name).to.equal(`Sheet${n}`);
            expect(sheet.rows).to.have.length(1);
            expect(sheet.rows[0].values[1]).to.equal(n === 1 ? "hello" : `text ${n}`);
            expect(sheet.rows[0].values[2]).to.equal(n);
        });
    });

    it("names worksheets when only workbook.xml comes after the worksheets", async () => {
        const buffer = await baseWorkbookBuffer((wb) => {
            wb.addWorksheet("Second").getCell("A1").value = "two";
        });
        // shared strings and rels first, then the worksheets, workbook.xml last
        const rank = (name) => {
            if (name === "xl/workbook.xml") return 2;
            return isWorksheet(name) ? 1 : 0;
        };
        const reordered = await reorderZip(buffer, rank);

        const sheets = await readAll(chunkedStream(reordered));
        expect(sheets.map((sheet) => sheet.name)).to.deep.equal(["Sheet1", "Second"]);
        expect(sheets[1].rows[0].values[1]).to.equal("two");
    });

    it("applies cached styles when styles.xml comes after the worksheets", async () => {
        const date = new Date(Date.UTC(2024, 1, 29));
        const buffer = await baseWorkbookBuffer((wb, ws) => {
            ws.getCell("B1").value = date;
            ws.getCell("B1").numFmt = "yyyy-mm-dd";
        });
        // shared strings, rels and workbook first, then the worksheets, styles.xml last
        const rank = (name) => {
            if (name === "xl/styles.xml") return 2;
            return isWorksheet(name) ? 1 : 0;
        };
        const reordered = await reorderZip(buffer, rank);

        const sheets = await readAll(chunkedStream(reordered), { styles: "cache" });
        expect(sheets[0].rows[0].values[2]).to.equalTime(date);
    });

    it("reads error and boolean formula results like the document reader", async () => {
        const buffer = await tamper(
            await baseWorkbookBuffer((wb, ws) => {
                ws.getCell("A1").value = { formula: "NA()", result: { error: "#N/A" } };
                ws.getCell("A2").value = { formula: "1=1", result: true };
                ws.getCell("A3").value = { formula: "1=2", result: false };
                ws.getCell("A4").value = { formula: "1+1", result: 2 };
                ws.getCell("A5").value = { formula: '"a"', result: "a" };
            }),
            {
                // as Excel writes them
                "xl/worksheets/sheet1.xml": (xml) => {
                    expect(xml).to.match(/<c r="A1" t="e"><f>NA\(\)<\/f><v>#N\/A<\/v><\/c>/);
                    expect(xml).to.match(/<c r="A2" t="b"><f>1=1<\/f><v>1<\/v><\/c>/);
                    return xml;
                },
            }
        );
        const doc = new ExcelJS.Workbook();
        await doc.xlsx.load(buffer);
        const expected = doc.getWorksheet("Sheet1").getColumn(1).values.slice(1, 6);
        expect(expected.map((value) => value.result)).to.deep.equal([{ error: "#N/A" }, true, false, 2, "a"]);

        const [sheet] = await readAll(chunkedStream(buffer));
        const streamed = sheet.rows.filter((row) => row.number <= 5).map((row) => row.values[1]);
        expect(streamed).to.deep.equal(expected);
    });

    describe("columns and rows", () => {
        it("reads column widths, styles and hidden flags", async () => {
            const buffer = await baseWorkbookBuffer((wb, ws) => {
                ws.getColumn(1).width = 25;
                ws.getColumn(2).hidden = true;
                ws.getColumn(3).width = 12.5;
                ws.getColumn(3).numFmt = "0.00";
                ws.getColumn(3).outlineLevel = 1;
                ws.getCell("C1").value = 3;
            });
            const [sheet] = await readAll(chunkedStream(buffer), { styles: "cache" });
            const { worksheet } = sheet;
            expect(worksheet.columns).to.have.length(3);
            expect(worksheet.getColumn(1).width).to.equal(25);
            expect(worksheet.getColumn(1).hidden).to.equal(false);
            expect(worksheet.getColumn(2).hidden).to.equal(true);
            expect(worksheet.getColumn("C").width).to.equal(12.5);
            expect(worksheet.getColumn("C").outlineLevel).to.equal(1);
            expect(worksheet.getColumn("C").numFmt).to.equal("0.00");
        });

        it("reads outline levels and sheet format properties", async () => {
            const buffer = await baseWorkbookBuffer((wb, ws) => {
                ws.properties.outlineLevelCol = 1;
                ws.properties.outlineLevelRow = 1;
                ws.getColumn(2).outlineLevel = 1;
                ws.getCell("A2").value = "grouped";
                ws.getRow(2).outlineLevel = 1;
            });
            const reader = new ExcelJS.stream.xlsx.WorkbookReader(chunkedStream(buffer));
            const rows = [];
            let worksheet;
            for await (worksheet of reader) {
                for await (const row of worksheet) {
                    rows.push({ outlineLevel: row.outlineLevel, collapsed: row.collapsed });
                }
            }
            expect(worksheet.properties.outlineLevelCol).to.equal(1);
            expect(worksheet.properties.outlineLevelRow).to.equal(1);
            expect(worksheet.getColumn(1).collapsed).to.equal(false);
            expect(worksheet.getColumn(2).outlineLevel).to.equal(1);
            expect(worksheet.getColumn(2).collapsed).to.equal(true);
            expect(rows).to.deep.equal([
                { outlineLevel: 0, collapsed: false },
                { outlineLevel: 1, collapsed: true },
            ]);
        });

        it("treats column keys named like Object.prototype members as plain keys", async () => {
            const [{ worksheet }] = await readAll(chunkedStream(await baseWorkbookBuffer()));
            expect(worksheet.getColumnKey("constructor")).to.equal(undefined);
            worksheet.getColumn(1).key = "__proto__";
            worksheet.getColumn(2).key = "constructor";
            expect(worksheet.getColumn("__proto__").number).to.equal(1);
            expect(worksheet.getColumn("constructor").number).to.equal(2);
            const keys = [];
            worksheet.eachColumnKey((column, key) => keys.push(`${key}:${column.number}`));
            expect(keys).to.deep.equal(["__proto__:1", "constructor:2"]);
        });

        it("clamps column ranges to the worksheet grid", async () => {
            const buffer = await tamper(await baseWorkbookBuffer(), {
                "xl/worksheets/sheet1.xml": (xml) =>
                    xml.replace("<sheetData>", '<cols><col min="2" max="2000000000" width="5"/></cols><sheetData>'),
            });
            const [sheet] = await readAll(chunkedStream(buffer));
            expect(sheet.worksheet.columns).to.have.length(16384);
            expect(sheet.worksheet.getColumn(16384).width).to.equal(5);
        });

        it("returns every column up to the last one read or used", async () => {
            const buffer = await baseWorkbookBuffer((wb, ws) => {
                ws.getColumn(2).width = 12;
                ws.getColumn(3).width = 12;
                ws.getCell("F1").value = "f";
            });
            const [{ worksheet, rows }] = await readAll(chunkedStream(buffer));
            expect(rows[0].values).to.deep.equal([undefined, "hello", undefined, undefined, undefined, undefined, "f"]);
            expect(worksheet.getColumn(4).width).to.equal(undefined);
            const { columns } = worksheet;
            expect(columns.map((column) => [column.number, column.width])).to.deep.equal([
                [1, undefined],
                [2, 12],
                [3, 12],
                [4, undefined],
                [5, undefined],
                [6, undefined],
            ]);
            expect(columns[3]).to.equal(worksheet.getColumn(4));
            expect(worksheet.getColumn(0)).to.equal(undefined);
        });

        it("reads hidden rows", async () => {
            const buffer = await baseWorkbookBuffer((wb, ws) => {
                ws.getCell("A2").value = "hidden";
                ws.getRow(2).hidden = true;
                ws.getCell("A3").value = "visible";
            });
            const reader = new ExcelJS.stream.xlsx.WorkbookReader(chunkedStream(buffer));
            const hidden = [];
            for await (const worksheet of reader) {
                for await (const row of worksheet) {
                    hidden.push(row.hidden);
                }
            }
            expect(hidden).to.deep.equal([false, true, false]);
        });
    });

    it("names worksheets whose relationship targets are absolute paths", async () => {
        const buffer = await baseWorkbookBuffer((wb) => {
            wb.addWorksheet("Named").getCell("A1").value = "two";
        });
        const tampered = await tamper(buffer, {
            // openpyxl style absolute targets, plus a relationship without Target
            "xl/_rels/workbook.xml.rels": (xml) =>
                xml
                    .replace(/Target="worksheets\//g, 'Target="/xl/worksheets/')
                    .replace("</Relationships>", '<Relationship Id="rId99" Type="http://example.com/none"/></Relationships>'),
        });
        const sheets = await readAll(chunkedStream(tampered));
        expect(sheets.map((sheet) => sheet.name)).to.deep.equal(["Sheet1", "Named"]);
    });

    it("names many worksheets among many relationships in linear time", async function () {
        this.timeout(20000);
        const W = 3000;
        const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
        const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        // the worksheet relationships come last: every lookup used to scan the 150,000 relationships and the sheets
        let rels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
        rels += '<Relationship Id="x" Type="t" Target="other.xml"/>'.repeat(150000);
        let sheets = "";
        const parts = {};
        for (let i = 1; i <= W; i++) {
            rels += `<Relationship Id="rId${i}" Type="${R}/worksheet" Target="worksheets/sheet${i}.xml"/>`;
            sheets += `<sheet name="S${i}" sheetId="${i}" r:id="rId${i}"/>`;
        }
        parts["xl/_rels/workbook.xml.rels"] = `${rels}</Relationships>`;
        parts["xl/workbook.xml"] = `<workbook ${NS} xmlns:r="${R}"><sheets>${sheets}</sheets></workbook>`;
        parts["xl/sharedStrings.xml"] = `<sst ${NS}/>`;
        for (let i = 1; i <= W; i++) {
            parts[`xl/worksheets/sheet${i}.xml`] = `<worksheet ${NS}><sheetData/></worksheet>`;
        }
        const start = Date.now();
        const names = [];
        for await (const worksheet of new ExcelJS.stream.xlsx.WorkbookReader(Readable.from([zipParts(parts)]))) {
            names.push(worksheet.name);
        }
        const ms = Date.now() - start;
        expect(names).to.have.length(W);
        expect(names[W - 1]).to.equal(`S${W}`);
        // about 0.2 s; 7 s with the linear lookups
        expect(ms).to.be.below(4000);
    });

    describe("cell references", () => {
        async function rowsOf(buffer) {
            const [sheet] = await readAll(chunkedStream(buffer));
            return sheet.rows.map(({ number, values }) => ({ number, values: values.slice(1) }));
        }

        it("infers row and column numbers when r is missing", async () => {
            const buffer = await baseWorkbookBuffer((wb, ws) => {
                ws.getCell("B1").value = 1;
                ws.getCell("C1").value = "c1";
                ws.getCell("A2").value = "a2";
                ws.getCell("B2").value = 2;
                ws.getCell("A3").value = "a3";
            });
            const tampered = await tamper(buffer, {
                "xl/worksheets/sheet1.xml": (xml) => xml.replace(/ r="[A-Z]*\d+"/g, ""),
            });
            expect(await rowsOf(tampered)).to.deep.equal([
                { number: 1, values: ["hello", 1, "c1"] },
                { number: 2, values: ["a2", 2] },
                { number: 3, values: ["a3"] },
            ]);
        });

        it("continues numbering after explicit references", async () => {
            const sheetData =
                "<sheetData>" +
                '<row r="2"><c r="B2"><v>1</v></c><c><v>2</v></c></row>' +
                '<row><c><v>3</v></c><c r="D3"><v>4</v></c><c><v>5</v></c></row>' +
                "</sheetData>";
            const buffer = await tamper(await baseWorkbookBuffer(), {
                "xl/worksheets/sheet1.xml": (xml) => xml.replace(/<sheetData>.*<\/sheetData>/, sheetData),
            });
            expect(await rowsOf(buffer)).to.deep.equal([
                { number: 2, values: [undefined, 1, 2] },
                { number: 3, values: [3, undefined, undefined, 4, 5] },
            ]);
        });
    });

    it('reads ISO 8601 date cells (t="d") as UTC dates', async () => {
        const cells = [
            "2024-02-29T12:34:56Z",
            "2024-02-29T12:34:56.5",
            "2024-02-29",
            "2024-02-29T12:34:56+02:00",
            "2024-02-30",
            "not a date",
        ]
            .map((text) => `<c t="d"><v>${text}</v></c>`)
            .join("");
        const buffer = await tamper(await baseWorkbookBuffer(), {
            "xl/worksheets/sheet1.xml": (xml) =>
                xml.replace(/<sheetData>.*<\/sheetData>/, `<sheetData><row r="1">${cells}</row></sheetData>`),
        });
        const [sheet] = await readAll(chunkedStream(buffer));
        const values = sheet.rows[0].values.slice(1);
        expect(values.slice(0, 4).map((value) => value instanceof Date && value.toISOString())).to.deep.equal([
            "2024-02-29T12:34:56.000Z",
            "2024-02-29T12:34:56.500Z",
            "2024-02-29T00:00:00.000Z",
            "2024-02-29T10:34:56.000Z",
        ]);
        expect(values.slice(4)).to.deep.equal(["2024-02-30", "not a date"]);
    });

    it('reads t="d" cells like the document reader', async () => {
        const texts = [
            "2024-02-29T24:00",
            "2024-02-29T24:00:01",
            "0099-12-31T10:00",
            "2024-03-01T00:30+01:00",
            "2024-02-29T12:30+15:00",
            "2024-02-29T25:00",
            "2023-02-29",
        ];
        const cells = texts.map((text) => `<c t="d"><v>${text}</v></c>`).join("");
        const buffer = await tamper(await baseWorkbookBuffer(), {
            "xl/worksheets/sheet1.xml": (xml) =>
                xml.replace(/<sheetData>.*<\/sheetData>/, `<sheetData><row r="1">${cells}</row></sheetData>`),
        });
        const [sheet] = await readAll(chunkedStream(buffer));
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buffer);
        const documentValues = wb.getWorksheet(1).getRow(1).values.slice(1);
        expect(sheet.rows[0].values.slice(1)).to.deep.equal(documentValues);
        expect(documentValues.map((value) => (value instanceof Date ? value.toISOString() : value))).to.deep.equal([
            "2024-03-01T00:00:00.000Z",
            "2024-02-29T24:00:01",
            "0099-12-31T10:00:00.000Z",
            "2024-02-29T23:30:00.000Z",
            "2024-02-29T12:30+15:00",
            "2024-02-29T25:00",
            "2023-02-29",
        ]);
    });

    describe("hyperlinks: emit with read()", () => {
        const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
        // relationships that do not compress well: the part spans many chunks of the zip input
        async function hyperlinksBuffer(count) {
            let rels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
            for (let i = 0; i < count; i++) {
                const target = require("crypto").randomBytes(24).toString("hex");
                rels += `<Relationship Id="rId${i}" Type="${REL}" Target="https://example.com/${target}" TargetMode="External"/>`;
            }
            const parts = unzipParts(await baseWorkbookBuffer());
            parts["xl/worksheets/_rels/sheet1.xml.rels"] = `${rels}</Relationships>`;
            return zipParts(parts);
        }

        async function readHyperlinks(buffer, startRead) {
            const reader = new ExcelJS.stream.xlsx.WorkbookReader(chunkedStream(buffer, 64 * 1024), { hyperlinks: "emit" });
            let count = 0;
            await new Promise((resolve, reject) => {
                reader.on("hyperlinks", (hyperlinks) => {
                    hyperlinks.on("hyperlink", () => count++);
                    if (startRead) {
                        // as in the README
                        hyperlinks.read();
                    }
                });
                reader.on("worksheet", (worksheet) => worksheet.on("row", () => {}));
                reader.on("end", resolve);
                reader.on("error", reject);
                reader.read();
            });
            return count;
        }

        it("reads all the relationships of a large part before moving on", async () => {
            const buffer = await hyperlinksBuffer(10000);
            expect(await readHyperlinks(buffer, true)).to.equal(10000);
            // the reader reads them when the listener does not
            expect(await readHyperlinks(buffer, false)).to.equal(10000);
        });
    });

    it("merges options passed to read() with the defaults", async () => {
        const buffer = await baseWorkbookBuffer();
        const reader = new ExcelJS.stream.xlsx.WorkbookReader();
        const events = [];
        await new Promise((resolve, reject) => {
            reader.on("worksheet", (worksheet) => {
                events.push(`worksheet ${worksheet.name}`);
                worksheet.on("row", (row) => events.push(`row ${JSON.stringify(row.values)}`));
            });
            reader.on("end", resolve);
            reader.on("error", reject);
            reader.read(chunkedStream(buffer), { sharedStrings: "emit" });
        });
        expect(reader.options.worksheets).to.equal("emit");
        expect(events).to.deep.equal(["worksheet Sheet1", 'row [null,{"sharedString":0}]']);
    });

    it("reads shared strings as the document reader does", async () => {
        const fonts = [
            { color: { argb: "FFFF0000" }, name: "Arial", size: 9, family: 2, charset: 1, scheme: "minor" },
            { bold: true, italic: true, underline: "double", strike: true, outline: true, vertAlign: "superscript" },
            { color: { theme: 1, tint: 0.5 }, underline: true },
        ];
        const buffer = await tamper(
            await baseWorkbookBuffer((wb, ws) => {
                ws.getCell("B1").value = { richText: fonts.map((font, i) => ({ font, text: `run ${i}` })) };
                ws.getCell("C1").value = { richText: [{ text: "plain" }, { font: { bold: true }, text: " bold" }] };
                ws.getCell("D1").value = "last";
            }),
            {
                // a phonetic run (not part of the text), an escaped character and an empty string
                "xl/sharedStrings.xml": (xml) =>
                    xml.replace(
                        "<si><t>last</t></si>",
                        '<si><t>last</t><rPh sb="0" eb="1"><t>phonetic</t></rPh></si>' + "<si><t>a_x0041_b</t></si><si><t/></si>"
                    ),
                "xl/worksheets/sheet1.xml": (xml) =>
                    xml.replace("</row>", '<c r="E1" t="s"><v>4</v></c><c r="F1" t="s"><v>5</v></c></row>'),
            }
        );
        const doc = new ExcelJS.Workbook();
        await doc.xlsx.load(buffer);
        const expected = doc.getWorksheet(1).getRow(1).values;
        expect(expected[4]).to.equal("last");
        expect(expected[5]).to.equal("aAb");
        const [sheet] = await readAll(chunkedStream(buffer));
        expect(sheet.rows[0].values).to.deep.equal(expected);
    });

    it("gives the cells of the shared string 0 a shared string value (sharedStrings: emit)", async () => {
        const buffer = await baseWorkbookBuffer((wb, ws) => {
            ws.getCell("B1").value = "other";
        });
        const reader = new ExcelJS.stream.xlsx.WorkbookReader(chunkedStream(buffer), { sharedStrings: "emit" });
        const cells = [];
        for await (const worksheet of reader) {
            for await (const row of worksheet) {
                row.eachCell((cell) => cells.push([cell.address, cell.type, cell.value]));
            }
        }
        expect(cells).to.deep.equal([
            ["A1", ExcelJS.ValueType.SharedString, { sharedString: 0 }],
            ["B1", ExcelJS.ValueType.SharedString, { sharedString: 1 }],
        ]);
    });

    it("gives a shared formula clone a formula value (without formula) and its result", async () => {
        const buffer = await baseWorkbookBuffer((wb, ws) => {
            ws.getCell("B1").value = { formula: "A1&1", result: "hello1", shareType: "shared", ref: "B1:B2" };
            ws.getCell("B2").value = { sharedFormula: "B1", result: 0 };
        });
        const [sheet] = await readAll(chunkedStream(buffer));
        const clone = sheet.rows[1].values[2];
        expect(clone).to.deep.equal({ formula: "", result: 0 });
        const cells = [];
        for await (const worksheet of new ExcelJS.stream.xlsx.WorkbookReader(chunkedStream(buffer))) {
            for await (const row of worksheet) {
                const cell = row.getCell(2);
                cells.push([cell.type, cell.formulaType, cell.result, cell.text]);
            }
        }
        expect(cells).to.deep.equal([
            [ExcelJS.ValueType.Formula, ExcelJS.FormulaType.Master, "hello1", "hello1"],
            [ExcelJS.ValueType.Formula, ExcelJS.FormulaType.None, 0, "0"],
        ]);
    });

    it("emits shared strings in sharedStrings: emit mode", async () => {
        const buffer = await baseWorkbookBuffer((wb, ws) => {
            ws.getCell("B1").value = { richText: [{ font: { bold: true }, text: "bold" }] };
        });
        const reader = new ExcelJS.stream.xlsx.WorkbookReader(chunkedStream(buffer), {
            sharedStrings: "emit",
            worksheets: "ignore",
        });
        const strings = [];
        await new Promise((resolve, reject) => {
            reader.on("shared-strings", (value) => strings.push(value));
            reader.on("end", resolve);
            reader.on("error", reject);
            reader.read();
        });
        expect(strings).to.deep.equal([
            { index: 0, text: "hello" },
            { index: 1, text: { richText: [{ font: { bold: true }, text: "bold" }] } },
        ]);
    });
});
