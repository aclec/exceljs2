// Streaming reader (ExcelJS.stream.xlsx.WorkbookReader) robustness tests.
// Fixtures are built in-test with ExcelJS + fflate so they stay reviewable.
const fs = require("fs");
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

describe("WorkbookReader (stream)", () => {
    describe("temp files for worksheets read before shared strings", () => {
        let tracker;
        beforeEach(() => {
            tracker = trackTempDirs();
        });
        afterEach(() => {
            tracker.restore();
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

        it("closes the spooled worksheet files when the consumer stops early", async () => {
            const opened = [];
            const original = fs.createReadStream;
            fs.createReadStream = (...args) => {
                const stream = original.apply(fs, args);
                opened.push(stream);
                return stream;
            };
            try {
                const buffer = await bigSheetFirstBuffer();
                const reader = new ExcelJS.stream.xlsx.WorkbookReader(chunkedStream(buffer));
                // skip the first worksheet without reading it, stop at the second
                for await (const worksheet of reader) {
                    if (worksheet.name === "Second") break;
                }
            } finally {
                fs.createReadStream = original;
            }
            expect(opened).to.have.length(2);
            // closed (not only destroyed) before the directory is removed: Windows cannot remove open files
            opened.forEach((stream) => expect(stream.closed).to.equal(true));
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
            await readAll(chunkedStream(buffer), { worksheets: "ignore" });
            expect(tracker.created).to.have.length(2);
            tracker.created.forEach((dir) => expect(fs.existsSync(dir)).to.equal(false));
        });

        it("closes the spooled worksheet files read to the end before removing the temp dir", async () => {
            const opened = [];
            const original = fs.createReadStream;
            fs.createReadStream = (...args) => {
                const stream = original.apply(fs, args);
                opened.push(stream);
                return stream;
            };
            try {
                await readAll(chunkedStream(await bigSheetFirstBuffer()));
            } finally {
                fs.createReadStream = original;
            }
            expect(opened).to.have.length(2);
            opened.forEach((stream) => expect(stream.closed).to.equal(true));
            expect(fs.existsSync(tracker.created[0])).to.equal(false);
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
