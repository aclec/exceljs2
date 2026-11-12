// Regressions of the streaming writer / reader found while typing src/stream (TypeScript port).
const { PassThrough, Readable } = require("stream");
const { baseWorkbookBuffer, tamper, readPart } = require("../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

// writes a streamed workbook to memory
async function streamedBuffer(build) {
    const stream = new PassThrough();
    const chunks = [];
    stream.on("data", (chunk) => chunks.push(chunk));
    const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ stream });
    await build(wb);
    await wb.commit();
    return Buffer.concat(chunks);
}

describe("stream I/O regressions", () => {
    describe("WorkbookWriter", () => {
        it("numbers the shared formulas of the worksheet once (si is not restarted on each row)", async () => {
            const buffer = await streamedBuffer((wb) => {
                const ws = wb.addWorksheet("F");
                ws.getCell("A1").value = { formula: "ROW()", shareType: "shared", ref: "A1:A2", result: 1 };
                ws.getCell("A2").value = { sharedFormula: "A1", result: 2 };
                ws.getCell("B3").value = { formula: "COLUMN()", shareType: "shared", ref: "B3:B4", result: 2 };
                ws.getCell("B4").value = { sharedFormula: "B3", result: 2 };
                ws.commit();
            });

            const xml = await readPart(buffer, "xl/worksheets/sheet1.xml");
            const si = (ref) => xml.match(new RegExp(`<c r="${ref}"[^>]*><f [^>]*si="(\\d+)"`))[1];
            expect(si("A1")).to.equal(si("A2"));
            expect(si("B3")).to.equal(si("B4"));
            expect(si("A1")).not.to.equal(si("B3"));

            const wb = new ExcelJS.Workbook();
            await wb.xlsx.load(buffer);
            const ws = wb.getWorksheet("F");
            expect(ws.getCell("A2").value).to.deep.equal({ sharedFormula: "A1", result: 2 });
            expect(ws.getCell("B4").value).to.deep.equal({ sharedFormula: "B3", result: 2 });
        });

        it("keeps numbering the shared formulas across committed rows", async () => {
            const buffer = await streamedBuffer((wb) => {
                const ws = wb.addWorksheet("F");
                for (let r = 1; r <= 3; r++) {
                    ws.getCell(`A${r}`).value = { formula: `ROW()*${r}`, shareType: "shared", ref: `A${r}`, result: r };
                    ws.getRow(r).commit();
                }
                ws.commit();
            });
            const xml = await readPart(buffer, "xl/worksheets/sheet1.xml");
            expect([...xml.matchAll(/si="(\d+)"/g)].map((m) => m[1])).to.deep.equal(["0", "1", "2"]);
        });
    });

    describe("WorkbookReader", () => {
        it('reads a cell with a hyperlink listed before <sheetData> (hyperlinks: "cache") without losing its value', async () => {
            const base = await baseWorkbookBuffer((wb, ws) => {
                ws.getCell("A2").value = { text: "link", hyperlink: "http://example.com/" };
            });
            // <hyperlinks> before <sheetData>: out of the ECMA-376 order, but the reader must not fail on it
            const buffer = await tamper(base, {
                "xl/worksheets/sheet1.xml": (xml) => {
                    const hyperlinks = xml.match(/<hyperlinks>.*<\/hyperlinks>/)[0];
                    return xml.replace(hyperlinks, "").replace("<sheetData>", `${hyperlinks}<sheetData>`);
                },
            });

            const reader = new ExcelJS.stream.xlsx.WorkbookReader(Readable.from([buffer]), { hyperlinks: "cache" });
            const values = [];
            let worksheet;
            for await (const ws of reader) {
                worksheet = ws;
                for await (const row of ws) {
                    values.push(row.getCell(1).value);
                }
            }
            expect(values).to.deep.equal(["hello", "link"]);
            expect(worksheet.hyperlinks.A2).to.include({ ref: "A2", rId: "rId1" });
        });

        it("does not report a row for a <row> outside <sheetData>", async () => {
            const buffer = await tamper(await baseWorkbookBuffer(), {
                "xl/worksheets/sheet1.xml": (xml) => xml.replace("</sheetData>", '</sheetData><x><row r="9"/></x>'),
            });
            const rows = [];
            for await (const ws of new ExcelJS.stream.xlsx.WorkbookReader(Readable.from([buffer]))) {
                for await (const row of ws) {
                    rows.push(row);
                }
            }
            expect(rows.map((row) => row && row.number)).to.deep.equal([1]);
        });
    });

    describe("HyperlinkReader", () => {
        const HyperlinkReader = verquire("stream/xlsx/hyperlink-reader");

        it('counts and iterates the cached relationships (hyperlinks: "cache")', async () => {
            const xml =
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="http://a/" TargetMode="External"/>' +
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>' +
                '<Relationship Id="__proto__" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="http://b/" TargetMode="External"/>' +
                "</Relationships>";
            const reader = new HyperlinkReader({ workbook: {}, id: "1", iterator: [xml], options: { hyperlinks: "cache" } });
            expect(reader.count).to.equal(0);
            await reader.read();
            expect(reader.count).to.equal(2);
            const targets = [];
            reader.each((relationship) => targets.push(relationship.target));
            expect(targets).to.deep.equal(["http://a/", "http://b/"]);
            expect({}.target).to.equal(undefined);
        });
    });
});
