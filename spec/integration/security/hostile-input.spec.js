// Regression tests for hostile xlsx/csv input: each case used to crash the process
// (heap OOM), block the event loop, or pollute Object.prototype.
const { Readable } = require("stream");
const { baseWorkbookBuffer, tamper, readPart } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

const SHEET = "xl/worksheets/sheet1.xml";
const afterSheetData = (xml) => (fragment) => xml.replace("</sheetData>", `</sheetData>${fragment}`);

async function load(buffer) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    return wb;
}

async function timed(fn) {
    const start = Date.now();
    const result = await fn();
    return { result, ms: Date.now() - start };
}

// a workbook with sheets Sheet1..SheetN, and patches applying fragment(i) after the sheetData of each
async function sheetsBuffer(count) {
    return baseWorkbookBuffer((wb) => {
        for (let i = 2; i <= count; i++) {
            wb.addWorksheet(`Sheet${i}`).getCell("A1").value = i;
        }
    });
}
function afterEachSheetData(count, fragment) {
    const patches = {};
    for (let i = 1; i <= count; i++) {
        patches[`xl/worksheets/sheet${i}.xml`] = (xml) => afterSheetData(xml)(fragment(i));
    }
    return patches;
}

describe("hostile input", function () {
    this.timeout(20000);

    describe("merge cells", () => {
        it("rejects a merge covering the whole grid instead of exhausting memory", async () => {
            const buffer = await tamper(await baseWorkbookBuffer(), {
                [SHEET]: (xml) => afterSheetData(xml)('<mergeCells count="1"><mergeCell ref="A3:XFD1048576"/></mergeCells>'),
            });
            let error;
            await load(buffer).catch((e) => {
                error = e;
            });
            expect(error).to.be.an("error");
            expect(error.message).to.match(/merged cells/);
        });

        it("rejects references beyond the last Excel row", async () => {
            const buffer = await tamper(await baseWorkbookBuffer(), {
                [SHEET]: (xml) => afterSheetData(xml)('<mergeCells count="1"><mergeCell ref="A1:B99999999999"/></mergeCells>'),
            });
            let error;
            await load(buffer).catch((e) => {
                error = e;
            });
            expect(error).to.be.an("error");
            expect(error.message).to.match(/Out of bounds/);
        });

        it("caps the merged cells of the whole workbook, not of each worksheet", async () => {
            // each worksheet is under the cap of 1048576 merged cells, the workbook is not
            const buffer = await tamper(
                await sheetsBuffer(2),
                afterEachSheetData(2, () => '<mergeCells count="1"><mergeCell ref="B1:B524289"/></mergeCells>')
            );
            const { result: error, ms } = await timed(() => load(buffer).catch((e) => e));
            expect(error).to.be.an("error");
            expect(error.message).to.match(/merged cells/);
            expect(ms).to.be.below(2000); // rejected before the cells are created
        });

        it("still loads merges spread over several worksheets", async () => {
            const buffer = await tamper(
                await sheetsBuffer(3),
                afterEachSheetData(3, () => '<mergeCells count="1"><mergeCell ref="B2:K1001"/></mergeCells>')
            );
            const wb = await load(buffer);
            wb.eachSheet((ws) => expect(ws.getCell("K1001").master.address).to.equal("B2"));
        });

        it("loads many small merges in linear time", async () => {
            const merges = [];
            for (let row = 1; row <= 20000; row++) {
                merges.push(`<mergeCell ref="B${row}:C${row}"/>`);
            }
            const buffer = await tamper(await baseWorkbookBuffer(), {
                [SHEET]: (xml) => afterSheetData(xml)(`<mergeCells count="${merges.length}">${merges.join("")}</mergeCells>`),
            });
            const { result: wb, ms } = await timed(() => load(buffer));
            expect(wb.getWorksheet("Sheet1").getCell("C20000").isMerged).to.equal(true);
            expect(ms).to.be.below(5000);
        });

        it("still rejects overlapping merges without modifying the sheet", () => {
            const wb = new ExcelJS.Workbook();
            const ws = wb.addWorksheet("s");
            ws.mergeCells("B2:C3");
            expect(() => ws.mergeCells("A1:B2")).to.throw(/already merged/);
            expect(() => ws.mergeCells("C3:D4")).to.throw(/already merged/);
            expect(ws.getCell("A1").isMerged).to.equal(false);
            expect(ws.getCell("D4").isMerged).to.equal(false);
            ws.unMergeCells("B2");
            ws.mergeCells("A1:B2");
            expect(ws.getCell("B2").master.address).to.equal("A1");
        });
    });

    describe("data validations", () => {
        it("keeps a huge sqref as a range instead of expanding it per cell", async () => {
            const buffer = await tamper(await baseWorkbookBuffer(), {
                [SHEET]: (xml) =>
                    afterSheetData(xml)(
                        '<dataValidations count="1"><dataValidation type="list" allowBlank="1" sqref="B1:XFD1048576">' +
                            '<formula1>"a,b"</formula1></dataValidation></dataValidations>'
                    ),
            });
            const { result: wb, ms } = await timed(() => load(buffer));
            expect(ms).to.be.below(5000);
            const ws = wb.getWorksheet("Sheet1");
            expect(ws.getCell("C5").dataValidation).to.deep.include({ type: "list", formulae: ['"a,b"'] });
            expect(ws.getCell("A5").dataValidation).to.equal(undefined);

            // the range survives a write / read round trip
            const out = await wb.xlsx.writeBuffer();
            expect(await readPart(out, SHEET)).to.include('sqref="B1:XFD1048576"');
        });

        it("still expands small ranges per cell", async () => {
            const buffer = await tamper(await baseWorkbookBuffer(), {
                [SHEET]: (xml) =>
                    afterSheetData(xml)(
                        '<dataValidations count="1"><dataValidation type="whole" operator="greaterThan" sqref="A2:B3 D4">' +
                            "<formula1>5</formula1></dataValidation></dataValidations>"
                    ),
            });
            const ws = (await load(buffer)).getWorksheet("Sheet1");
            expect(Object.keys(ws.dataValidations.model).sort()).to.deep.equal(["A2", "A3", "B2", "B3", "D4"]);
        });

        it("ignores sqref tokens that are not cell references", async () => {
            const buffer = await tamper(await baseWorkbookBuffer(), {
                [SHEET]: (xml) =>
                    afterSheetData(xml)(
                        '<dataValidations count="1"><dataValidation type="list" sqref="__proto__ constructor A1">' +
                            '<formula1>"a"</formula1></dataValidation></dataValidations>'
                    ),
            });
            const ws = (await load(buffer)).getWorksheet("Sheet1");
            expect(Object.keys(ws.dataValidations.model)).to.deep.equal(["A1"]);
            expect(Object.getPrototypeOf(ws.dataValidations.model)).to.equal(Object.prototype);
        });

        it("shares the expansion budget between the worksheets of a workbook", async () => {
            const buffer = await tamper(
                await sheetsBuffer(3),
                afterEachSheetData(
                    3,
                    () =>
                        '<dataValidations count="1"><dataValidation type="list" sqref="A1:CV1000">' +
                        '<formula1>"a"</formula1></dataValidation></dataValidations>'
                )
            );
            const wb = await load(buffer);
            // 100,000 cells: the first worksheet uses the whole budget, the others keep a range key
            expect(Object.keys(wb.getWorksheet("Sheet1").dataValidations.model)).to.have.length(100000);
            expect(Object.keys(wb.getWorksheet("Sheet2").dataValidations.model)).to.deep.equal(["A1:CV1000"]);
            expect(Object.keys(wb.getWorksheet("Sheet3").dataValidations.model)).to.deep.equal(["A1:CV1000"]);
            wb.eachSheet((ws) => {
                expect(ws.getCell("CV1000").dataValidation).to.deep.include({ type: "list" });
                expect(ws.getCell("CW1000").dataValidation).to.equal(undefined);
            });
        });

        it("keeps whole-column validations on several worksheets", async () => {
            const buffer = await tamper(
                await sheetsBuffer(3),
                afterEachSheetData(
                    3,
                    (i) =>
                        `<dataValidations count="1"><dataValidation type="whole" operator="greaterThan" sqref="C1:C1048576">` +
                        `<formula1>${i}</formula1></dataValidation></dataValidations>`
                )
            );
            const wb = await load(buffer);
            const out = await load(await wb.xlsx.writeBuffer());
            [1, 2, 3].forEach((i) => {
                expect(out.getWorksheet(`Sheet${i}`).getCell("C77").dataValidation).to.deep.include({ formulae: [i] });
            });
        });

        it("keeps removals and overrides of cells inside a range key after a write", async () => {
            const buffer = await tamper(await baseWorkbookBuffer(), {
                [SHEET]: (xml) =>
                    afterSheetData(xml)(
                        '<dataValidations count="1"><dataValidation type="list" sqref="B1:C1048576">' +
                            '<formula1>"a,b"</formula1></dataValidation></dataValidations>'
                    ),
            });
            const ws = (await load(buffer)).getWorksheet("Sheet1");
            ws.dataValidations.remove("B5");
            ws.getCell("C7").dataValidation = { type: "whole", operator: "greaterThan", formulae: [3] };
            const out = await ws.workbook.xlsx.writeBuffer();
            const ws2 = (await load(out)).getWorksheet("Sheet1");
            expect(ws2.getCell("B5").dataValidation).to.equal(undefined);
            expect(ws2.getCell("C7").dataValidation).to.deep.include({ type: "whole", formulae: [3] });
            ["B4", "B6", "C5", "B7", "C6", "C8", "C1048576"].forEach((address) => {
                expect(ws2.getCell(address).dataValidation, address).to.deep.include({ type: "list" });
            });
            // no overlapping sqref is written
            const sqrefs = [...(await readPart(out, SHEET)).matchAll(/sqref="([^"]+)"/g)].map((m) => m[1]);
            expect(sqrefs).to.have.members(["B1:C4", "C5", "B6:C6", "B7", "B8:C1048576", "C7"]);
        });

        it("does not write validations that were removed", async () => {
            const wb = new ExcelJS.Workbook();
            const ws = wb.addWorksheet("s");
            ws.getCell("A1").dataValidation = { type: "list", formulae: ['"x"'] };
            ws.getCell("A2").dataValidation = { type: "list", formulae: ['"x"'] };
            ws.dataValidations.remove("A2");
            const xml = await readPart(await wb.xlsx.writeBuffer(), SHEET);
            expect(xml).to.include('<dataValidations count="1">');
            expect(xml).to.include('sqref="A1"');
        });
    });

    describe("row numbers", () => {
        const withRow = (r) => ({ [SHEET]: (xml) => xml.replace("</sheetData>", `<row r="${r}"/></sheetData>`) });

        it("rejects an empty row beyond the last Excel row instead of creating a sparse row array", async () => {
            for (const r of ["1048577", "4294967294", "0", "-3"]) {
                const buffer = await tamper(await baseWorkbookBuffer(), withRow(r));
                const error = await load(buffer).catch((e) => e);
                expect(error, r).to.be.an("error");
                expect(error.message).to.match(/Out of bounds. Excel supports rows from 1 to 1048576/);
            }
        });

        it("rejects it in the streaming reader too", async () => {
            const buffer = await tamper(await baseWorkbookBuffer(), withRow("4294967294"));
            const reader = new ExcelJS.stream.xlsx.WorkbookReader(Readable.from([buffer]));
            const read = async () => {
                for await (const ws of reader) {
                    for await (const row of ws) {
                        expect(row.number).to.equal(1);
                    }
                }
            };
            const error = await read().catch((e) => e);
            expect(error).to.be.an("error");
            expect(error.message).to.match(/Out of bounds. Excel supports rows from 1 to 1048576/);
        });

        it("still reads the last Excel row", async () => {
            const buffer = await tamper(await baseWorkbookBuffer(), {
                [SHEET]: (xml) => xml.replace("</sheetData>", '<row r="1048576"><c r="A1048576"><v>5</v></c></row></sheetData>'),
            });
            const ws = (await load(buffer)).getWorksheet("Sheet1");
            expect(ws.getCell("A1048576").value).to.equal(5);
        });
    });

    describe("sheet ids", () => {
        const withSheetIds = (...ids) => ({
            "xl/workbook.xml": (xml) => {
                let i = 0;
                return xml.replace(/sheetId="\d+"/g, () => `sheetId="${ids[i++]}"`);
            },
        });
        const twoSheets = () =>
            baseWorkbookBuffer((wb) => {
                wb.addWorksheet("Sheet2").getCell("A1").value = "two";
            });

        it("renumbers a huge sheetId instead of creating a sparse worksheet array", async () => {
            // (up to 4294967294 in a file: 10 million is enough to show the cost and stays bounded)
            const buffer = await tamper(await baseWorkbookBuffer(), withSheetIds("10000000"));
            const { result: wb, ms } = await timed(async () => {
                const book = await load(buffer);
                book.getWorksheet("Sheet1");
                book.addWorksheet("added");
                await book.xlsx.writeBuffer();
                return book;
            });
            expect(ms).to.be.below(500);
            expect(wb.worksheets.map((ws) => [ws.id, ws.name])).to.deep.equal([
                [1, "Sheet1"],
                [2, "added"],
            ]);
        });

        it("keeps ids up to 65535 and renumbers invalid or duplicate ones to free ids", async () => {
            const check = async (ids, expected) => {
                const wb = await load(await tamper(await twoSheets(), withSheetIds(...ids)));
                expect(
                    wb.worksheets.map((ws) => [ws.id, ws.name]),
                    ids.join()
                ).to.deep.equal(expected);
                expect(wb.getWorksheet("Sheet2").getCell("A1").value).to.equal("two");
            };
            await check(
                ["7", "65535"],
                [
                    [7, "Sheet1"],
                    [65535, "Sheet2"],
                ]
            );
            // a duplicate id used to replace the first worksheet with the second
            await check(
                ["3", "3"],
                [
                    [3, "Sheet1"],
                    [1, "Sheet2"],
                ]
            );
            await check(
                ["65536", "1"],
                [
                    [2, "Sheet1"],
                    [1, "Sheet2"],
                ]
            );
            await check(
                ["0", "x"],
                [
                    [1, "Sheet1"],
                    [2, "Sheet2"],
                ]
            );
        });
    });

    describe("number formats", () => {
        // a numeric cell styled with the custom format formatCode
        const withNumFmt = (formatCode, cells = 1) => ({
            "xl/styles.xml": (xml) =>
                xml
                    .replace(/<numFmts[^>]*>/, (open) => `${open}<numFmt numFmtId="200" formatCode="${formatCode}"/>`)
                    .replace(/<cellXfs count="(\d+)">(.*?)<\/cellXfs>/s, (m, count, xfs) => {
                        const xf = '<xf numFmtId="200" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>';
                        return `<cellXfs count="${+count + 1}">${xfs}${xf}</cellXfs>`;
                    }),
            [SHEET]: (xml) => {
                const rows = Array.from({ length: cells }, (_, i) => `<row r="${i + 2}"><c r="A${i + 2}" s="XF"><v>1</v></c></row>`);
                return xml.replace("</sheetData>", `${rows.join("")}</sheetData>`);
            },
        });
        const build = async (formatCode, cells) => {
            // a base whose styles.xml has a <numFmts> element and two cell formats
            const base = await baseWorkbookBuffer((wb, ws) => {
                ws.getCell("B1").numFmt = "0.000";
            });
            const styles = await readPart(base, "xl/styles.xml");
            const xfIndex = (styles.match(/<cellXfs[^>]*>(.*?)<\/cellXfs>/s)[1].match(/<xf /g) || []).length;
            const patches = withNumFmt(formatCode, cells);
            const sheetPatch = patches[SHEET];
            patches[SHEET] = (xml) => sheetPatch(xml).replace(/s="XF"/g, `s="${xfIndex}"`);
            return tamper(base, patches);
        };

        it("keeps custom formats up to the Excel limit of 255 characters", async () => {
            const formatCode = `yyyy-mm-dd${" ".repeat(245)}`;
            const ws = (await load(await build(formatCode))).getWorksheet("Sheet1");
            expect(ws.getCell("A2").numFmt).to.equal(formatCode);
            expect(ws.getCell("A2").value).to.be.an.instanceof(Date);
        });

        it("ignores custom formats longer than 255 characters", async () => {
            const { result: wb, ms } = await timed(async () => load(await build("[".repeat(20000), 50)));
            expect(ms).to.be.below(2000);
            const ws = wb.getWorksheet("Sheet1");
            expect(ws.getCell("A2").numFmt).to.equal(undefined);
            expect(ws.getCell("A51").value).to.equal(1);
        });
    });

    it("clamps <col max> to the last Excel column", async () => {
        const buffer = await tamper(await baseWorkbookBuffer(), {
            [SHEET]: (xml) =>
                xml.replace("<sheetData>", '<cols><col min="1" max="2000000000" width="12" customWidth="1"/></cols><sheetData>'),
        });
        const ws = (await load(buffer)).getWorksheet("Sheet1");
        expect(ws.columns.length).to.equal(16384);
        expect(ws.getColumn(16384).width).to.equal(12);
    });

    describe("defined names", () => {
        const withDefinedNames = (names) => ({
            "xl/workbook.xml": (xml) => xml.replace("</sheets>", `</sheets><definedNames>${names}</definedNames>`),
        });

        it("keeps a huge defined-name range as a rectangle", async () => {
            const buffer = await tamper(
                await baseWorkbookBuffer(),
                withDefinedNames('<definedName name="everything">Sheet1!$A$1:$XFD$1048576</definedName>')
            );
            const { result: wb, ms } = await timed(() => load(buffer));
            expect(ms).to.be.below(5000);
            expect(wb.getWorksheet("Sheet1").getCell("Z99").names).to.deep.equal(["everything"]);
            expect(wb.definedNames.model).to.deep.equal([{ name: "everything", ranges: ["Sheet1!$A$1:$XFD$1048576"] }]);
        });

        it("shares the expansion budget between the names of a workbook", async () => {
            const names = [0, 1, 2].map((i) => `<definedName name="n${i}">Sheet1!$A$1:$CV$1000</definedName>`).join("");
            const wb = await load(await tamper(await baseWorkbookBuffer(), withDefinedNames(names)));
            // 100,000 cells: the first name uses the whole budget, the others are kept as rectangles
            expect(Object.keys(wb.definedNames.rangeMap)).to.deep.equal(["n1", "n2"]);
            expect(wb.getWorksheet("Sheet1").getCell("CV1000").names).to.deep.equal(["n0", "n1", "n2"]);
            expect(wb.definedNames.model).to.deep.equal(["n0", "n1", "n2"].map((name) => ({ name, ranges: ["Sheet1!$A$1:$CV$1000"] })));
        });

        it("lists the names of a workbook with many large ranges in linear time", async () => {
            const names = Array.from({ length: 40000 }, (_, i) => `<definedName name="n${i}">Sheet1!$A$1:$A$1048576</definedName>`);
            const wb = await load(await tamper(await baseWorkbookBuffer(), withDefinedNames(names.join(""))));
            const { result: model, ms } = await timed(() => wb.definedNames.model);
            expect(model).to.have.length(40000);
            expect(ms).to.be.below(300); // was 640 ms: names.includes for every name
            const { result: cellNames, ms: cellMs } = await timed(() => wb.getWorksheet("Sheet1").getCell("A5").names);
            expect(cellNames).to.have.length(40000);
            expect(cellMs).to.be.below(300);
        });

        it("shifts and removes large ranges with row splices", () => {
            const wb = new ExcelJS.Workbook();
            const ws = wb.addWorksheet("s");
            wb.definedNames.add("s!$A$1:$B$100000", "big");
            ws.spliceRows(1, 2);
            expect(wb.definedNames.getRanges("big").ranges).to.deep.equal(["s!$A$1:$B$99998"]);
            ws.spliceRows(1, 0, [], []);
            expect(wb.definedNames.getRanges("big").ranges).to.deep.equal(["s!$A$3:$B$100000"]);
            ws.getCell("A50").removeName("big");
            expect(wb.definedNames.getRanges("big").ranges).to.deep.equal([]);
        });

        it('does not pollute Object.prototype through a "__proto__" sheet name', async () => {
            const buffer = await tamper(
                await baseWorkbookBuffer(),
                withDefinedNames('<definedName name="a">__proto__!$A$7</definedName><definedName name="b">\'__proto__\'!$B$9</definedName>')
            );
            await load(buffer);
            expect({}[7]).to.equal(undefined);
            expect({}[9]).to.equal(undefined);
            expect(7 in {}).to.equal(false);
        });
    });

    it("parses a drawing relationship with a pathological Target in linear time", async () => {
        const target = `..${"/media/".repeat(20000)}`;
        const base = await baseWorkbookBuffer((wb, ws) => {
            const imageId = wb.addImage({ filename: "./spec/integration/data/image.png", extension: "png" });
            ws.addImage(imageId, "B2:C3");
        });
        const buffer = await tamper(base, {
            "xl/drawings/_rels/drawing1.xml.rels": (xml) => xml.replace("../media/image1.png", target),
        });
        const { ms } = await timed(() => load(buffer));
        expect(ms).to.be.below(2000);
    });

    it("rejects formulas longer than the Excel limit instead of scanning them quadratically", () => {
        const { slideFormula } = verquire("utils/shared-formula");
        expect(() => slideFormula("a-".repeat(5000), "A1", "A2")).to.throw(/Excel limit/);
        expect(slideFormula("A1+$B$2+Sheet1!C3", "A1", "B2")).to.equal("B2+$B$2+Sheet1!D4");
    });

    it('does not merge "__proto__" keys from app-supplied objects (cell notes)', async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet("s");
        ws.getCell("A1").note = JSON.parse('{"texts":[{"text":"hi"}],"__proto__":{"polluted":"yes"}}');
        await wb.xlsx.writeBuffer();
        expect({}.polluted).to.equal(undefined);
    });

    it("reads CSV cells named like Object.prototype members as plain text", async () => {
        const wb = new ExcelJS.Workbook();
        const ws = await wb.csv.read(Readable.from(["__proto__,constructor,toString,true,#N/A\n"]));
        expect(ws.getRow(1).values.slice(1)).to.deep.equal(["__proto__", "constructor", "toString", true, { error: "#N/A" }]);
        await wb.xlsx.writeBuffer(); // used to throw on the Object.prototype cell value
    });

    it("strips characters that are not allowed in XML 1.0 when writing", async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet("s");
        ws.getCell("A1").value = "a￾b￿c\uD800d";
        ws.getCell("A2").value = "emoji 😀 stays";
        const wb2 = await load(await wb.xlsx.writeBuffer());
        const ws2 = wb2.getWorksheet("s");
        expect(ws2.getCell("A1").value).to.equal("abcd");
        expect(ws2.getCell("A2").value).to.equal("emoji 😀 stays");
    });

    it("escapes pivot table names and values", async () => {
        const wb = new ExcelJS.Workbook();
        const source = wb.addWorksheet("src");
        const evil = 'x" injected="1"/><evil/><s v="';
        source.addRows([
            ["A", "B", "C"],
            [evil, "b1", 5],
            ["a2", "b2", 7],
        ]);
        const pivot = wb.addWorksheet("pivot");
        pivot.addPivotTable({
            sourceSheet: source,
            rows: ["A"],
            columns: ["B"],
            values: ["C"],
            metric: "sum",
        });
        const out = await wb.xlsx.writeBuffer();
        const definition = await readPart(out, "xl/pivotCache/pivotCacheDefinition1.xml");
        expect(definition).not.to.include("<evil/>");
        expect(definition).to.include("x&quot; injected=&quot;1&quot;/&gt;&lt;evil/&gt;");
        await load(out); // output is well-formed
    });
});
