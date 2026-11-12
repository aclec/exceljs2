// Every id, name, rId, ref... read from an xlsx part may be used as an object key. Feeding the
// names of Object.prototype members must neither pollute the global prototypes nor leave the
// library in a state where a later, unrelated write fails.
const { Readable } = require("stream");
const { strFromU8 } = require("fflate");
const { baseWorkbookBuffer, tamper, readPart, unzipParts } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

const KEYS = ["__proto__", "constructor", "prototype", "toString"];
const PROTOTYPES = [Object.prototype, Array.prototype, Function.prototype, String.prototype];

function snapshotPrototypes() {
    return PROTOTYPES.map((proto) => Object.getOwnPropertyNames(proto).sort().join(","));
}
const PRISTINE = snapshotPrototypes();

// removes polluted members, so that one finding does not mask the others
function restorePrototypes() {
    PROTOTYPES.forEach((proto, i) => {
        const known = new Set(PRISTINE[i].split(","));
        Object.getOwnPropertyNames(proto)
            .filter((name) => !known.has(name))
            .forEach((name) => delete proto[name]);
    });
}

// A workbook that uses as many id/name/rId keyed parts as possible.
function buildRichWorkbook(wb, ws) {
    ws.columns = [
        { header: "A", key: "a", width: 10 },
        { header: "B", key: "b" },
    ];
    for (let i = 0; i < 5; i++) {
        ws.addRow({ a: i, b: `x${i}` });
    }
    ws.getCell("C1").value = { formula: "A1*2", result: 0, shareType: "shared", ref: "C1:C3" };
    ws.getCell("C2").value = { sharedFormula: "C1", result: 0 };
    ws.getCell("D1").value = { text: "link", hyperlink: "http://example.com/" };
    ws.getCell("D2").value = { richText: [{ text: "r", font: { bold: true } }, { text: "s" }] };
    ws.getCell("D3").note = "hello";
    ws.getCell("A3").numFmt = "yyyy-mm-dd";
    ws.mergeCells("E1:F2");
    ws.getCell("A2").dataValidation = { type: "list", formulae: ['"a,b"'] };
    ws.addConditionalFormatting({
        ref: "A1:A5",
        rules: [
            {
                type: "expression",
                formulae: ["A1>2"],
                style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFF0000" } } },
            },
            { type: "dataBar", cfvo: [{ type: "min" }, { type: "max" }], color: { argb: "FF00FF00" } },
        ],
    });
    ws.autoFilter = "A1:B5";
    const ws2 = wb.addWorksheet("T");
    ws2.addTable({
        name: "Tbl",
        ref: "A1",
        headerRow: true,
        columns: [{ name: "x", filterButton: true }, { name: "y" }],
        rows: [
            [1, 2],
            [3, 4],
        ],
    });
    ws2.getCell("D1").value = { text: "internal", hyperlink: "#Sheet1!A1" };
    wb.definedNames.add("Sheet1!$A$1:$B$3", "nm");
    const imageId = wb.addImage({ filename: "./spec/integration/data/image.png", extension: "png" });
    ws.addImage(imageId, "H2:I4");
    ws2.addBackgroundImage(imageId);
}

async function partsOf(buffer) {
    const parts = unzipParts(buffer);
    const texts = {};
    for (const name of Object.keys(parts).filter((n) => /\.(xml|rels|vml)$/.test(n))) {
        texts[name] = strFromU8(parts[name]);
    }
    return texts;
}

// [part, attribute] pairs: every attribute name used in every part (namespaces excluded)
function attributeCases(texts) {
    const cases = [];
    for (const [part, xml] of Object.entries(texts)) {
        const attributes = new Set([...xml.matchAll(/\s([\w:]+)="[^"]*"/g)].map((m) => m[1]).filter((n) => !n.startsWith("xmlns")));
        for (const attribute of attributes) {
            cases.push([part, attribute]);
        }
    }
    return cases;
}

const replaceAttribute = (attribute, value) => (xml) => xml.replace(new RegExp(`(\\s${attribute}=")[^"]*(")`, "g"), `$1${value}$2`);

// Resolves with "rejected" instead of throwing: hostile input may be refused, but must settle.
const settle = (promise) =>
    promise.then(
        () => "resolved",
        () => "rejected"
    );

function attempt(fn) {
    try {
        return fn();
    } catch (error) {
        return error;
    }
}

function isPolluted() {
    return snapshotPrototypes().join("|") !== PRISTINE.join("|");
}

// Loads (or rejects), reads every cell, then writes (or rejects). Returns the step after which
// the prototypes were polluted, checked before the next step can be affected by the pollution.
async function readAll(buffer) {
    const wb = new ExcelJS.Workbook();
    const loaded = await settle(wb.xlsx.load(buffer));
    if (isPolluted()) {
        return "load";
    }
    if (loaded === "rejected") {
        return undefined;
    }
    wb.eachSheet((ws) =>
        ws.eachRow((row) =>
            row.eachCell((cell) => {
                for (const getter of ["value", "formula", "dataValidation", "names", "text"]) {
                    // a getter may throw on inconsistent input: that is caught by the caller,
                    // only pollution and unsettled promises matter here
                    attempt(() => cell[getter]);
                }
            })
        )
    );
    if (isPolluted()) {
        return "cell getters";
    }
    await settle(wb.xlsx.writeBuffer());
    return isPolluted() ? "write" : undefined;
}

async function streamAll(buffer) {
    const reader = new ExcelJS.stream.xlsx.WorkbookReader(Readable.from([buffer]), {
        sharedStrings: "cache",
        hyperlinks: "cache",
        styles: "cache",
    });
    const iterate = async () => {
        for await (const ws of reader) {
            for await (const row of ws) {
                row.eachCell((cell) => [attempt(() => cell.value), attempt(() => cell.text)]);
            }
        }
    };
    await settle(iterate());
    return isPolluted() ? "stream" : undefined;
}

async function writeUnrelatedWorkbook() {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet("ok").getCell("A1").value = "hello";
    const back = new ExcelJS.Workbook();
    await back.xlsx.load(await wb.xlsx.writeBuffer());
    return back.getWorksheet("ok").getCell("A1").value;
}

describe("prototype-named keys in xlsx parts", function () {
    this.timeout(120000);

    let base;
    let texts;
    before(async () => {
        base = await baseWorkbookBuffer(buildRichWorkbook);
        texts = await partsOf(base);
    });
    afterEach(restorePrototypes);

    it('does not pollute Object.prototype through an x14 conditional formatting rule id="__proto__"', async () => {
        const ext =
            '<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
            '<x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">' +
            '<x14:cfRule type="attacker-type" priority="7" id="__proto__" operator="attacker-op"><xm:f>attacker-formula</xm:f></x14:cfRule>' +
            "<xm:sqref>A1</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>";
        const buffer = await tamper(await baseWorkbookBuffer(), {
            "xl/worksheets/sheet1.xml": (xml) =>
                xml
                    .replace(
                        "</sheetData>",
                        '</sheetData><conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>1</formula></cfRule></conditionalFormatting>'
                    )
                    .replace("</worksheet>", `${ext}</worksheet>`),
        });
        const before = snapshotPrototypes();
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buffer);
        expect(snapshotPrototypes()).to.deep.equal(before);
        expect({}.type).to.equal(undefined);
        // the x14 rule is kept as a rule of its own
        const rules = wb.getWorksheet("Sheet1").conditionalFormattings.flatMap((cf) => cf.rules);
        expect(rules.map((rule) => rule.type)).to.deep.equal(["expression", "attacker-type"]);
        expect(await writeUnrelatedWorkbook()).to.equal("hello");
    });

    describe("treats ids named like Object.prototype members as unknown", () => {
        it("a sheet relationship id", async () => {
            const buffer = await tamper(await baseWorkbookBuffer(), {
                "xl/workbook.xml": (xml) => xml.replace(/r:id="rId\d+"/, 'r:id="constructor"'),
            });
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.load(buffer); // used to throw a TypeError on the inherited constructor
            expect(wb.worksheets).to.have.length(0);
        });

        it("a shared formula index", async () => {
            const buffer = await tamper(await baseWorkbookBuffer(), {
                "xl/worksheets/sheet1.xml": (xml) =>
                    xml.replace("</sheetData>", '<row r="5"><c r="A5"><f t="shared" si="toString"/><v>1</v></c></row></sheetData>'),
            });
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.load(buffer);
            expect(wb.getWorksheet("Sheet1").getCell("A5").value).to.deep.equal({ result: 1 });
            await wb.xlsx.writeBuffer(); // used to throw on a missing shared formula master
        });

        it('a hyperlink ref="__proto__" or r:id="__proto__"', async () => {
            const base = await baseWorkbookBuffer((wb, ws) => {
                ws.getCell("B2").value = { text: "x", hyperlink: "http://example.com/" };
                ws.getCell("B3").value = { text: "y", hyperlink: "http://example.org/" };
            });
            const buffer = await tamper(base, {
                "xl/worksheets/sheet1.xml": (xml) =>
                    xml
                        .replace('<hyperlink ref="B2"', '<hyperlink ref="__proto__"')
                        .replace(
                            "</hyperlinks>",
                            '<hyperlink ref="A1" r:id="__proto__"/><hyperlink ref="constructor" location="Sheet1!A1"/></hyperlinks>'
                        ),
            });
            const before = snapshotPrototypes();
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.load(buffer);
            expect(snapshotPrototypes()).to.deep.equal(before);
            const ws = wb.getWorksheet("Sheet1");
            expect(ws.getCell("A1").value).to.equal("hello");
            expect(ws.getCell("B2").value).to.equal("x");
            expect(ws.getCell("B3").value).to.deep.equal({ text: "y", hyperlink: "http://example.org/" });
            const out = await wb.xlsx.writeBuffer();
            expect(await readPart(out, "xl/worksheets/sheet1.xml")).not.to.include("__proto__");
        });

        it("a background image target", async () => {
            const base = await baseWorkbookBuffer((wb, ws) => {
                ws.addBackgroundImage(wb.addImage({ filename: "./spec/integration/data/image.png", extension: "png" }));
            });
            const buffer = await tamper(base, {
                "xl/worksheets/_rels/sheet1.xml.rels": (xml) => xml.replace(/\.\.\/media\/[^"]+/, "../media/constructor"),
            });
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.load(buffer);
            expect(wb.getWorksheet("Sheet1").getBackgroundImageId()).to.equal(undefined);
            await wb.xlsx.writeBuffer(); // used to reject on the inherited Object constructor
        });

        it("a number format", async () => {
            const wb = new ExcelJS.Workbook();
            const ws = wb.addWorksheet("s");
            KEYS.forEach((key, i) => {
                ws.getCell(1, i + 1).value = i;
                ws.getCell(1, i + 1).numFmt = key;
            });
            const out = await wb.xlsx.writeBuffer();
            // used to write numFmtId="function Object() { [native code] }"
            expect(await readPart(out, "xl/styles.xml")).not.to.include("native code");
            const back = new ExcelJS.Workbook();
            await back.xlsx.load(out);
            expect(KEYS.map((key, i) => back.getWorksheet("s").getCell(1, i + 1).numFmt)).to.deep.equal(KEYS);
        });

        it("a cell reference looked up in the streamed hyperlinks", async () => {
            const base = await baseWorkbookBuffer((wb, ws) => {
                ws.getCell("B2").value = { text: "x", hyperlink: "http://example.com/" };
            });
            const buffer = await tamper(base, {
                "xl/worksheets/sheet1.xml": (xml) => xml.replace('r="A1"', 'r="toString"'),
            });
            const reader = new ExcelJS.stream.xlsx.WorkbookReader(Readable.from([buffer]), { hyperlinks: "cache", sharedStrings: "cache" });
            const cells = [];
            for await (const ws of reader) {
                for await (const row of ws) {
                    // used to throw: the inherited toString was taken for the cell's hyperlink
                    row.eachCell((cell) => cells.push([cell.value, cell.hyperlink]));
                }
            }
            expect(cells).to.deep.equal([
                ["hello", undefined],
                ["x", undefined],
            ]);
        });
    });

    it("leaves the global prototypes untouched for every attribute of every part (document and streaming readers)", async () => {
        const cases = attributeCases(texts);
        expect(cases.length).to.be.above(200);
        const polluted = [];
        for (const [part, attribute] of cases) {
            for (const key of KEYS) {
                const buffer = await tamper(base, { [part]: replaceAttribute(attribute, key) });
                for (const run of [readAll, streamAll]) {
                    const step = await run(buffer);
                    if (step) {
                        polluted.push(`${part} ${attribute}="${key}" (${step})`);
                        restorePrototypes();
                    }
                }
            }
        }
        expect(polluted).to.deep.equal([]);
        expect(await writeUnrelatedWorkbook()).to.equal("hello");
    });

    describe("pivot table field names", () => {
        async function pivotCacheDefinition(valueHeader) {
            const wb = new ExcelJS.Workbook();
            const ws = wb.addWorksheet("src");
            ws.addRows([
                ["A", "B", valueHeader],
                ["a", "b", 5],
                ["a2", "b2", 7],
            ]);
            wb.addWorksheet("p").addPivotTable({ sourceSheet: ws, rows: ["A"], columns: ["B"], values: [valueHeader], metric: "sum" });
            return readPart(await wb.xlsx.writeBuffer(), "xl/pivotCache/pivotCacheDefinition1.xml");
        }

        it("writes a value field named like an Object.prototype member as a numeric field", async () => {
            const expected = (await pivotCacheDefinition("C")).replace('name="C"', 'name="constructor"');
            expect(await pivotCacheDefinition("constructor")).to.equal(expected);
        });

        it("rejects fields named like Object.prototype members that are not headers", () => {
            const wb = new ExcelJS.Workbook();
            const ws = wb.addWorksheet("src");
            ws.addRows([
                ["A", "B", "C"],
                ["a", "b", 5],
            ]);
            for (const name of KEYS) {
                expect(() =>
                    wb.addWorksheet(`p${name}`).addPivotTable({ sourceSheet: ws, rows: [name], columns: ["B"], values: ["C"] })
                ).to.throw(`The header name "${name}" was not found in src.`);
            }
        });
    });
});
