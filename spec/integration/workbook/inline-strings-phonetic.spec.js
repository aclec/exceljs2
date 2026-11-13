// Phonetic runs (<rPh>, furigana) and properties (<phoneticPr>) of inline strings (<is>): read as for shared strings
// (the reading is dropped, the base text kept), by the document and the streaming readers, without losing the
// following cells.
const { Readable } = require("stream");
const { baseWorkbookBuffer, tamper } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

// B1 holds the inline string `is`; C1 and A2 follow it
async function inlineString(is) {
    const buffer = await baseWorkbookBuffer((wb, ws) => {
        ws.getCell("B1").value = 1;
        ws.getCell("C1").value = "after";
        ws.getCell("A2").value = 2;
    });
    return tamper(buffer, {
        "xl/worksheets/sheet1.xml": (xml) =>
            xml
                .replace(/<c r="B1"[^>]*>.*?<\/c>/, `<c r="B1" t="inlineStr">${is}</c>`)
                // the sheet-level phonetic properties that Excel writes in a Japanese workbook
                .replace("<pageMargins", '<phoneticPr fontId="1" type="noConversion"/><pageMargins'),
    });
}

async function documentValues(buffer) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    const ws = wb.getWorksheet("Sheet1");
    return ["B1", "C1", "A2"].map((address) => ws.getCell(address).value);
}

async function streamValues(buffer, options = {}) {
    const values = [];
    const reader = new ExcelJS.stream.xlsx.WorkbookReader(Readable.from([buffer]), options);
    for await (const worksheet of reader) {
        for await (const row of worksheet) {
            values.push(...row.values.slice(row.number === 1 ? 2 : 1));
        }
    }
    return values;
}

const CASES = [
    {
        name: "a text with its reading and phonetic properties",
        is: '<is><t>東京</t><rPh sb="0" eb="2"><t>トウキョウ</t></rPh><phoneticPr fontId="1" type="noConversion"/></is>',
        document: "東京",
        stream: "東京",
    },
    {
        name: "a text with readings only",
        is: '<is><t>東京都</t><rPh sb="0" eb="2"><t>トウキョウ</t></rPh><rPh sb="2" eb="3"><t>ト</t></rPh></is>',
        document: "東京都",
        stream: "東京都",
    },
    {
        name: "a text with phonetic properties only",
        is: '<is><t>東京</t><phoneticPr fontId="1"/></is>',
        document: "東京",
        stream: "東京",
    },
    {
        name: "rich text runs with a reading",
        is: '<is><r><t>東</t></r><r><rPr><b/></rPr><t>京</t></r><rPh sb="0" eb="2"><t>トウキョウ</t></rPh><phoneticPr fontId="1"/></is>',
        document: { richText: [{ text: "東" }, { font: { bold: true }, text: "京" }] },
        // the streaming reader reads an inline string as its text
        stream: "東京",
    },
    {
        name: "rich text runs",
        is: "<is><r><t>a</t></r><r><t>b</t></r></is>",
        document: { richText: [{ text: "a" }, { text: "b" }] },
        stream: "ab",
    },
];

describe("Inline strings with phonetic runs", () => {
    for (const { name, is, document, stream } of CASES) {
        it(`document reader: ${name}`, async () => {
            expect(await documentValues(await inlineString(is))).to.deep.equal([document, "after", 2]);
        });

        it(`streaming reader: ${name}`, async () => {
            expect(await streamValues(await inlineString(is))).to.deep.equal([stream, "after", 2]);
        });
    }

    it("reads the same text as a shared string with the same phonetic runs", async () => {
        const si = '<t>東京</t><rPh sb="0" eb="2"><t>トウキョウ</t></rPh><phoneticPr fontId="1" type="noConversion"/>';
        const buffer = await tamper(await inlineString(`<is>${si}</is>`), {
            "xl/sharedStrings.xml": (xml) => xml.replace("<si><t>after</t></si>", `<si>${si}</si>`),
        });
        expect(await documentValues(buffer)).to.deep.equal(["東京", "東京", 2]);
        expect(await streamValues(buffer, { sharedStrings: "cache" })).to.deep.equal(["東京", "東京", 2]);
    });

    it("ignores the other elements of a cell (extLst)", async () => {
        const buffer = await inlineString('<is><t>x</t></is><extLst><ext uri="u"><y a="1"/></ext></extLst>');
        expect(await documentValues(buffer)).to.deep.equal(["x", "after", 2]);
        expect(await streamValues(buffer)).to.deep.equal(["x", "after", 2]);
    });

    it("streaming reader: the white space after a value is not part of it", async () => {
        const buffer = await tamper(await inlineString("<is><t>x</t></is>"), {
            "xl/worksheets/sheet1.xml": (xml) =>
                xml
                    .replace(/<c r="C1"[^>]*>.*?<\/c>/, '<c r="C1" t="str">\n  <v>after</v>\n</c>')
                    .replace("<is><t>x</t></is>", "\n<is>\n<t>x</t>\n</is>\n"),
        });
        expect(await streamValues(buffer)).to.deep.equal(["x", "after", 2]);
        expect(await documentValues(buffer)).to.deep.equal(["x", "after", 2]);
    });
});
