// Streaming writer (ExcelJS.stream.xlsx.WorkbookWriter): sheet XML structure and shared strings.
const { PassThrough, Readable } = require("stream");
const { readPart } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

const IMAGE_FILENAME = `${__dirname}/../data/image.png`;

// commit the workbook into memory and return the xlsx buffer
async function writeBuffer(build, options = {}) {
    const stream = new PassThrough();
    const chunks = [];
    stream.on("data", (chunk) => chunks.push(chunk));
    const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ stream, ...options });
    await build(wb);
    await wb.commit();
    return Buffer.concat(chunks);
}

describe("WorkbookWriter (stream)", () => {
    it("writes worksheet elements in CT_Worksheet order", async () => {
        const buffer = await writeBuffer((wb) => {
            const ws = wb.addWorksheet("Sheet", {
                properties: { tabColor: { argb: "FFFF0000" } },
                views: [{ state: "frozen", ySplit: 1 }],
                pageSetup: { showGridLines: true, orientation: "landscape" },
                headerFooter: { oddHeader: "header" },
            });
            ws.autoFilter = "A1:B1";
            ws.mergeCells("C1:D1");
            ws.getCell("A1").value = "text";
            ws.getCell("A1").note = "a note";
            ws.getCell("B1").value = { text: "link", hyperlink: "https://example.com" };
            ws.getRow(1).addPageBreak();
            ws.addConditionalFormatting({
                ref: "A2:A3",
                rules: [{ type: "dataBar", cfvo: [{ type: "min" }, { type: "max" }], color: { argb: "FF638EC6" } }],
            });
            ws.dataValidations.add("A2", { type: "whole", operator: "between", formulae: [1, 9] });
            ws.addBackgroundImage(wb.addImage({ filename: IMAGE_FILENAME, extension: "png" }));
            ws.getCell("A2").value = 1;
            ws.getCell("A3").value = 2;
            return ws.protect("", { selectLockedCells: true });
        });
        const xml = await readPart(buffer, "xl/worksheets/sheet1.xml");

        // ECMA-376 CT_Worksheet sequence (elements the stream writer can produce)
        const order = [
            "sheetPr",
            "sheetViews",
            "sheetFormatPr",
            "sheetData",
            "sheetProtection",
            "autoFilter",
            "mergeCells",
            "conditionalFormatting",
            "dataValidations",
            "hyperlinks",
            "printOptions",
            "pageMargins",
            "pageSetup",
            "headerFooter",
            "rowBreaks",
            "legacyDrawing",
            "picture",
            "extLst",
        ];
        // last occurrence: a dataBar cfRule has its own nested extLst
        const positions = order.map((tag) => ({
            tag,
            index: Math.max(xml.lastIndexOf(`<${tag} `), xml.lastIndexOf(`<${tag}>`), xml.lastIndexOf(`<${tag}/>`)),
        }));
        expect(positions.filter(({ index }) => index < 0).map(({ tag }) => tag)).to.deep.equal([]);
        const sorted = [...positions].sort((a, b) => a.index - b.index).map(({ tag }) => tag);
        expect(sorted).to.deep.equal(order);

        // the dataBar cfRule references its x14 extension by id
        const ids = [...xml.matchAll(/<x14:id>([^<]+)<\/x14:id>|<x14:cfRule [^>]*id="([^"]+)"/g)];
        expect(ids).to.have.length(2);
        expect(ids[0][1]).to.be.a("string");
        expect(ids[0][1]).to.equal(ids[1][2]);

        // and the document reader still understands it
        const wb2 = new ExcelJS.Workbook();
        await wb2.xlsx.load(buffer);
        const ws2 = wb2.getWorksheet("Sheet");
        expect(ws2.pageSetup.showGridLines).to.equal(true);
        expect(ws2.conditionalFormattings[0].rules[0].type).to.equal("dataBar");
        expect(ws2.getCell("B1").hyperlink).to.equal("https://example.com");
    });

    it("writes the differential formats of conditional formatting rules without useStyles", async () => {
        const buffer = await writeBuffer((wb) => {
            const ws = wb.addWorksheet("Sheet");
            ws.getCell("A1").value = 1;
            ws.addConditionalFormatting({
                ref: "A1:A5",
                rules: [
                    { type: "expression", formulae: ["A1>1"], style: { font: { bold: true }, numFmt: '0.0 "kg"' } },
                    { type: "cellIs", operator: "equal", formulae: ["Other!$A$1"], style: { numFmt: "0.00%" } },
                    {
                        type: "cellIs",
                        operator: "lessThan",
                        formulae: ["0"],
                        style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFF0000" } } },
                    },
                ],
            });
            ws.commit();
        });

        const styles = await readPart(buffer, "xl/styles.xml");
        const [, numFmtId] = styles.match(/<numFmt numFmtId="(\d+)" formatCode="0.0 &quot;kg&quot;"\/>/);
        expect(styles).to.include(
            `<dxfs count="3"><dxf><font><b/></font><numFmt numFmtId="${numFmtId}" formatCode="0.0 &quot;kg&quot;"/></dxf>`
        );
        const sheet = await readPart(buffer, "xl/worksheets/sheet1.xml");
        expect(sheet).to.match(/<cfRule type="expression" dxfId="0" priority="1">/);
        expect(sheet).to.match(/<cfRule type="cellIs" dxfId="2" priority="3" operator="lessThan">/);
        // the x14 rule carries its differential format inline, with a built-in number format
        expect(sheet).to.include('<x14:dxf><numFmt numFmtId="10" formatCode="0.00%"/></x14:dxf>');

        const wb2 = new ExcelJS.Workbook();
        await wb2.xlsx.load(buffer);
        const [cf] = wb2.getWorksheet("Sheet").conditionalFormattings;
        // the x14 rule is read after the main ones
        const [expression, lessThan, crossSheet] = cf.rules;
        expect(expression.style.numFmt).to.equal('0.0 "kg"');
        expect(expression.style.font.bold).to.equal(true);
        expect(lessThan.style.fill.bgColor).to.deep.equal({ argb: "FFFF0000" });
        expect(crossSheet.style.numFmt).to.equal("0.00%");
    });

    it("keeps distinct rich text values apart in shared strings", async () => {
        const bold = { richText: [{ font: { bold: true }, text: "first" }] };
        const mixed = { richText: [{ text: "second " }, { font: { italic: true }, text: "part" }] };
        const json = JSON.stringify(bold);
        const values = [bold, mixed, json, "[object Object]", bold, "plain"];
        const buffer = await writeBuffer(
            (wb) => {
                const ws = wb.addWorksheet("Sheet");
                values.forEach((value, index) => {
                    ws.getCell(index + 1, 1).value = value;
                });
            },
            { useSharedStrings: true }
        );

        const wb2 = new ExcelJS.Workbook();
        await wb2.xlsx.load(buffer);
        const read = values.map((value, index) => wb2.getWorksheet("Sheet").getCell(index + 1, 1).value);
        expect(read[0].richText.map(({ text }) => text)).to.deep.equal(["first"]);
        expect(read[0].richText[0].font.bold).to.equal(true);
        expect(read[1].richText.map(({ text }) => text)).to.deep.equal(["second ", "part"]);
        expect(read.slice(2, 4)).to.deep.equal([json, "[object Object]"]);
        expect(read[4]).to.deep.equal(read[0]);
        expect(read[5]).to.equal("plain");

        // the streaming reader sees the same values
        const streamed = [];
        for await (const worksheet of new ExcelJS.stream.xlsx.WorkbookReader(Readable.from([buffer]))) {
            for await (const row of worksheet) {
                streamed.push(row.values[1]);
            }
        }
        const text = (value) => (value.richText ? value.richText.map((run) => run.text).join("|") : value);
        expect(streamed.map(text)).to.deep.equal(["first", "second |part", json, "[object Object]", "first", "plain"]);

        // count: the references to the shared strings, uniqueCount: the shared strings
        const xml = await readPart(buffer, "xl/sharedStrings.xml");
        expect(xml).to.match(/<sst [^>]* count="6" uniqueCount="5">/);
    });

    it("treats column keys named like Object.prototype members as plain keys", async () => {
        let worksheet;
        const buffer = await writeBuffer((wb) => {
            worksheet = wb.addWorksheet("Sheet");
            expect(worksheet.getColumnKey("constructor")).to.equal(undefined);
            worksheet.columns = [
                { header: "proto", key: "__proto__" },
                { header: "ctor", key: "constructor" },
                { header: "name", key: "name" },
            ];
            expect(worksheet.getColumn("constructor").number).to.equal(2);
            expect(worksheet.getColumn("__proto__").number).to.equal(1);
            // JSON.parse creates an own "__proto__" property (an object literal would not)
            worksheet.addRow(JSON.parse('{"__proto__": "p", "constructor": "c", "name": "n"}')).commit();
        });
        const keys = [];
        worksheet.eachColumnKey((column, key) => keys.push(`${key}:${column.number}`));
        expect(keys).to.deep.equal(["__proto__:1", "constructor:2", "name:3"]);

        const wb2 = new ExcelJS.Workbook();
        await wb2.xlsx.load(buffer);
        expect(wb2.getWorksheet("Sheet").getRow(2).values.slice(1)).to.deep.equal(["p", "c", "n"]);
    });

    it("gives comment VML shapes workbook-unique ids on every worksheet", async () => {
        const buffer = await writeBuffer((wb) => {
            [1, 2, 3].forEach((n) => {
                const ws = wb.addWorksheet(`Sheet${n}`);
                ws.getCell("A1").value = n;
                ws.getCell("A1").note = `note ${n}a`;
                ws.getCell("B2").value = n;
                ws.getCell("B2").note = `note ${n}b`;
                ws.commit();
            });
        });
        const blocks = [];
        const shapeIds = [];
        for (const n of [1, 2, 3]) {
            const vml = await readPart(buffer, `xl/drawings/vmlDrawing${n}.vml`);
            blocks.push(/<o:idmap [^>]*data="(\d+)"/.exec(vml)[1]);
            shapeIds.push(...[...vml.matchAll(/<v:shape [^>]*id="([^"]+)"/g)].map((match) => match[1]));
        }
        // as the document writer (and Excel): block n = sheet id, ids n*1024+1 ..
        expect(blocks).to.deep.equal(["1", "2", "3"]);
        expect(shapeIds).to.deep.equal(["_x0000_s1025", "_x0000_s1026", "_x0000_s2049", "_x0000_s2050", "_x0000_s3073", "_x0000_s3074"]);

        const wb2 = new ExcelJS.Workbook();
        await wb2.xlsx.load(buffer);
        expect(wb2.getWorksheet("Sheet3").getCell("B2").note).to.equal("note 3b");
    });
});
