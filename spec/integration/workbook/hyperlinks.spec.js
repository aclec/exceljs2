// Cell hyperlinks: external links are relationships, internal links (a place in the workbook,
// written by the user as "#Sheet!A1" or "#DefinedName") are a location attribute only,
// as Excel writes them (ECMA-376 CT_Hyperlink).
const { PassThrough } = require("stream");
const JSZip = require("jszip");
const { baseWorkbookBuffer, tamper, readPart } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

const SHEET = "xl/worksheets/sheet1.xml";
const RELS = "xl/worksheets/_rels/sheet1.xml.rels";

async function load(buffer) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    return wb;
}

async function hyperlinkNodes(buffer) {
    return (await readPart(buffer, SHEET)).match(/<hyperlink [^>]*>/g);
}

async function hyperlinkRels(buffer) {
    const zip = await JSZip.loadAsync(buffer);
    const rels = zip.file(RELS);
    return rels ? (await rels.async("string")).match(/<Relationship [^>]*\/hyperlink"[^>]*>/g) || [] : [];
}

const LINKS = {
    A1: { text: "sheet ref", hyperlink: "#Data!A1" },
    A2: { text: "quoted range", hyperlink: "#'My Data'!B2:C3" },
    A3: { text: "defined name", hyperlink: "#MyName" },
    A4: { text: "external", hyperlink: "https://example.com/page" },
    A5: { text: "external with anchor", hyperlink: "https://example.com/doc#part2" },
};

function addLinks(wb, ws) {
    wb.addWorksheet("Data");
    wb.addWorksheet("My Data");
    Object.entries(LINKS).forEach(([address, value]) => {
        ws.getCell(address).value = value;
    });
}

describe("Hyperlinks", () => {
    it("writes internal links as a location without '#' and without a relationship", async () => {
        const buffer = await baseWorkbookBuffer(addLinks);
        expect(await hyperlinkNodes(buffer)).to.deep.equal([
            '<hyperlink ref="A1" location="Data!A1"/>',
            '<hyperlink ref="A2" location="&apos;My Data&apos;!B2:C3"/>',
            '<hyperlink ref="A3" location="MyName"/>',
            '<hyperlink ref="A4" r:id="rId1"/>',
            '<hyperlink ref="A5" r:id="rId2"/>',
        ]);
        const rels = await hyperlinkRels(buffer);
        expect(rels).to.have.length(2);
        expect(rels[0]).to.include('Target="https://example.com/page"').and.include('TargetMode="External"');
        expect(rels[1]).to.include('Target="https://example.com/doc#part2"');
    });

    it("round-trips internal and external links without loss", async () => {
        let buffer = await baseWorkbookBuffer(addLinks);
        for (let i = 0; i < 2; i++) {
            const wb = await load(buffer);
            const ws = wb.getWorksheet("Sheet1");
            Object.entries(LINKS).forEach(([address, value]) => {
                expect(ws.getCell(address).value, address).to.deep.equal(value);
                expect(ws.getCell(address).type, address).to.equal(ExcelJS.ValueType.Hyperlink);
            });
            buffer = await wb.xlsx.writeBuffer();
        }
        expect(await hyperlinkRels(buffer)).to.have.length(2);
    });

    it("round-trips Excel-style links (location only, and an external target with a location)", async () => {
        const buffer = await tamper(
            await baseWorkbookBuffer((wb, ws) => {
                wb.addWorksheet("Data");
                ws.getCell("B1").value = 42;
                ws.getCell("B2").value = { text: "doc", hyperlink: "https://example.com/doc" };
            }),
            {
                [SHEET]: (xml) =>
                    xml
                        .replace('<hyperlink ref="B2" r:id="rId1"/>', '<hyperlink ref="B2" r:id="rId1" location="part2"/>')
                        .replace("</hyperlinks>", '<hyperlink ref="B1" location="Data!A1" display="Data"/></hyperlinks>'),
            }
        );
        const expectCells = (ws) => {
            expect(ws.getCell("B1").value).to.deep.equal({ text: 42, hyperlink: "#Data!A1" });
            expect(ws.getCell("B2").value).to.deep.equal({ text: "doc", hyperlink: "https://example.com/doc#part2" });
        };
        const wb = await load(buffer);
        expectCells(wb.getWorksheet("Sheet1"));

        const written = await wb.xlsx.writeBuffer();
        expect(await hyperlinkNodes(written)).to.include('<hyperlink ref="B1" location="Data!A1"/>');
        expect(await hyperlinkRels(written)).to.have.length(1);
        expectCells((await load(written)).getWorksheet("Sheet1"));
    });

    it("keeps reading the internal links written by exceljs 4", async () => {
        // 4.x wrote an External relationship to "#..." and, for Sheet!A1 targets, a location with '#'
        const buffer = await tamper(await baseWorkbookBuffer(addLinks), {
            [SHEET]: (xml) =>
                xml
                    .replace('<hyperlink ref="A1" location="Data!A1"/>', '<hyperlink ref="A1" r:id="rId8" location="#Data!A1"/>')
                    .replace('<hyperlink ref="A3" location="MyName"/>', '<hyperlink ref="A3" r:id="rId9"/>'),
            [RELS]: (xml) =>
                xml.replace(
                    "</Relationships>",
                    '<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="#Data!A1" TargetMode="External"/>' +
                        '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="#MyName" TargetMode="External"/>' +
                        "</Relationships>"
                ),
        });
        const wb = await load(buffer);
        const ws = wb.getWorksheet("Sheet1");
        expect(ws.getCell("A1").hyperlink).to.equal("#Data!A1");
        expect(ws.getCell("A3").hyperlink).to.equal("#MyName");
        const nodes = await hyperlinkNodes(await wb.xlsx.writeBuffer());
        expect(nodes).to.include('<hyperlink ref="A1" location="Data!A1"/>');
        expect(nodes).to.include('<hyperlink ref="A3" location="MyName"/>');
    });

    it("writes a legacy unprefixed Sheet!A1 target as an internal link", async () => {
        const buffer = await baseWorkbookBuffer((wb, ws) => {
            wb.addWorksheet("Data");
            ws.getCell("A1").value = { text: "legacy", hyperlink: "Data!B3" };
        });
        expect(await hyperlinkNodes(buffer)).to.deep.equal(['<hyperlink ref="A1" location="Data!B3"/>']);
        expect(await hyperlinkRels(buffer)).to.deep.equal([]);
        expect((await load(buffer)).getWorksheet("Sheet1").getCell("A1").hyperlink).to.equal("#Data!B3");
    });

    it("streaming writer: internal links are a location, external links a relationship", async () => {
        const stream = new PassThrough();
        const chunks = [];
        stream.on("data", (chunk) => chunks.push(chunk));
        const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ stream });
        const ws = wb.addWorksheet("Sheet1");
        wb.addWorksheet("Data").commit();
        ws.getCell("A1").value = { text: "internal", hyperlink: "#Data!A1", tooltip: "go" };
        ws.getCell("A2").value = { text: "name", hyperlink: "#MyName" };
        ws.getCell("A3").value = { text: "external", hyperlink: "https://example.com/" };
        ws.commit();
        await wb.commit();
        const buffer = Buffer.concat(chunks);

        expect(await hyperlinkNodes(buffer)).to.deep.equal([
            '<hyperlink ref="A1" tooltip="go" location="Data!A1"/>',
            '<hyperlink ref="A2" location="MyName"/>',
            '<hyperlink ref="A3" r:id="rId1"/>',
        ]);
        const rels = await hyperlinkRels(buffer);
        expect(rels).to.have.length(1);
        expect(rels[0]).to.include('Target="https://example.com/"');

        const loaded = (await load(buffer)).getWorksheet("Sheet1");
        expect(loaded.getCell("A1").hyperlink).to.equal("#Data!A1");
        expect(loaded.getCell("A2").hyperlink).to.equal("#MyName");
        expect(loaded.getCell("A3").hyperlink).to.equal("https://example.com/");
    });
});
