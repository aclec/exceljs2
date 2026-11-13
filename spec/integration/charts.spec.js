// Charts of the worksheets are carried opaque from a file read to the file written: the chart anchor of the drawing, the
// chart part and the parts it references (style, colors, embedded workbook...), with their relationships. The fixtures are
// built in code: a workbook written by ExcelJS, to which a hand-written bar chart is added (xlsx-tamper).
const fs = require("fs");
const path = require("path");
const { strFromU8, strToU8 } = require("fflate");
const { baseWorkbookBuffer, editParts, unzipParts } = require("../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");
const { loadChartParts, resolvePartName, relativeTarget } = verquire("xlsx/chart-parts");

const png = fs.readFileSync(path.join(__dirname, "data/image.png"));

const NS = {
    c: "http://schemas.openxmlformats.org/drawingml/2006/chart",
    r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    a: "http://schemas.openxmlformats.org/drawingml/2006/main",
    xdr: "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    pkg: "http://schemas.openxmlformats.org/package/2006/relationships",
};
const REL = {
    drawing: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
    chart: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
    chartStyle: "http://schemas.microsoft.com/office/2011/relationships/chartStyle",
    chartColorStyle: "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle",
    package: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
    worksheet: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
    oleObject: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
};
const CT = {
    chart: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
    chartStyle: "application/vnd.ms-office.chartstyle+xml",
    chartColorStyle: "application/vnd.ms-office.chartcolorstyle+xml",
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    drawing: "application/vnd.openxmlformats-officedocument.drawing+xml",
};

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

// a minimal bar chart of Sheet1!B1:B3, its data in an embedded workbook (rId3)
const CHART_XML = `${XML_DECL}<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}"><c:chart><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>Sheet1!$B$1:$B$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart><c:externalData r:id="rId3"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>`;
const STYLE_XML = `${XML_DECL}<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" id="201"/>`;
const COLORS_XML = `${XML_DECL}<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" meth="cycle" id="10"/>`;
// any bytes: an embedded package is never parsed
const EMBEDDED = new Uint8Array([0x50, 0x4b, 0x05, 0x06, 1, 2, 3, 0, 255]);

function rels(list) {
    const items = list
        .map(
            ([id, type, target, external]) =>
                `<Relationship Id="${id}" Type="${type}" Target="${target}"${external ? ' TargetMode="External"' : ""}/>`
        )
        .join("");
    return `${XML_DECL}<Relationships xmlns="${NS.pkg}">${items}</Relationships>`;
}

const CHART_RELS = rels([
    ["rId1", REL.chartStyle, "style1.xml"],
    ["rId2", REL.chartColorStyle, "colors1.xml"],
    ["rId3", REL.package, "../embeddings/Microsoft_Excel_Worksheet.xlsx"],
]);

// a chart anchor (D2:J15) referencing the drawing relationship `rId`, its namespaces declared on <c:chart> (as Excel does)
function chartAnchor(rId, { name = "Chart 1", shapeId = 7, chartNamespaces = true } = {}) {
    const ns = chartNamespaces ? ` xmlns:c="${NS.c}" xmlns:r="${NS.r}"` : "";
    return (
        '<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
        "<xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>14</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>" +
        `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${shapeId}" name="${name}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
        '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
        `<c:chart${ns} r:id="${rId}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`
    );
}

const DRAWING_OPEN = `${XML_DECL}<xdr:wsDr xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}">`;

/**
 * A workbook (Sheet1: B1:B3 = 1, 2, 3) with a bar chart in the drawing of Sheet1, and optionally an image in the same
 * drawing. `edit(parts)` may change the parts (strings, or bytes for the embedded workbook) before they are zipped.
 */
async function chartWorkbook({ image = false, anchors = [chartAnchor("rId9")], edit } = {}) {
    const base = await baseWorkbookBuffer((wb, ws) => {
        ws.getCell("B1").value = 1;
        ws.getCell("B2").value = 2;
        ws.getCell("B3").value = 3;
        if (image) {
            ws.addImage(wb.addImage({ buffer: png, extension: "png" }), "L2:M4");
        }
    });
    return editParts(base, (parts) => {
        const text = {};
        if (image) {
            const drawing = strFromU8(parts["xl/drawings/drawing1.xml"]);
            text["xl/drawings/drawing1.xml"] = drawing.replace("</xdr:wsDr>", `${anchors.join("")}</xdr:wsDr>`);
            const drawingRels = strFromU8(parts["xl/drawings/_rels/drawing1.xml.rels"]);
            text["xl/drawings/_rels/drawing1.xml.rels"] = drawingRels.replace(
                "</Relationships>",
                `<Relationship Id="rId9" Type="${REL.chart}" Target="../charts/chart1.xml"/></Relationships>`
            );
        } else {
            const sheet = strFromU8(parts["xl/worksheets/sheet1.xml"]);
            text["xl/worksheets/sheet1.xml"] = sheet.replace("</worksheet>", '<drawing r:id="rId1"/></worksheet>');
            text["xl/worksheets/_rels/sheet1.xml.rels"] = rels([["rId1", REL.drawing, "../drawings/drawing1.xml"]]);
            text["xl/drawings/drawing1.xml"] = `${DRAWING_OPEN}${anchors.join("")}</xdr:wsDr>`;
            text["xl/drawings/_rels/drawing1.xml.rels"] = rels([["rId9", REL.chart, "../charts/chart1.xml"]]);
        }
        text["xl/charts/chart1.xml"] = CHART_XML;
        text["xl/charts/_rels/chart1.xml.rels"] = CHART_RELS;
        text["xl/charts/style1.xml"] = STYLE_XML;
        text["xl/charts/colors1.xml"] = COLORS_XML;
        const contentTypes = strFromU8(parts["[Content_Types].xml"]);
        text["[Content_Types].xml"] = contentTypes.replace(
            "</Types>",
            `<Default Extension="xlsx" ContentType="${CT.xlsx}"/>` +
                (image ? "" : `<Override PartName="/xl/drawings/drawing1.xml" ContentType="${CT.drawing}"/>`) +
                `<Override PartName="/xl/charts/chart1.xml" ContentType="${CT.chart}"/>` +
                `<Override PartName="/xl/charts/style1.xml" ContentType="${CT.chartStyle}"/>` +
                `<Override PartName="/xl/charts/colors1.xml" ContentType="${CT.chartColorStyle}"/></Types>`
        );
        for (const [name, value] of Object.entries(text)) {
            parts[name] = strToU8(value);
        }
        parts["xl/embeddings/Microsoft_Excel_Worksheet.xlsx"] = EMBEDDED;
        if (edit) {
            edit(parts);
            for (const [name, value] of Object.entries(parts)) {
                if (typeof value === "string") {
                    parts[name] = strToU8(value);
                }
            }
        }
    });
}

async function roundTrip(buffer) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    return { wb, buffer: await wb.xlsx.writeBuffer() };
}

// the parts of a zip, as text
function textParts(buffer) {
    const parts = unzipParts(buffer);
    return { parts, text: (name) => (parts[name] ? strFromU8(parts[name]) : undefined) };
}

// the relationships of a .rels part: [{Id, Type, Target, TargetMode}]
function parseRels(xml) {
    return [...xml.matchAll(/<Relationship ([^>]*)\/>/g)].map(([, attributes]) =>
        Object.fromEntries([...attributes.matchAll(/(\w+)="([^"]*)"/g)].map(([, name, value]) => [name, value]))
    );
}

function overrides(contentTypes) {
    return Object.fromEntries(
        [...contentTypes.matchAll(/<Override PartName="([^"]+)" ContentType="([^"]+)"\/>/g)].map(([, n, t]) => [n, t])
    );
}

function chartPartNames(buffer) {
    return Object.keys(unzipParts(buffer)).filter((name) => name.startsWith("xl/charts/") || name.startsWith("xl/embeddings/"));
}

// checks the written chart of Sheet1: the drawing anchor, its relationship, the chart parts and their content types
function expectChart(buffer) {
    const { parts, text } = textParts(buffer);
    // the worksheet references its drawing
    const sheetRels = parseRels(text("xl/worksheets/_rels/sheet1.xml.rels"));
    const drawingRel = sheetRels.find((rel) => rel.Type === REL.drawing);
    expect(drawingRel.Target).to.equal("../drawings/drawing1.xml");
    expect(text("xl/worksheets/sheet1.xml")).to.include(`<drawing r:id="${drawingRel.Id}"/>`);

    // the anchor references the chart through a relationship of the drawing
    const drawing = text("xl/drawings/drawing1.xml");
    const [, chartId] = /<c:chart [^>]*r:id="([^"]+)"/.exec(drawing);
    const chartRel = parseRels(text("xl/drawings/_rels/drawing1.xml.rels")).find((rel) => rel.Id === chartId);
    expect(chartRel).to.deep.equal({ Id: chartId, Type: REL.chart, Target: "../charts/chart1.xml" });

    // the parts: the same bytes, the relationships of the chart retargeted to the names written
    expect(text("xl/charts/chart1.xml")).to.equal(CHART_XML);
    expect(text("xl/charts/style1.xml")).to.equal(STYLE_XML);
    expect(text("xl/charts/colors1.xml")).to.equal(COLORS_XML);
    expect(Array.from(parts["xl/embeddings/Microsoft_Excel_Worksheet1.xlsx"])).to.deep.equal(Array.from(EMBEDDED));
    expect(parseRels(text("xl/charts/_rels/chart1.xml.rels"))).to.deep.equal([
        { Id: "rId1", Type: REL.chartStyle, Target: "style1.xml" },
        { Id: "rId2", Type: REL.chartColorStyle, Target: "colors1.xml" },
        { Id: "rId3", Type: REL.package, Target: "../embeddings/Microsoft_Excel_Worksheet1.xlsx" },
    ]);
    const types = overrides(text("[Content_Types].xml"));
    expect(types["/xl/drawings/drawing1.xml"]).to.equal(CT.drawing);
    expect(types["/xl/charts/chart1.xml"]).to.equal(CT.chart);
    expect(types["/xl/charts/style1.xml"]).to.equal(CT.chartStyle);
    expect(types["/xl/charts/colors1.xml"]).to.equal(CT.chartColorStyle);
    expect(types["/xl/embeddings/Microsoft_Excel_Worksheet1.xlsx"]).to.equal(CT.xlsx);
    return { drawing, text };
}

describe("Charts carried from a file read to the file written", () => {
    it("keeps a bar chart: its anchor, chart, style, colors and embedded workbook, with consistent relationships", async () => {
        const first = await roundTrip(await chartWorkbook());
        const { drawing } = expectChart(first.buffer);
        // the anchor as read, its namespaces declared
        expect(drawing).to.include('<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>3</xdr:col>');
        expect(drawing).to.include(`xmlns:c="${NS.c}"`);
        expect(drawing).to.include('<xdr:cNvPr id="1" name="Chart 1"/>');
        expect(first.wb.getWorksheet("Sheet1").getImages()).to.have.length(0);

        // load → write again: the same package
        const second = await roundTrip(first.buffer);
        expectChart(second.buffer);
        const a = textParts(first.buffer);
        const b = textParts(second.buffer);
        for (const name of [
            "xl/drawings/drawing1.xml",
            "xl/drawings/_rels/drawing1.xml.rels",
            "xl/charts/_rels/chart1.xml.rels",
            "[Content_Types].xml",
        ]) {
            expect(b.text(name), name).to.equal(a.text(name));
        }
    });

    it("keeps the images of a drawing that also holds a chart", async () => {
        const { wb, buffer } = await roundTrip(await chartWorkbook({ image: true }));
        expect(wb.getWorksheet("Sheet1").getImages()).to.have.length(1);
        const { drawing, text } = expectChart(buffer);
        expect(drawing).to.include("<xdr:pic>");
        // unique relationship ids and shape ids in the drawing
        const drawingRels = parseRels(text("xl/drawings/_rels/drawing1.xml.rels"));
        expect(new Set(drawingRels.map((rel) => rel.Id)).size).to.equal(drawingRels.length);
        expect(drawingRels.map((rel) => rel.Type)).to.include("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
        const shapeIds = [...drawing.matchAll(/<xdr:cNvPr id="(\d+)"/g)].map(([, id]) => id);
        expect(shapeIds).to.deep.equal(["1", "2"]);

        // the images still load from the written file
        const again = new ExcelJS.Workbook();
        await again.xlsx.load(buffer);
        expect(again.getWorksheet("Sheet1").getImages()).to.have.length(1);
    });

    it("writes the namespaces of the drawing in scope on the anchor, and its text and attributes encoded", async () => {
        const anchor = chartAnchor("rId9", { name: "a &lt;b&gt; &amp; &quot;c&quot;", chartNamespaces: false });
        const buffer = await chartWorkbook({
            anchors: [anchor],
            edit(parts) {
                parts["xl/drawings/drawing1.xml"] =
                    `${XML_DECL}<xdr:wsDr xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}" xmlns:c="${NS.c}" xmlns:r="${NS.r}">${anchor}</xdr:wsDr>`;
            },
        });
        const first = await roundTrip(buffer);
        const { drawing } = expectChart(first.buffer);
        expect(drawing).to.include(`<xdr:twoCellAnchor editAs="oneCell" xmlns:c="${NS.c}" xmlns:r="${NS.r}">`);
        expect(drawing).to.include('name="a &lt;b&gt; &amp; &quot;c&quot;"');
        // and it reads again
        expectChart((await roundTrip(first.buffer)).buffer);
    });

    it("writes several charts, numbered in order, and a chart of a renamed worksheet as it was read", async () => {
        const buffer = await chartWorkbook({
            anchors: [chartAnchor("rId9"), chartAnchor("rId10", { name: "Chart 2" })],
            edit(parts) {
                parts["xl/drawings/_rels/drawing1.xml.rels"] = rels([
                    ["rId9", REL.chart, "../charts/chart1.xml"],
                    ["rId10", REL.chart, "../charts/chart7.xml"],
                ]);
                parts["xl/charts/chart7.xml"] = CHART_XML.replace('<c:barDir val="col"/>', '<c:barDir val="bar"/>');
            },
        });
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buffer);
        // the formulas of a chart are not rewritten: its series still reference Sheet1
        wb.getWorksheet("Sheet1").name = "Renamed";
        const { text } = textParts(await wb.xlsx.writeBuffer());
        expect(text("xl/charts/chart1.xml")).to.equal(CHART_XML);
        expect(text("xl/charts/chart2.xml")).to.include('<c:barDir val="bar"/>');
        expect(text("xl/charts/chart2.xml")).to.include("<c:f>Sheet1!$B$1:$B$3</c:f>");
        expect(text("xl/charts/_rels/chart2.xml.rels")).to.equal(undefined);
        const drawingRels = parseRels(text("xl/drawings/_rels/drawing1.xml.rels"));
        expect(drawingRels.map((rel) => rel.Target)).to.deep.equal(["../charts/chart1.xml", "../charts/chart2.xml"]);
    });

    it("writes a workbook without charts as before (no chart part, no content type)", async () => {
        const base = await baseWorkbookBuffer((wb, ws) => ws.addImage(wb.addImage({ buffer: png, extension: "png" }), "B2:C3"));
        const { buffer } = await roundTrip(base);
        expect(chartPartNames(buffer)).to.deep.equal([]);
        expect(textParts(buffer).text("[Content_Types].xml")).to.not.match(/chart|embeddings/);
        // parts written the same, read or not
        const a = textParts(base);
        const b = textParts(buffer);
        for (const name of ["xl/drawings/drawing1.xml", "xl/drawings/_rels/drawing1.xml.rels", "[Content_Types].xml"]) {
            expect(b.text(name), name).to.equal(a.text(name));
        }
    });

    it("is kept by a worksheet model copied to another worksheet", async () => {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(await chartWorkbook());
        const copy = wb.addWorksheet("Copy");
        copy.model = { ...wb.getWorksheet("Sheet1").model, name: "Copy" };
        wb.removeWorksheet(wb.getWorksheet("Sheet1").id);
        const { text } = textParts(await wb.xlsx.writeBuffer());
        expect(text("xl/charts/chart1.xml")).to.equal(CHART_XML);
    });

    it("is written as an independent copy for a second worksheet that shares it (a copied worksheet model)", async () => {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(await chartWorkbook());
        wb.addWorksheet("Copy").model = { ...wb.getWorksheet("Sheet1").model, name: "Copy" };
        const output = await wb.xlsx.writeBuffer();
        const { parts, text } = textParts(output);
        const target = (drawing) => parseRels(text(`xl/drawings/_rels/${drawing}.xml.rels`)).map((rel) => rel.Target);
        expect(target("drawing1")).to.deep.equal(["../charts/chart1.xml"]);
        expect(target("drawing2")).to.deep.equal(["../charts/chart2.xml"]);
        // the parts of the chart are copied too: no part is shared by the two charts
        expect(text("xl/charts/chart2.xml")).to.equal(CHART_XML);
        expect(parseRels(text("xl/charts/_rels/chart2.xml.rels"))).to.deep.equal([
            { Id: "rId1", Type: REL.chartStyle, Target: "style2.xml" },
            { Id: "rId2", Type: REL.chartColorStyle, Target: "colors2.xml" },
            { Id: "rId3", Type: REL.package, Target: "../embeddings/Microsoft_Excel_Worksheet2.xlsx" },
        ]);
        expect(Array.from(parts["xl/embeddings/Microsoft_Excel_Worksheet2.xlsx"])).to.deep.equal(Array.from(EMBEDDED));
        expect(overrides(text("[Content_Types].xml"))["/xl/charts/colors2.xml"]).to.equal(CT.chartColorStyle);
        // and both charts read again
        const again = await roundTrip(output);
        expect(chartPartNames(again.buffer).filter((name) => /chart\d[.]xml$/.test(name))).to.have.length(2);
    });

    describe("chartEx (as Excel writes it, in <mc:AlternateContent>)", () => {
        const CX = "http://schemas.microsoft.com/office/drawing/2014/chartex";
        const CX1 = "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex";
        const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        const CHART_EX_XML = `${XML_DECL}<cx:chartSpace xmlns:cx="${CX}"/>`;
        const corners =
            "<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>" +
            "<xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>";
        const frame =
            '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>' +
            `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${CX}">` +
            `<cx:chart xmlns:cx="${CX}" xmlns:r="${NS.r}" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame>`;
        const fallback =
            '<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="0" name=""/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/>' +
            "<xdr:txBody><a:bodyPr/><a:p><a:r><a:t>This chart isn't available in your version of Excel.</a:t></a:r></a:p></xdr:txBody></xdr:sp>";
        const anchor = (inner) => `<xdr:twoCellAnchor>${corners}${inner}<xdr:clientData/></xdr:twoCellAnchor>`;

        function chartExWorkbook(drawing) {
            return chartWorkbook({
                edit(parts) {
                    parts["xl/drawings/drawing1.xml"] = drawing;
                    parts["xl/drawings/_rels/drawing1.xml.rels"] = rels([
                        ["rId1", "http://schemas.microsoft.com/office/2014/relationships/chartEx", "../charts/chartEx1.xml"],
                    ]);
                    parts["xl/charts/chartEx1.xml"] = CHART_EX_XML;
                },
            });
        }

        it("keeps the chart, its <mc:Choice> and its <mc:Fallback>", async () => {
            const wrapper =
                `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:cx1="${CX1}" Requires="cx1">${anchor(frame)}</mc:Choice>` +
                `<mc:Fallback>${anchor(fallback)}</mc:Fallback></mc:AlternateContent>`;
            const first = await roundTrip(await chartExWorkbook(`${DRAWING_OPEN}${wrapper}</xdr:wsDr>`));
            const { text } = textParts(first.buffer);
            expect(text("xl/charts/chartEx1.xml")).to.equal(CHART_EX_XML);
            expect(overrides(text("[Content_Types].xml"))["/xl/charts/chartEx1.xml"]).to.equal("application/vnd.ms-office.chartex+xml");
            const drawing = text("xl/drawings/drawing1.xml");
            expect(drawing).to.include(
                `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:cx1="${CX1}" Requires="cx1"><xdr:twoCellAnchor>`
            );
            expect(drawing).to.include("<mc:Fallback><xdr:twoCellAnchor>");
            expect(drawing).to.include("This chart isn&apos;t available");
            const [, id] = /<cx:chart [^>]*r:id="([^"]+)"/.exec(drawing);
            expect(parseRels(text("xl/drawings/_rels/drawing1.xml.rels")).find((rel) => rel.Id === id).Target).to.equal(
                "../charts/chartEx1.xml"
            );
            // written again as read
            const second = textParts((await roundTrip(first.buffer)).buffer);
            expect(second.text("xl/drawings/drawing1.xml")).to.equal(drawing);
        });

        it("declares on the wrapper the namespaces of <xdr:wsDr> it uses, also those named by Requires", async () => {
            const wrapper =
                `<mc:AlternateContent><mc:Choice Requires="cx1">${anchor(frame)}</mc:Choice>` +
                `<mc:Fallback>${anchor(fallback)}</mc:Fallback></mc:AlternateContent>`;
            const root = `${XML_DECL}<xdr:wsDr xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}" xmlns:mc="${MC}" xmlns:cx1="${CX1}" xmlns:unused="u">`;
            const { text } = textParts((await roundTrip(await chartExWorkbook(`${root}${wrapper}</xdr:wsDr>`))).buffer);
            const drawing = text("xl/drawings/drawing1.xml");
            expect(drawing).to.match(/<mc:AlternateContent( xmlns:(mc|cx1)="[^"]+"){2}>/);
            expect(drawing).to.include(`xmlns:mc="${MC}"`);
            expect(drawing).to.include(`xmlns:cx1="${CX1}"`);
            expect(drawing).to.not.include("xmlns:unused");
        });

        it("reads a wrapper holding a picture as before (the image, no chart)", async () => {
            const buffer = await chartWorkbook({
                image: true,
                edit(parts) {
                    const drawing = strFromU8(parts["xl/drawings/drawing1.xml"]);
                    const picture = /<xdr:(two|one)CellAnchor[\s\S]*?<\/xdr:(two|one)CellAnchor>/.exec(drawing)[0];
                    const wrapper =
                        `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice Requires="a">${chartAnchor("rId9")}</mc:Choice>` +
                        `<mc:Fallback>${picture}</mc:Fallback></mc:AlternateContent>`;
                    parts["xl/drawings/drawing1.xml"] = drawing.replace(
                        /<xdr:(two|one)CellAnchor[\s\S]*<\/xdr:wsDr>/,
                        `${wrapper}</xdr:wsDr>`
                    );
                },
            });
            const { wb, buffer: output } = await roundTrip(buffer);
            expect(wb.getWorksheet("Sheet1").getImages()).to.have.length(1);
            expect(chartPartNames(output)).to.deep.equal([]);
        });
    });

    it("is read by the streaming reader as before (cells only)", async () => {
        const buffer = await chartWorkbook({ image: true });
        const reader = new ExcelJS.stream.xlsx.WorkbookReader(require("stream").Readable.from([buffer]), {});
        const values = [];
        for await (const worksheet of reader) {
            for await (const row of worksheet) {
                values.push(row.getCell(2).value);
            }
        }
        expect(values).to.deep.equal([1, 2, 3]);
    });

    describe("written by Excel (chart-sheet.xlsx)", () => {
        const input = fs.readFileSync(path.join(__dirname, "data/chart-sheet.xlsx"));

        it("carries the chart of a chartsheet with the chartsheet (see chartsheets.spec.js)", async () => {
            const { wb, buffer } = await roundTrip(input);
            expect(wb.worksheets).to.have.length(1);
            expect(wb.chartsheets).to.have.length(1);
            expect(chartPartNames(buffer)).to.deep.equal([
                "xl/charts/chart1.xml",
                "xl/charts/_rels/chart1.xml.rels",
                "xl/charts/colors1.xml",
                "xl/charts/style1.xml",
            ]);
        });

        it("keeps its chart (an absolute anchor) once anchored on the worksheet", async () => {
            const buffer = editParts(input, (parts) => {
                const sheet = strFromU8(parts["xl/worksheets/sheet1.xml"]);
                parts["xl/worksheets/sheet1.xml"] = strToU8(sheet.replace("</worksheet>", '<drawing r:id="rId1"/></worksheet>'));
                parts["xl/worksheets/_rels/sheet1.xml.rels"] = strToU8(rels([["rId1", REL.drawing, "../drawings/drawing1.xml"]]));
            });
            const original = textParts(input);
            const { text } = textParts((await roundTrip(buffer)).buffer);
            for (const name of ["xl/charts/chart1.xml", "xl/charts/style1.xml", "xl/charts/colors1.xml"]) {
                expect(text(name), name).to.equal(original.text(name));
            }
            expect(parseRels(text("xl/charts/_rels/chart1.xml.rels"))).to.deep.equal([
                { Id: "rId2", Type: REL.chartColorStyle, Target: "colors1.xml" },
                { Id: "rId1", Type: REL.chartStyle, Target: "style1.xml" },
            ]);
            const drawing = text("xl/drawings/drawing1.xml");
            expect(drawing).to.include('<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="9304587" cy="6077107"/>');
            expect(drawing).to.include('<xdr:cNvPr id="1" name="Chart 1">');
            expect(drawing).to.include('id="{C0CAAAD9-BEAB-4404-9D93-A04790160044}"');
            expect(drawing).to.include('<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"');
        });
    });
});

describe("Charts of a hostile file", () => {
    // loads, writes and loads again; returns the parts written
    async function written(buffer) {
        const { buffer: output } = await roundTrip(buffer);
        const again = new ExcelJS.Workbook();
        await again.xlsx.load(output);
        return output;
    }

    async function expectDropped(options) {
        const output = await written(await chartWorkbook(options));
        expect(chartPartNames(output)).to.deep.equal([]);
        expect(textParts(output).text("xl/drawings/drawing1.xml") || "").to.not.include("c:chart");
        return output;
    }

    it("drops a chart whose part is missing", async () => {
        await expectDropped({ edit: (parts) => delete parts["xl/charts/chart1.xml"] });
    });

    it("drops a chart when a part it references is missing (all or nothing)", async () => {
        await expectDropped({ edit: (parts) => delete parts["xl/charts/colors1.xml"] });
    });

    it("drops a chart whose anchor references a relationship missing from the drawing", async () => {
        await expectDropped({ anchors: [chartAnchor("rId404")] });
    });

    it("never resolves a target outside the zip, or to a relationships part or the content types", async () => {
        for (const target of [
            "../../../../etc/passwd",
            "../../../charts/chart1.xml",
            "C:\\charts\\chart1.xml",
            "file:///etc/passwd",
            "http://127.0.0.1/chart1.xml",
            "../charts/%63hart1.xml",
            "../charts//chart1.xml",
            "../charts/_rels/chart1.xml.rels",
            "../../[Content_Types].xml",
        ]) {
            await expectDropped({
                edit(parts) {
                    parts["xl/drawings/_rels/drawing1.xml.rels"] = rels([["rId9", REL.chart, target]]);
                },
            });
        }
        // a chart relationship escaping the package
        await expectDropped({
            edit(parts) {
                parts["xl/charts/_rels/chart1.xml.rels"] = CHART_RELS.replace("style1.xml", "../../../style1.xml");
            },
        });
    });

    it("accepts an absolute part name (from the package root)", async () => {
        const output = await written(
            await chartWorkbook({
                edit(parts) {
                    parts["xl/drawings/_rels/drawing1.xml.rels"] = rels([["rId9", REL.chart, "/xl/charts/chart1.xml"]]);
                },
            })
        );
        expectChart(output);
    });

    it("drops a chart with a relationship of a type that is not a chart part's", async () => {
        await expectDropped({
            edit(parts) {
                parts["xl/charts/_rels/chart1.xml.rels"] = rels([["rId3", REL.worksheet, "../worksheets/sheet1.xml"]]);
            },
        });
        // an OLE object or a package with an unexpected extension
        await expectDropped({
            edit(parts) {
                parts["xl/charts/_rels/chart1.xml.rels"] = rels([["rId3", REL.oleObject, "../worksheets/sheet1.xml"]]);
            },
        });
    });

    it("drops a chart whose relationships part is malformed or has duplicate ids, and still loads the file", async () => {
        await expectDropped({ edit: (parts) => (parts["xl/charts/_rels/chart1.xml.rels"] = "<Relationships><Oops/></Relationships>") });
        await expectDropped({
            edit(parts) {
                parts["xl/charts/_rels/chart1.xml.rels"] = rels([
                    ["rId1", REL.chartStyle, "style1.xml"],
                    ["rId1", REL.chartColorStyle, "colors1.xml"],
                ]);
            },
        });
    });

    it("ends on a cycle of relationships, and writes each part once", async () => {
        const output = await written(
            await chartWorkbook({
                edit(parts) {
                    // style → chart → style, and the chart to itself
                    parts["xl/charts/_rels/style1.xml.rels"] = rels([["rId1", REL.chart, "chart1.xml"]]);
                    parts["xl/charts/_rels/chart1.xml.rels"] = CHART_RELS.replace(
                        "</Relationships>",
                        `<Relationship Id="rId4" Type="${REL.chart}" Target="chart1.xml"/></Relationships>`
                    );
                },
            })
        );
        expect(chartPartNames(output).sort()).to.deep.equal(
            [
                "xl/charts/_rels/chart1.xml.rels",
                "xl/charts/_rels/style1.xml.rels",
                "xl/charts/chart1.xml",
                "xl/charts/colors1.xml",
                "xl/charts/style1.xml",
                "xl/embeddings/Microsoft_Excel_Worksheet1.xlsx",
            ].sort()
        );
        const { text } = textParts(output);
        expect(parseRels(text("xl/charts/_rels/style1.xml.rels"))).to.deep.equal([{ Id: "rId1", Type: REL.chart, Target: "chart1.xml" }]);
    });

    it("keeps an external relationship as it is, without following it", async () => {
        const output = await written(
            await chartWorkbook({
                edit(parts) {
                    parts["xl/charts/_rels/chart1.xml.rels"] = CHART_RELS.replace(
                        `Target="../embeddings/Microsoft_Excel_Worksheet.xlsx"/>`,
                        `Target="file:///C:/data/source&amp;.xlsx" TargetMode="External"/>`
                    ).replace(REL.package, REL.oleObject);
                },
            })
        );
        const chartRels = parseRels(textParts(output).text("xl/charts/_rels/chart1.xml.rels"));
        expect(chartRels[2]).to.deep.equal({
            Id: "rId3",
            Type: REL.oleObject,
            Target: "file:///C:/data/source&amp;.xlsx",
            TargetMode: "External",
        });
        expect(chartPartNames(output)).to.not.include("xl/embeddings/Microsoft_Excel_Worksheet1.xlsx");
    });

    it("gives a chart part to one anchor only", async () => {
        const anchors = Array.from({ length: 200 }, () => chartAnchor("rId9"));
        const output = await written(await chartWorkbook({ anchors }));
        const drawing = textParts(output).text("xl/drawings/drawing1.xml");
        expect(drawing.match(/<c:chart /g)).to.have.length(1);
    });

    it("drops the anchors of a drawing that no worksheet references", async () => {
        const output = await written(
            await chartWorkbook({
                edit(parts) {
                    parts["xl/drawings/drawing2.xml"] = parts["xl/drawings/drawing1.xml"];
                    parts["xl/drawings/_rels/drawing2.xml.rels"] = parts["xl/drawings/_rels/drawing1.xml.rels"];
                    delete parts["xl/drawings/drawing1.xml"];
                },
            })
        );
        expect(chartPartNames(output)).to.deep.equal([]);
    });

    // a drawing of many namespace declarations (on <xdr:wsDr>) and many elements: the read stays linear (it was proportional
    // to their product: seconds for these sizes, and quadrupling with each doubling)
    async function loadTime(drawingXml) {
        const buffer = await chartWorkbook({
            edit(parts) {
                parts["xl/drawings/drawing1.xml"] = drawingXml;
            },
        });
        const start = Date.now();
        await new ExcelJS.Workbook().xlsx.load(buffer);
        return Date.now() - start;
    }
    const manyDeclarations = Array.from({ length: 8000 }, (_, i) => ` xmlns:p${i}="u${i}"`).join("");

    it("reads many anchors under many namespace declarations in linear time", async () => {
        const drawing = `${XML_DECL}<xdr:wsDr xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}"${manyDeclarations}>${"<xdr:absoluteAnchor/>".repeat(20000)}</xdr:wsDr>`;
        expect(await loadTime(drawing)).to.be.below(1000);
    });

    it("reads many namespace declarations inside an anchor in linear time", async () => {
        const inner = '<a:x xmlns:q="v"/>'.repeat(20000);
        const drawing = `${XML_DECL}<xdr:wsDr xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}"${manyDeclarations}><xdr:absoluteAnchor>${inner}</xdr:absoluteAnchor></xdr:wsDr>`;
        expect(await loadTime(drawing)).to.be.below(1000);
    });

    it("bounds the number of chart parts read", async () => {
        const entries = new Map();
        const drawingRels = [];
        const chartFrames = [];
        for (let i = 1; i <= 50; i++) {
            entries.set(`xl/charts/chart${i}.xml`, { data: async () => strToU8(CHART_XML) });
            drawingRels.push({ Id: `rId${i}`, Type: REL.chart, Target: `../charts/chart${i}.xml` });
            chartFrames.push({ name: "c:chart", attributes: [{ name: "r:id", value: `rId${i}`, role: "rel" }], children: [] });
        }
        let inflated = 0;
        for (const entry of entries.values()) {
            const { data } = entry;
            entry.data = () => {
                inflated++;
                return data();
            };
        }
        const parts = await loadChartParts({ drawing1: { anchors: [], chartFrames } }, { drawing1: drawingRels }, entries, 10);
        expect(parts.size).to.equal(10);
        expect(inflated).to.equal(10);
    });

    it("resolves part names and relative targets", () => {
        expect(resolvePartName("xl/drawings/drawing1.xml", "../charts/chart1.xml")).to.equal("xl/charts/chart1.xml");
        expect(resolvePartName("xl/drawings/drawing1.xml", "./../charts/./chart1.xml")).to.equal("xl/charts/chart1.xml");
        expect(resolvePartName("xl/drawings/drawing1.xml", "/xl/charts/chart1.xml")).to.equal("xl/charts/chart1.xml");
        expect(resolvePartName("xl/drawings/drawing1.xml", "../../../chart1.xml")).to.equal(undefined);
        expect(resolvePartName("xl/drawings/drawing1.xml", "")).to.equal(undefined);
        expect(relativeTarget("xl/drawings/drawing1.xml", "xl/charts/chart1.xml")).to.equal("../charts/chart1.xml");
        expect(relativeTarget("xl/charts/chart1.xml", "xl/charts/style1.xml")).to.equal("style1.xml");
        expect(relativeTarget("xl/charts/chart1.xml", "xl/embeddings/a.xlsx")).to.equal("../embeddings/a.xlsx");
    });
});
