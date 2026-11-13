// Charts created by worksheet.addChart: validated when added, written with their cached values read from the cells
// (src/doc/chart.ts, src/xlsx/chart-space.ts), through the code path of the charts read from a file (chart-parts.ts),
// then carried as read. Sample files for Excel are written to spec/out/charts-samples/; with OOXML_SCHEMAS set
// (spec/utils/xmllint.js), every chart and drawing part written is validated against the schemas.
const fs = require("fs");
const path = require("path");
const { strFromU8 } = require("fflate");
const { unzipParts } = require("../utils/xlsx-tamper");
const { schemaValidator } = require("../utils/xmllint");

const ExcelJS = verquire("exceljs");

const png = fs.readFileSync(path.join(__dirname, "data/image.png"));
const SAMPLES = path.join(__dirname, "../out/charts-samples");

const REL_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const CT_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

function text(buffer, name) {
    const part = unzipParts(buffer)[name];
    return part ? strFromU8(part) : undefined;
}

function partNames(buffer) {
    return Object.keys(unzipParts(buffer));
}

async function reload(buffer) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    return wb;
}

// a workbook with the data of the samples: Data!A1:D5 (months, two numeric series, a formula series)
function dataWorkbook(sheetName = "Data") {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet(sheetName);
    ws.addRow(["Month", "North", "South", "Total"]);
    [
        ["Jan", 10, 7],
        ["Feb", 12, 9],
        ["Mar", 9, 11],
        ["Apr", 14, 10],
    ].forEach(([month, north, south], index) => {
        const row = index + 2;
        ws.addRow([month, north, south, { formula: `B${row}+C${row}`, result: north + south }]);
    });
    return { wb, ws };
}

function columnChart(overrides = {}) {
    return {
        type: "column",
        title: "Sales",
        anchor: "F2:M18",
        series: [
            { name: { ref: "B1" }, categories: "A2:A5", values: "B2:B5" },
            { name: { ref: "C1" }, categories: "A2:A5", values: "C2:C5" },
        ],
        ...overrides,
    };
}

// the chart parts and the drawings of a file, validated against the schemas when they are available
const validate = schemaValidator();
function validateParts(buffer, label) {
    if (!validate) {
        return;
    }
    const dir = path.join(SAMPLES, "xml", label);
    fs.rmSync(dir, { recursive: true, force: true });
    fs.mkdirSync(dir, { recursive: true });
    const parts = unzipParts(buffer);
    const files = { "dml-chart.xsd": [], "dml-spreadsheetDrawing.xsd": [] };
    for (const name of Object.keys(parts)) {
        const match = /^xl\/(charts|drawings)\/([a-z]+\d+\.xml)$/.exec(name);
        if (match) {
            const file = path.join(dir, `${match[1]}-${match[2]}`);
            fs.writeFileSync(file, parts[name]);
            files[match[1] === "charts" ? "dml-chart.xsd" : "dml-spreadsheetDrawing.xsd"].push(file);
        }
    }
    for (const [xsd, list] of Object.entries(files)) {
        if (list.length) {
            const result = validate(xsd, list);
            expect(result.ok, `${label}: ${result.output}`).to.be.true();
        }
    }
}

describe("Charts created by worksheet.addChart", () => {
    it("writes the chart part, its relationship, its content type and its anchor", async () => {
        const { wb, ws } = dataWorkbook();
        const chart = ws.addChart(columnChart());
        expect(chart.type).to.equal("column");
        expect(chart.title).to.equal("Sales");
        const buffer = await wb.xlsx.writeBuffer();
        validateParts(buffer, "basic");

        expect(text(buffer, "xl/worksheets/sheet1.xml")).to.match(/<drawing r:id="rId\d+"\/>/);
        const drawing = text(buffer, "xl/drawings/drawing1.xml");
        expect(drawing).to.include(
            '<xdr:twoCellAnchor><xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>13</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>18</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="1" name="Chart 1"/>'
        );
        expect(drawing).to.include(
            '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>'
        );
        expect(text(buffer, "xl/drawings/_rels/drawing1.xml.rels")).to.include(
            `<Relationship Id="rId1" Type="${REL_CHART}" Target="../charts/chart1.xml"/>`
        );
        expect(text(buffer, "[Content_Types].xml")).to.include(`<Override PartName="/xl/charts/chart1.xml" ContentType="${CT_CHART}"/>`);
        // a generated chart has no other part
        expect(partNames(buffer).filter((name) => name.startsWith("xl/charts/"))).to.deep.equal(["xl/charts/chart1.xml"]);

        const xml = text(buffer, "xl/charts/chart1.xml");
        expect(xml).to.include('<c:f>Data!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>North</c:v></c:pt></c:strCache>');
        expect(xml).to.include(
            '<c:cat><c:strRef><c:f>Data!$A$2:$A$5</c:f><c:strCache><c:ptCount val="4"/><c:pt idx="0"><c:v>Jan</c:v></c:pt><c:pt idx="1"><c:v>Feb</c:v></c:pt><c:pt idx="2"><c:v>Mar</c:v></c:pt><c:pt idx="3"><c:v>Apr</c:v></c:pt></c:strCache></c:strRef></c:cat>'
        );
        expect(xml).to.include(
            '<c:val><c:numRef><c:f>Data!$C$2:$C$5</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="4"/><c:pt idx="0"><c:v>7</c:v></c:pt><c:pt idx="1"><c:v>9</c:v></c:pt><c:pt idx="2"><c:v>11</c:v></c:pt><c:pt idx="3"><c:v>10</c:v></c:pt></c:numCache></c:numRef></c:val>'
        );
    });

    it("caches the values of the cells when written: numbers, formula results, dates, texts, empty cells", async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet("S");
        ws.getCell("A1").value = "label";
        ws.getCell("A2").value = { richText: [{ text: "rich " }, { text: "text" }] };
        // A3 empty
        ws.getCell("A4").value = 7;
        ws.getCell("B1").value = 1.5;
        ws.getCell("B1").numFmt = "0.00";
        ws.getCell("B2").value = { formula: "B1*2", result: 3 };
        ws.getCell("B3").value = { formula: "NA()", result: { error: "#N/A" } };
        ws.getCell("B4").value = "text";
        ws.getCell("B5").value = { formula: "B1+1" }; // no result: no point
        ws.getCell("B6").value = 4;
        ws.getCell("C1").value = new Date(Date.UTC(2024, 0, 1));
        ws.getCell("C2").value = new Date(Date.UTC(2024, 0, 2));
        ws.getCell("D1").value = 10;
        ws.getCell("D2").value = 20;
        ws.mergeCells("D2:D3");
        ws.getCell("E1").value = "Total";
        ws.getCell("E2").value = new Date(Date.UTC(2024, 0, 2, 3, 4, 5));
        ws.getCell("E3").value = new Date(Date.UTC(2024, 0, 3));
        ws.addChart({
            type: "line",
            anchor: "F2:M18",
            series: [
                { categories: "A1:A4", values: "B1:B6" },
                { categories: "C1:C2", values: "D1:D3" },
                { categories: "E1:E3", values: "D1:D3" },
            ],
        });
        const xml = text(await wb.xlsx.writeBuffer(), "xl/charts/chart1.xml");
        // texts: rich text as its text, a number as its text, the empty cell without point
        expect(xml).to.include(
            '<c:strCache><c:ptCount val="4"/><c:pt idx="0"><c:v>label</c:v></c:pt><c:pt idx="1"><c:v>rich text</c:v></c:pt><c:pt idx="3"><c:v>7</c:v></c:pt></c:strCache>'
        );
        // numbers: the format of the first number, the formula results, no point for an error, a text, no result
        expect(xml).to.include(
            '<c:numCache><c:formatCode>0.00</c:formatCode><c:ptCount val="6"/><c:pt idx="0"><c:v>1.5</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt><c:pt idx="5"><c:v>4</c:v></c:pt></c:numCache>'
        );
        // dates as categories: serial numbers with a date format
        expect(xml).to.include(
            '<c:cat><c:numRef><c:f>S!$C$1:$C$2</c:f><c:numCache><c:formatCode>mm-dd-yy</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>45292</c:v></c:pt><c:pt idx="1"><c:v>45293</c:v></c:pt></c:numCache></c:numRef></c:cat>'
        );
        // the cells merged into another one are empty, as Excel reads them: no point
        expect(xml).to.include('<c:ptCount val="3"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numCache>');
        // dates among texts: yyyy-mm-dd, with the time when it has one
        expect(xml).to.include(
            '<c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>Total</c:v></c:pt><c:pt idx="1"><c:v>2024-01-02 03:04:05</c:v></c:pt><c:pt idx="2"><c:v>2024-01-03</c:v></c:pt></c:strCache>'
        );
    });

    it("reads the values when the workbook is written, in the date system of the workbook", async () => {
        const wb = new ExcelJS.Workbook();
        wb.properties.date1904 = true;
        const ws = wb.addWorksheet("S");
        ws.getCell("A1").value = new Date(Date.UTC(1904, 0, 2));
        ws.getCell("B1").value = 1;
        ws.addChart({ type: "column", anchor: "D2:H10", series: [{ categories: "A1", values: "B1" }] });
        ws.getCell("B1").value = 2;
        const xml = text(await wb.xlsx.writeBuffer(), "xl/charts/chart1.xml");
        expect(xml).to.include('<c:date1904 val="1"/>');
        expect(xml).to.include(
            '<c:f>S!$A$1</c:f><c:numCache><c:formatCode>mm-dd-yy</c:formatCode><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt>'
        );
        expect(xml).to.include(
            '<c:f>S!$B$1</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="1"/><c:pt idx="0"><c:v>2</c:v></c:pt>'
        );
    });

    it("writes the sheet names of the formulas quoted when needed, and follows a renamed worksheet", async () => {
        const wb = new ExcelJS.Workbook();
        const names = ["My Sheet", "It's", "Data_1.x", "A1", "R1C1", "1st", "Données"];
        for (const name of names) {
            wb.addWorksheet(name).getCell("B2").value = 1;
        }
        const ws = wb.addWorksheet("Chart");
        ws.addChart({
            type: "bar",
            anchor: "A1:H20",
            series: [
                { values: "'My Sheet'!$B$2" },
                { values: "'It''s'!B2" },
                { values: "Data_1.x!B2" },
                { values: "'A1'!B2" },
                { values: "'r1c1'!B2" },
                { values: "'1st'!B2" },
                { values: "Données!B2" },
            ],
        });
        wb.getWorksheet("Data_1.x").name = "Renamed";
        const xml = text(await wb.xlsx.writeBuffer(), "xl/charts/chart1.xml");
        const formulas = [...xml.matchAll(/<c:f>([^<]*)<\/c:f>/g)].map(([, formula]) => formula);
        expect(formulas).to.deep.equal([
            "&apos;My Sheet&apos;!$B$2",
            "&apos;It&apos;&apos;s&apos;!$B$2",
            "Renamed!$B$2",
            "&apos;A1&apos;!$B$2",
            "&apos;R1C1&apos;!$B$2",
            "&apos;1st&apos;!$B$2",
            "&apos;Données&apos;!$B$2",
        ]);
    });

    it("is carried as read once written and read back, the second write byte for byte the first", async () => {
        const { wb, ws } = dataWorkbook("My Data");
        ws.addChart(columnChart({ legend: { position: "b" } }));
        ws.addChart({
            type: "pie",
            anchor: { tl: { col: 5, row: 20 }, ext: { width: 320, height: 240 } },
            series: [{ categories: "A2:A5", values: "D2:D5" }],
        });
        const first = await wb.xlsx.writeBuffer();
        const read = await reload(first);
        expect(read.getWorksheet("My Data").getCharts()).to.deep.equal([]);
        const second = await read.xlsx.writeBuffer();
        const third = await (await reload(second)).xlsx.writeBuffer();
        for (const name of [
            "xl/charts/chart1.xml",
            "xl/charts/chart2.xml",
            "xl/drawings/drawing1.xml",
            "xl/drawings/_rels/drawing1.xml.rels",
            "[Content_Types].xml",
        ]) {
            expect(text(second, name), name).to.equal(text(first, name));
            expect(text(third, name), name).to.equal(text(second, name));
        }
    });

    it("writes created charts, carried charts and images in one drawing, with unique shape ids", async () => {
        const { wb, ws } = dataWorkbook();
        ws.addChart(columnChart());
        const read = await reload(await wb.xlsx.writeBuffer());
        const sheet = read.getWorksheet("Data");
        sheet.addImage(read.addImage({ buffer: png, extension: "png" }), "O2:P4");
        sheet.addChart(columnChart({ type: "line", title: "Added", anchor: "F20:M36" }));
        const buffer = await read.xlsx.writeBuffer();
        validateParts(buffer, "mixed");

        const drawing = text(buffer, "xl/drawings/drawing1.xml");
        // the picture first, then the carried chart, then the created one
        const ids = [...drawing.matchAll(/<xdr:cNvPr id="(\d+)" name="([^"]*)"/g)].map(([, id, name]) => `${id}:${name}`);
        expect(ids).to.deep.equal(["1:Picture 1", "2:Chart 1", "3:Chart 2"]);
        expect(drawing.indexOf("<xdr:pic>")).to.be.lessThan(drawing.indexOf("<c:chart "));
        const rels = text(buffer, "xl/drawings/_rels/drawing1.xml.rels");
        expect(rels).to.include('Target="../media/image1.png"');
        expect(rels).to.include('Target="../charts/chart1.xml"');
        expect(rels).to.include('Target="../charts/chart2.xml"');
        expect(text(buffer, "xl/charts/chart1.xml")).to.include("<a:t>Sales</a:t>");
        expect(text(buffer, "xl/charts/chart2.xml")).to.include("<a:t>Added</a:t>");
    });

    it("leaves the parts of a worksheet without chart unchanged", async () => {
        const build = (withChart) => {
            const { wb, ws } = dataWorkbook();
            const other = wb.addWorksheet("Other");
            other.getCell("A1").value = 1;
            if (withChart) {
                ws.addChart(columnChart());
            }
            return wb.xlsx.writeBuffer();
        };
        const [plain, charted] = await Promise.all([build(false), build(true)]);
        expect(text(charted, "xl/worksheets/sheet2.xml")).to.equal(text(plain, "xl/worksheets/sheet2.xml"));
        expect(text(charted, "xl/worksheets/_rels/sheet2.xml.rels")).to.equal(undefined);
        expect(text(plain, "xl/worksheets/sheet1.xml")).to.not.include("<drawing");
        expect(partNames(plain).some((name) => name.startsWith("xl/charts/") || name.startsWith("xl/drawings/"))).to.be.false();
        // a chart removed is not written
        const { wb, ws } = dataWorkbook();
        ws.addChart(columnChart()).remove();
        const removed = await wb.xlsx.writeBuffer();
        expect(text(removed, "xl/worksheets/sheet1.xml")).to.equal(text(plain, "xl/worksheets/sheet1.xml"));
        expect(partNames(removed).some((name) => name.startsWith("xl/charts/"))).to.be.false();
    });

    it("lists the charts added, and removes them", () => {
        const { ws } = dataWorkbook();
        const first = ws.addChart(columnChart());
        const second = ws.addChart(columnChart({ type: "area", title: undefined }));
        expect(ws.getCharts()).to.deep.equal([first, second]);
        expect(second.title).to.equal(undefined);
        first.remove();
        first.remove();
        expect(ws.getCharts()).to.deep.equal([second]);
    });

    it("is kept by a worksheet model copied to another worksheet", async () => {
        const { wb, ws } = dataWorkbook();
        ws.addChart(columnChart());
        const copy = wb.addWorksheet("Copy");
        copy.model = { ...ws.model, name: "Copy" };
        const buffer = await wb.xlsx.writeBuffer();
        // each worksheet has its own chart, both of the data of the original
        expect(text(buffer, "xl/charts/chart1.xml")).to.equal(text(buffer, "xl/charts/chart2.xml"));
        expect(text(buffer, "xl/charts/chart2.xml")).to.include("<c:f>Data!$B$2:$B$5</c:f>");
    });

    it("refuses to write a chart whose data worksheet was removed", async () => {
        const { wb } = dataWorkbook();
        const ws = wb.addWorksheet("Charts");
        const chart = ws.addChart(columnChart({ series: [{ values: "Data!B2:B5" }] }));
        wb.removeWorksheet("Data");
        let error;
        try {
            await wb.xlsx.writeBuffer();
        } catch (e) {
            error = e;
        }
        expect(error && error.message).to.equal(
            'A column chart of the worksheet "Charts" uses the data of the worksheet "Data", which was removed: remove the chart (chart.remove())'
        );
        chart.remove();
        await wb.xlsx.writeBuffer();
    });

    describe("invalid options", () => {
        const cases = [
            ["no options", undefined, /^Invalid chart options: undefined, expected an object$/],
            [
                "an unknown type",
                { type: "radar" },
                /^Invalid chart type: "radar", expected one of bar, column, line, pie, doughnut, area, scatter$/,
            ],
            [
                "a grouping of another type",
                { type: "line", grouping: "clustered" },
                /^Invalid grouping for a line chart: "clustered", expected one of standard, stacked, percentStacked$/,
            ],
            ["a grouping of a pie", { type: "pie", grouping: "stacked" }, /^A pie chart has no grouping$/],
            ["a marker of a bar chart", { type: "bar", marker: true }, /^The option marker does not apply to a bar chart$/],
            ["a line of a line chart", { type: "line", line: true }, /^The option line does not apply to a line chart$/],
            ["a scatter showing nothing", { type: "scatter", marker: false, line: false }, /^A scatter chart needs markers or lines/],
            ["a title that is not a string", { title: 1 }, /^Invalid chart title: 1, expected a string$/],
            ["a legend position", { legend: { position: "x" } }, /^Invalid chart legend position: "x", expected one of r, l, t, b, tr$/],
            ["a legend", { legend: true }, /^Invalid chart legend: true, expected \{ position \} or false$/],
            ["no series", { series: [] }, /^Invalid chart series: an array, expected an array of 1 to 255 series$/],
            [
                "too many series",
                { series: Array.from({ length: 256 }, () => ({ values: "B2:B5" })) },
                /expected an array of 1 to 255 series$/,
            ],
            ["a series without values", { series: [{ categories: "A2:A5" }] }, /^Invalid chart series 1: values is required/],
            [
                "x values on a column chart",
                { series: [{ values: "B2:B5", xValues: "A2:A5" }] },
                /^Invalid chart series 1: xValues is for scatter charts/,
            ],
            [
                "values on a scatter chart",
                { type: "scatter", series: [{ values: "B2:B5" }] },
                /^Invalid chart series 1: values is for other charts, a scatter chart takes xValues and yValues$/,
            ],
            [
                "a range of two columns",
                { series: [{ values: "B2:C5" }] },
                /^Invalid chart series 1 values: "B2:C5" is not one row or one column$/,
            ],
            [
                "a range out of the grid",
                { series: [{ values: "B2:B1048577" }] },
                /^Invalid chart series 1 values: "B2:B1048577", expected a range such as/,
            ],
            [
                "a column out of the grid",
                { series: [{ values: "XFE1:XFE2" }] },
                /^Invalid chart series 1 values: "XFE1:XFE2", expected a range/,
            ],
            ["whole columns", { series: [{ values: "B:B" }] }, /^Invalid chart series 1 values: "B:B", expected a range/],
            ["lower case", { series: [{ values: "b2:b5" }] }, /^Invalid chart series 1 values: "b2:b5", expected a range/],
            ["text after the range", { series: [{ values: "B2:B5x" }] }, /expected a range/],
            ["a misquoted sheet name", { series: [{ values: "'It's'!B2:B5" }] }, /expected a range/],
            [
                "an unknown worksheet",
                { series: [{ values: "Nope!B2:B5" }] },
                /^Invalid chart series 1 values: there is no worksheet "Nope"$/,
            ],
            [
                "categories of two rows",
                { series: [{ values: "B2:B5", categories: "A1:B2" }] },
                /^Invalid chart series 1 categories: "A1:B2" is not one row/,
            ],
            [
                "a name of two cells",
                { series: [{ values: "B2:B5", name: { ref: "B1:C1" } }] },
                /^Invalid chart series 1 name: the reference must be one cell$/,
            ],
            [
                "a name",
                { series: [{ values: "B2:B5", name: 3 }] },
                /^Invalid chart series 1 name: 3, expected a string or \{ ref: "B1" \}$/,
            ],
            [
                "a color",
                { series: [{ values: "B2:B5", color: "red" }] },
                /^Invalid chart series 1 color: "red", expected 6 hexadecimal digits RRGGBB$/,
            ],
            ["a series", { series: ["B2:B5"] }, /^Invalid chart series 1: "B2:B5", expected an object$/],
            ["the axes of a pie", { type: "pie", axes: {} }, /^A pie chart has no axes$/],
            ["an unknown axis", { axes: { x: {} } }, /^Invalid chart axis "x": a column chart has the axes category and value$/],
            ["a category axis bound", { axes: { category: { min: 0 } } }, /^The chart axis category is a category axis: it has no min$/],
            ["a bound", { axes: { value: { max: Infinity } } }, /^Invalid chart axis value max: Infinity, expected a finite number$/],
            [
                "bounds in reverse",
                { axes: { value: { min: 2, max: 1 } } },
                /^Invalid chart axis value: min \(2\) must be less than max \(1\)$/,
            ],
            ["a numFmt", { axes: { value: { numFmt: "" } } }, /^Invalid chart axis value numFmt: an empty string$/],
            [
                "gridlines",
                { axes: { value: { majorGridlines: "yes" } } },
                /^Invalid chart axis value majorGridlines: "yes", expected a boolean$/,
            ],
            ["no anchor", { anchor: undefined }, /^Invalid chart anchor: undefined, expected a range/],
            [
                "an anchor on a sheet",
                { anchor: "Data!F2:M18" },
                /^Invalid chart anchor: "Data!F2:M18", expected a range of cells such as "E2:L20"$/,
            ],
            [
                "an anchor with br and ext",
                { anchor: { tl: { col: 1, row: 1 }, br: { col: 2, row: 2 }, ext: { width: 1, height: 1 } } },
                /expected either br/,
            ],
            [
                "an anchor corner",
                { anchor: { tl: { col: -1, row: 1 }, br: { col: 2, row: 2 } } },
                /^Invalid chart anchor tl.col: -1, expected a number from 0 to less than 16384$/,
            ],
            [
                "a native corner",
                { anchor: { tl: { nativeCol: 1.5, nativeRow: 0 }, br: { col: 2, row: 2 } } },
                /^Invalid chart anchor tl.nativeCol: 1.5, expected an integer/,
            ],
            [
                "a corner below the grid",
                { anchor: { tl: { col: 1, row: 1048576 }, br: { col: 2, row: 2 } } },
                /^Invalid chart anchor tl.row: 1048576/,
            ],
            [
                "an offset",
                { anchor: { tl: { nativeCol: 1, nativeRow: 1, nativeColOff: -5 }, br: { col: 2, row: 2 } } },
                /^Invalid chart anchor tl.nativeColOff: -5/,
            ],
            ["br before tl", { anchor: { tl: { col: 3, row: 1 }, br: { col: 2, row: 5 } } }, /^Invalid chart anchor: br is left of tl$/],
            ["br above tl", { anchor: { tl: { col: 1, row: 3 }, br: { col: 2, row: 2 } } }, /^Invalid chart anchor: br is above tl$/],
            [
                "editAs",
                { anchor: { tl: { col: 1, row: 1 }, br: { col: 2, row: 2 }, editAs: "free" } },
                /^Invalid chart anchor editAs: "free", expected one of twoCell, oneCell, absolute$/,
            ],
            [
                "editAs with a size",
                { anchor: { tl: { col: 1, row: 1 }, ext: { width: 1, height: 1 }, editAs: "oneCell" } },
                /editAs applies to an anchor with two corners/,
            ],
            [
                "a size",
                { anchor: { tl: { col: 1, row: 1 }, ext: { width: 0, height: 1 } } },
                /^Invalid chart anchor ext width: 0, expected a positive number of pixels$/,
            ],
            [
                "a size beyond the coordinates of DrawingML",
                { anchor: { tl: { col: 1, row: 1 }, ext: { width: 1e20, height: 1 } } },
                /^Invalid chart anchor ext width: 100000000000000000000, expected a positive number of pixels up to 2863311529$/,
            ],
            [
                "an offset beyond the coordinates of DrawingML",
                { anchor: { tl: { nativeCol: 1, nativeRow: 1, nativeColOff: 1e21 }, br: { nativeCol: 3, nativeRow: 3 } } },
                /^Invalid chart anchor tl.nativeColOff: 1e\+21, expected an offset in EMU \(an integer from 0 to 27273042316900\)$/,
            ],
        ];
        for (const [what, overrides, message] of cases) {
            it(`refuses ${what}`, () => {
                const { ws } = dataWorkbook();
                const options = overrides === undefined ? undefined : { ...columnChart(), ...overrides };
                expect(() => ws.addChart(options)).to.throw(Error, message);
                expect(ws.getCharts()).to.deep.equal([]);
            });
        }

        it("decodes a long hostile range in linear time", () => {
            const { ws } = dataWorkbook();
            const hostile = `'${"''".repeat(50000)}x!${"$".repeat(50000)}B2:B5`;
            const start = Date.now();
            expect(() => ws.addChart({ ...columnChart(), series: [{ values: hostile }] })).to.throw(Error, /Invalid chart series 1 values/);
            expect(Date.now() - start).to.be.lessThan(50);
        });
    });

    it("refuses a corner whose offset in a huge column is beyond the coordinates of DrawingML", async () => {
        const { wb, ws } = dataWorkbook();
        ws.getColumn(2).width = 1e17;
        expect(() => ws.addChart({ ...columnChart(), anchor: { tl: { col: 1.5, row: 1 }, ext: { width: 10, height: 10 } } })).to.throw(
            Error,
            /^Invalid chart anchor tl: the offset in its cell \(500000000000000000000 EMU\) is beyond 27273042316900 EMU$/
        );
        // the largest size and offset are written without exponent, valid against the schema
        ws.addChart({
            ...columnChart(),
            anchor: { tl: { nativeCol: 1, nativeRow: 1, nativeRowOff: 27273042316900 }, ext: { width: 2863311529, height: 1 } },
        });
        const buffer = await wb.xlsx.writeBuffer();
        validateParts(buffer, "largest-anchor");
        const drawing = text(buffer, "xl/drawings/drawing1.xml");
        expect(drawing).to.include("<xdr:rowOff>27273042316900</xdr:rowOff>");
        expect(drawing).to.include('<xdr:ext cx="27273042313725" cy="9525"/>');
    });

    it("accepts the anchors of an image: corners, native corners with offsets, a corner and a size", async () => {
        const { wb, ws } = dataWorkbook();
        ws.addChart(
            columnChart({
                anchor: { tl: { col: 5.5, row: 1 }, br: { nativeCol: 12, nativeRow: 18, nativeColOff: 9525 }, editAs: "oneCell" },
            })
        );
        ws.addChart(columnChart({ anchor: { tl: { nativeCol: 5, nativeRow: 20 }, ext: { width: 480, height: 288 } } }));
        const drawing = text(await wb.xlsx.writeBuffer(), "xl/drawings/drawing1.xml");
        expect(drawing).to.include(
            '<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>5</xdr:col><xdr:colOff>320000</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>12</xdr:col><xdr:colOff>9525</xdr:colOff><xdr:row>18</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
        );
        expect(drawing).to.include(
            '<xdr:oneCellAnchor><xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>20</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="4572000" cy="2743200"/>'
        );
    });

    it("writes sample files of each type (spec/out/charts-samples) for Excel", async () => {
        fs.mkdirSync(SAMPLES, { recursive: true });
        const samples = {
            column: { type: "column", title: "Column (clustered)" },
            "column-stacked": { type: "column", grouping: "stacked", title: "Column (stacked)" },
            bar: { type: "bar", title: "Bar", legend: { position: "b" } },
            "bar-percent": { type: "bar", grouping: "percentStacked", title: "Bar (100%)" },
            line: { type: "line", title: "Line", axes: { value: { title: "Units", min: 0, max: 30 }, category: { title: "Month" } } },
            "line-no-markers": { type: "line", marker: false, title: "Line without markers" },
            area: { type: "area", grouping: "stacked", title: "Area (stacked)" },
            pie: { type: "pie", title: "Pie", series: [{ name: { ref: "D1" }, categories: "A2:A5", values: "D2:D5" }] },
            doughnut: { type: "doughnut", title: "Doughnut", legend: { position: "r" } },
            scatter: {
                type: "scatter",
                title: "Scatter",
                line: true,
                series: [{ name: "North vs South", xValues: "B2:B5", yValues: "C2:C5", color: "#C00000" }],
                axes: { x: { title: "North" }, y: { title: "South", numFmt: "0.0" } },
            },
        };
        for (const [name, options] of Object.entries(samples)) {
            const { wb, ws } = dataWorkbook();
            ws.addChart(columnChart(options));
            const buffer = await wb.xlsx.writeBuffer();
            validateParts(buffer, name);
            fs.writeFileSync(path.join(SAMPLES, `${name}.xlsx`), buffer);
        }

        // mixed: a chart carried from a file read, an image and created charts in one drawing, data on another sheet
        const { wb, ws } = dataWorkbook("Sales Data");
        ws.addChart(
            columnChart({
                title: "Carried (read from a file)",
                series: [{ name: { ref: "D1" }, categories: "A2:A5", values: "D2:D5", color: "4472C4" }],
            })
        );
        const read = await reload(await wb.xlsx.writeBuffer());
        const sheet = read.getWorksheet("Sales Data");
        sheet.addImage(read.addImage({ buffer: png, extension: "png" }), "O2:P4");
        const report = read.addWorksheet("Report");
        report.addChart({
            type: "line",
            title: "Created, data of 'Sales Data'",
            anchor: "A1:H16",
            series: [
                { name: { ref: "'Sales Data'!B1" }, categories: "'Sales Data'!A2:A5", values: "'Sales Data'!$B$2:$B$5", color: "ED7D31" },
                { name: { ref: "'Sales Data'!C1" }, categories: "'Sales Data'!A2:A5", values: "'Sales Data'!$C$2:$C$5" },
            ],
        });
        sheet.addChart(
            columnChart({
                type: "pie",
                title: "Created next to the carried chart",
                anchor: "F20:M36",
                series: [{ categories: "A2:A5", values: "B2:B5" }],
            })
        );
        const mixed = await read.xlsx.writeBuffer();
        validateParts(mixed, "mixed-sample");
        fs.writeFileSync(path.join(SAMPLES, "mixed.xlsx"), mixed);
        expect(partNames(mixed).filter((name) => /^xl\/charts\/chart\d+\.xml$/.test(name))).to.have.length(3);
    });
});
