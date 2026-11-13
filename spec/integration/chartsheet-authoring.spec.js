// Chartsheets created by workbook.addChartsheet: a sheet that is a single chart, its options validated as the charts of
// worksheet.addChart (src/doc/chart.ts), written through the code path of the chartsheets read from a file
// (src/xlsx/chartsheet-parts.ts): its part (createdChartsheetPart), a drawing with an absolute anchor, the chart part
// generated with the values of the cells when the workbook is written. Sample files for Excel are written to
// spec/out/chartsheet-samples/created-*.xlsx; with OOXML_SCHEMAS set (spec/utils/xmllint.js), the chartsheet, drawing
// and chart parts are validated against the schemas.
const fs = require("fs");
const path = require("path");
const { strFromU8 } = require("fflate");
const { unzipParts, zipParts } = require("../utils/xlsx-tamper");
const { validateDrawingParts } = require("../utils/xmllint");

const ExcelJS = verquire("exceljs");

const SAMPLES = path.join(__dirname, "../out/chartsheet-samples");
const EXCEL_FILE = path.join(__dirname, "data/chart-sheet.xlsx");
const REL_CHARTSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
const REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const CT_CHARTSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";
const CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml";
const CT_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

const textParts = (buffer) => Object.fromEntries(Object.entries(unzipParts(buffer)).map(([name, data]) => [name, strFromU8(data)]));

async function load(buffer) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    return wb;
}

// the chartsheet, drawing and chart parts (but those of `skip`: parts written by Excel, carried) against the schemas
function validate(buffer, label, skip = []) {
    const parts = unzipParts(buffer);
    for (const name of skip) {
        delete parts[name];
    }
    const result = validateDrawingParts(parts, path.join(SAMPLES, "xml", label));
    if (result) {
        expect(result.ok, `${label}: ${result.output}`).to.be.true();
        expect(result.count).to.be.greaterThan(0);
    }
}

// the <sheet> elements of a workbook part: [name, sheetId, state]
const sheetsOf = (parts) =>
    [...parts["xl/workbook.xml"].matchAll(/<sheet sheetId="(\d+)" name="([^"]*)"(?: state="(\w+)")?/g)].map(([, id, name, state]) => [
        name,
        Number(id),
        state ?? "visible",
    ]);

function dataWorkbook() {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Data");
    ws.addRow(["Month", "North", "South"]);
    ws.addRow(["Jan", 10, 7]);
    ws.addRow(["Feb", 12, 9]);
    ws.addRow(["Mar", 9, 11]);
    return { wb, ws };
}

const COLUMN = {
    type: "column",
    title: "Sales",
    series: [
        { name: { ref: "Data!B1" }, categories: "Data!A2:A4", values: "Data!$B$2:$B$4" },
        { name: { ref: "Data!C1" }, categories: "Data!A2:A4", values: "Data!C2:C4", color: "FF9900" },
    ],
};

describe("Chartsheets created by workbook.addChartsheet", () => {
    it("writes the chartsheet, its drawing and its chart, appended to the sheets", async () => {
        const { wb } = dataWorkbook();
        const added = wb.addChartsheet("Sales Chart", COLUMN);
        expect(added).to.deep.equal({ id: 2, name: "Sales Chart", state: "visible" });
        expect(wb.chartsheets).to.deep.equal([added]);
        // a worksheet added afterwards comes after it, with an id of its own
        expect(wb.addWorksheet("After").id).to.equal(3);
        expect(wb.worksheets.map((ws) => ws.name)).to.deep.equal(["Data", "After"]);

        const buffer = await wb.xlsx.writeBuffer();
        validate(buffer, "created-basic");
        const parts = textParts(buffer);
        expect(sheetsOf(parts)).to.deep.equal([
            ["Data", 1, "visible"],
            ["Sales Chart", 2, "visible"],
            ["After", 3, "visible"],
        ]);
        expect(parts["xl/_rels/workbook.xml.rels"]).to.include(`Type="${REL_CHARTSHEET}" Target="chartsheets/sheet1.xml"`);
        expect(parts["xl/chartsheets/sheet1.xml"]).to.equal(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr/><sheetViews><sheetView zoomToFit="1" workbookViewId="0"/></sheetViews><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><drawing r:id="rId1"/></chartsheet>'
        );
        expect(parts["xl/chartsheets/_rels/sheet1.xml.rels"]).to.include(
            `Id="rId1" Type="${REL_DRAWING}" Target="../drawings/drawing1.xml"`
        );
        expect(parts["xl/drawings/drawing1.xml"]).to.include(
            '<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="9304587" cy="6077107"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="1" name="Chart 1"/>'
        );
        expect(parts["xl/drawings/_rels/drawing1.xml.rels"]).to.include('Target="../charts/chart1.xml"');
        const contentTypes = parts["[Content_Types].xml"];
        expect(contentTypes).to.include(`<Override PartName="/xl/chartsheets/sheet1.xml" ContentType="${CT_CHARTSHEET}"/>`);
        expect(contentTypes).to.include(`<Override PartName="/xl/drawings/drawing1.xml" ContentType="${CT_DRAWING}"/>`);
        expect(contentTypes).to.include(`<Override PartName="/xl/charts/chart1.xml" ContentType="${CT_CHART}"/>`);
        expect(parts["docProps/app.xml"]).to.include("<vt:lpstr>Data</vt:lpstr><vt:lpstr>After</vt:lpstr><vt:lpstr>Sales Chart</vt:lpstr>");
        const chart = parts["xl/charts/chart1.xml"];
        expect(chart).to.include('<c:f>Data!$C$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>South</c:v>');
        expect(chart).to.include('<c:pt idx="2"><c:v>11</c:v></c:pt>');
        expect(chart).to.include('<a:srgbClr val="FF9900"/>');
    });

    it("reads the values when the workbook is written; refuses to write data of a removed worksheet", async () => {
        const { wb, ws } = dataWorkbook();
        const other = wb.addWorksheet("Other");
        other.getCell("A1").value = 5;
        wb.addChartsheet("Chart", { type: "pie", series: [{ values: "Other!A1:A2" }] });
        other.getCell("A2").value = 6;
        ws.getCell("B2").value = 99;
        let parts = textParts(await wb.xlsx.writeBuffer());
        expect(parts["xl/charts/chart1.xml"]).to.include('<c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>6</c:v></c:pt>');
        // a renamed worksheet: the formula follows it
        other.name = "My Other";
        parts = textParts(await wb.xlsx.writeBuffer());
        expect(parts["xl/charts/chart1.xml"]).to.include("<c:f>&apos;My Other&apos;!$A$1:$A$2</c:f>");
        wb.removeWorksheet("My Other");
        await wb.xlsx.writeBuffer().then(
            () => expect.fail("written"),
            (error) =>
                expect(error.message).to.equal(
                    'The chartsheet "Chart" uses the data of the worksheet "My Other", which was removed: remove the chartsheet (workbook.removeChartsheet("Chart"))'
                )
        );
        wb.removeChartsheet("Chart");
        expect(wb.chartsheets).to.deep.equal([]);
        parts = textParts(await wb.xlsx.writeBuffer());
        expect(Object.keys(parts).filter((name) => /^xl\/(charts|drawings|chartsheets)\//.test(name))).to.deep.equal([]);
    });

    it("takes the states, and every type of chart", async () => {
        const { wb } = dataWorkbook();
        wb.addChartsheet("Hidden", { type: "line", marker: false, series: [{ values: "Data!B2:B4" }] }, { state: "hidden" });
        wb.addChartsheet("Very", { type: "scatter", series: [{ xValues: "Data!B2:B4", yValues: "Data!C2:C4" }] }, { state: "veryHidden" });
        wb.addChartsheet("Shown", { type: "doughnut", legend: false, series: [{ values: "Data!B2:B4" }] }, {});
        wb.addChartsheet("Bars", { type: "bar", grouping: "percentStacked", series: [{ values: "Data!B2:B4" }, { values: "Data!C2:C4" }] });
        wb.addChartsheet("Area", { type: "area", axes: { value: { min: 0, max: 20, numFmt: "0" } }, series: [{ values: "Data!B2:B4" }] });
        const buffer = await wb.xlsx.writeBuffer();
        validate(buffer, "created-types");
        const parts = textParts(buffer);
        expect(sheetsOf(parts).map(([name, , state]) => [name, state])).to.deep.equal([
            ["Data", "visible"],
            ["Hidden", "hidden"],
            ["Very", "veryHidden"],
            ["Shown", "visible"],
            ["Bars", "visible"],
            ["Area", "visible"],
        ]);
        expect(parts["xl/charts/chart5.xml"]).to.include("<c:areaChart>");
        expect((await load(buffer)).chartsheets.map(({ name, state }) => `${name}:${state}`)).to.deep.equal([
            "Hidden:hidden",
            "Very:veryHidden",
            "Shown:visible",
            "Bars:visible",
            "Area:visible",
        ]);
    });

    it("is carried once read back, the second write byte for byte the first", async () => {
        const { wb, ws } = dataWorkbook();
        ws.addChart({ ...COLUMN, anchor: "E2:L18" });
        wb.addChartsheet("Chart", COLUMN);
        const buffer = await wb.xlsx.writeBuffer();
        const loaded = await load(buffer);
        expect(loaded.chartsheets).to.deep.equal([{ id: 2, name: "Chart", state: "visible" }]);
        const first = unzipParts(await loaded.xlsx.writeBuffer());
        const second = unzipParts(await (await load(zipParts(first))).xlsx.writeBuffer());
        const charted = (all) =>
            Object.keys(all)
                .filter((name) => /^xl\/(charts|drawings|chartsheets)\//.test(name))
                .sort();
        expect(charted(first)).to.deep.equal(charted(textParts(buffer)));
        expect(charted(second)).to.deep.equal(charted(first));
        for (const name of charted(first)) {
            expect(strFromU8(second[name]), name).to.equal(strFromU8(first[name]));
            expect(strFromU8(first[name]), name).to.equal(textParts(buffer)[name]);
        }
    });

    it("is placed after the chartsheets of a file read, its drawing numbered after theirs", async () => {
        const wb = await load(fs.readFileSync(EXCEL_FILE));
        wb.getWorksheet("Sheet1").getCell("H1").value = 1;
        const added = wb.addChartsheet("Created", { type: "line", series: [{ values: "Sheet1!H1:H2" }] });
        expect(added.id).to.equal(2);
        expect(wb.chartsheets.map(({ name }) => name)).to.deep.equal(["Chart1", "Created"]);
        expect(() => wb.addChartsheet("chart1", COLUMN)).to.throw("Worksheet name already exists: chart1");
        const buffer = await wb.xlsx.writeBuffer();
        // the chart of Excel has markup compatibility elements (mc:AlternateContent), which the schema does not know
        validate(buffer, "created-after-read", ["xl/charts/chart1.xml"]);
        const parts = textParts(buffer);
        expect(sheetsOf(parts).map(([name, id]) => [name, id])).to.deep.equal([
            ["Chart1", 9],
            ["Sheet1", 1],
            ["Created", 2],
        ]);
        expect(parts["xl/chartsheets/_rels/sheet2.xml.rels"]).to.include("../drawings/drawing2.xml");
        expect(parts["xl/drawings/drawing2.xml"]).to.include("<xdr:absoluteAnchor>");
        expect(parts["xl/drawings/_rels/drawing2.xml.rels"]).to.include("../charts/chart2.xml");
    });

    it("keeps the names unique among the worksheets and the chartsheets (case-insensitive)", () => {
        const { wb, ws } = dataWorkbook();
        wb.addChartsheet("Chart", COLUMN);
        const cases = [
            [() => wb.addChartsheet("DATA", COLUMN), "Worksheet name already exists: DATA"],
            [() => wb.addChartsheet("chart", COLUMN), "Worksheet name already exists: chart"],
            [() => wb.addWorksheet("CHART"), "Worksheet name already exists: CHART"],
            [() => (ws.name = "Chart"), "Worksheet name already exists: Chart"],
            [() => wb.addChartsheet("", COLUMN), "The name can't be empty."],
            [() => wb.addChartsheet("History", COLUMN), 'The name "History" is protected'],
            [() => wb.addChartsheet("a[1]", COLUMN), "cannot include any of the following characters"],
            [() => wb.addChartsheet("x'", COLUMN), "single quotation mark"],
            [() => wb.addChartsheet(42, COLUMN), "The name has to be a string."],
        ];
        for (const [call, message] of cases) {
            expect(call).to.throw(message);
        }
        // a name of more than 31 characters is truncated (with a warning), as a worksheet name
        const warnings = [];
        const { warn } = console;
        console.warn = (message) => warnings.push(message);
        let long;
        try {
            long = wb.addChartsheet("A chartsheet with a very long name indeed", COLUMN);
        } finally {
            console.warn = warn;
        }
        expect(warnings).to.deep.equal([
            "Worksheet name A chartsheet with a very long name indeed exceeds 31 chars. This will be truncated",
        ]);
        expect(long.name).to.equal("A chartsheet with a very long n");
        expect(wb.chartsheets.map(({ name }) => name)).to.deep.equal(["Chart", "A chartsheet with a very long n"]);
        // a chartsheet removed frees its name
        wb.removeChartsheet("Chart");
        expect(wb.addWorksheet("Chart").name).to.equal("Chart");
    });

    it("refuses invalid chart options and states, adding nothing", () => {
        const { wb } = dataWorkbook();
        const cases = [
            [{ type: "line", series: [{ values: "B2:B4" }] }, 'Invalid chart series 1 values: "B2:B4" names no worksheet'],
            [
                { type: "line", series: [{ values: "Data!B2:B4", name: { ref: "C1" } }] },
                'Invalid chart series 1 name: "C1" names no worksheet',
            ],
            [{ ...COLUMN, anchor: "A1:B2" }, "The chart of a chartsheet has no anchor: it fills the sheet"],
            [{ type: "line", series: [{ values: "Nope!B2:B4" }] }, 'Invalid chart series 1 values: there is no worksheet "Nope"'],
            [
                { type: "line", series: [{ values: "Data!B2:C4" }] },
                'Invalid chart series 1 values: "Data!B2:C4" is not one row or one column',
            ],
            [{ type: "radar", series: [{ values: "Data!B2:B4" }] }, "Invalid chart type"],
            [undefined, "Invalid chart options: undefined, expected an object"],
        ];
        for (const [options, message] of cases) {
            expect(() => wb.addChartsheet("New", options)).to.throw(message);
        }
        expect(() => wb.addChartsheet("New", COLUMN, { state: "invisible" })).to.throw(
            'Invalid chartsheet state: "invisible", expected one of visible, hidden, veryHidden'
        );
        expect(() => wb.addChartsheet("New", COLUMN, null)).to.throw("Invalid chartsheet options: expected an object");
        expect(wb.chartsheets).to.deep.equal([]);
    });

    it("bounds a hostile range: a whole column, read in time linear in the cells", async function () {
        this.timeout(10000);
        const { wb } = dataWorkbook();
        const started = Date.now();
        wb.addChartsheet("Chart", { type: "line", series: [{ values: "Data!B1:B1048576" }, { values: "Data!A1:XFD1" }] });
        const parts = textParts(await wb.xlsx.writeBuffer());
        expect(Date.now() - started).to.be.below(5000);
        expect(parts["xl/charts/chart1.xml"]).to.include('<c:ptCount val="1048576"/>').and.include('<c:ptCount val="16384"/>');
    });

    it("writes sample files (spec/out/chartsheet-samples/created-*.xlsx) for Excel", async () => {
        fs.mkdirSync(SAMPLES, { recursive: true });
        const samples = {
            "created-column": (wb) => wb.addChartsheet("Sales", COLUMN),
            "created-pie-hidden-data": (wb) => {
                wb.getWorksheet("Data").state = "hidden";
                wb.addChartsheet("Share", { type: "pie", title: "Share", series: [{ categories: "Data!A2:A4", values: "Data!B2:B4" }] });
            },
            "created-scatter-and-worksheet-chart": (wb) => {
                wb.getWorksheet("Data").addChart({ ...COLUMN, anchor: "E2:L18" });
                wb.addChartsheet("Scatter", {
                    type: "scatter",
                    line: true,
                    title: "North vs South",
                    axes: { x: { title: "North" }, y: { title: "South" } },
                    series: [{ name: "Points", xValues: "Data!B2:B4", yValues: "Data!C2:C4", color: "2E7D32" }],
                });
                wb.addWorksheet("Last").getCell("A1").value = "after the chartsheet";
            },
        };
        for (const [name, build] of Object.entries(samples)) {
            const { wb } = dataWorkbook();
            build(wb);
            const buffer = await wb.xlsx.writeBuffer();
            validate(buffer, name);
            fs.writeFileSync(path.join(SAMPLES, `${name}.xlsx`), buffer);
            await load(buffer);
        }
    });
});
