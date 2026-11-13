// Chartsheets (sheets that are a single chart: xl/chartsheets/sheetN.xml → drawing → chart) are carried from a file read
// to the file written, in their places among the worksheets. The fixtures are spec/integration/data/chart-sheet.xlsx (made
// by Excel) and workbooks built in code: a workbook written by ExcelJS, to which chartsheets are added (xlsx-tamper), the
// chart and drawing parts taken from the Excel file.
const fs = require("fs");
const path = require("path");
const { strFromU8, strToU8 } = require("fflate");
const { baseWorkbookBuffer, editParts, unzipParts } = require("../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");
const { loadChartsheets, mergeSheets, MAX_CHARTSHEETS } = verquire("xlsx/chartsheet-parts");
const ChartsheetXform = verquire("xlsx/xform/chartsheet/chartsheet-xform");

const EXCEL_FILE = path.join(__dirname, "data/chart-sheet.xlsx");
const excelParts = unzipParts(fs.readFileSync(EXCEL_FILE));
const excelPart = (name) => strFromU8(excelParts[name]);

const NS = {
    main: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    pkg: "http://schemas.openxmlformats.org/package/2006/relationships",
};
const REL = {
    chartsheet: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
    worksheet: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
    drawing: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
    chart: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
    styles: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
    theme: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
    sharedStrings: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
    image: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
};
const CT = {
    chartsheet: "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
    drawing: "application/vnd.openxmlformats-officedocument.drawing+xml",
    chart: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
};
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

function rels(list) {
    const items = list.map(([id, type, target]) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`).join("");
    return `${XML_DECL}<Relationships xmlns="${NS.pkg}">${items}</Relationships>`;
}

// a chartsheet part: `inner` goes before <drawing> (CT_Chartsheet: sheetPr ... headerFooter), `after` after it
function chartsheetXml({
    inner = '<sheetViews><sheetView workbookViewId="0" zoomToFit="1"/></sheetViews>',
    after = "",
    drawing = '<drawing r:id="rId1"/>',
} = {}) {
    return `${XML_DECL}<chartsheet xmlns="${NS.main}" xmlns:r="${NS.r}">${inner}${drawing}${after}</chartsheet>`;
}

const escapeXml = (text) => text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");

/**
 * A workbook whose sheets are, in order, `sheets`: `{ worksheet: name }` (cells A1 = name, B1:B3 = 1, 2, 3) or
 * `{ chartsheet: name, state, xml, rels, drawing }` (the drawing and the chart of chart-sheet.xlsx). `workbookXml(xml)`
 * edits xl/workbook.xml (the sheets, defined names, views...), `edit(parts)` any part (strings) before zipping.
 */
async function chartsheetWorkbook(sheets, { workbookXml = (xml) => xml, edit } = {}) {
    const worksheetNames = sheets.filter((sheet) => sheet.worksheet).map((sheet) => sheet.worksheet);
    const base = await baseWorkbookBuffer((wb, ws) => {
        ws.name = worksheetNames[0];
        for (const name of worksheetNames.slice(1)) {
            wb.addWorksheet(name).getCell("A1").value = name;
        }
        for (const worksheet of wb.worksheets) {
            worksheet.getCell("B1").value = 1;
            worksheet.getCell("B2").value = 2;
            worksheet.getCell("B3").value = 3;
        }
    });
    return editParts(base, (parts) => {
        const text = {};
        const workbookRels = [
            ["rId1", REL.styles, "styles.xml"],
            ["rId2", REL.theme, "theme/theme1.xml"],
            ["rId3", REL.sharedStrings, "sharedStrings.xml"],
        ];
        const sheetElements = [];
        let overrides = "";
        let worksheetNo = 0;
        let chartsheetNo = 0;
        sheets.forEach((sheet, i) => {
            const rId = `rId${10 + i}`;
            const state = sheet.state ? ` state="${sheet.state}"` : "";
            if (sheet.worksheet) {
                worksheetNo++;
                workbookRels.push([rId, REL.worksheet, `worksheets/sheet${worksheetNo}.xml`]);
                sheetElements.push(`<sheet name="${escapeXml(sheet.worksheet)}" sheetId="${worksheetNo}"${state} r:id="${rId}"/>`);
                return;
            }
            const n = ++chartsheetNo;
            workbookRels.push([rId, REL.chartsheet, sheet.target ?? `chartsheets/sheet${n}.xml`]);
            sheetElements.push(
                `<sheet name="${escapeXml(sheet.chartsheet)}" sheetId="${sheet.sheetId ?? 100 + n}"${state} r:id="${rId}"/>`
            );
            text[`xl/chartsheets/sheet${n}.xml`] = sheet.xml ?? chartsheetXml();
            text[`xl/chartsheets/_rels/sheet${n}.xml.rels`] = sheet.rels ?? rels([["rId1", REL.drawing, `../drawings/drawing${n}.xml`]]);
            text[`xl/drawings/drawing${n}.xml`] = sheet.drawing ?? excelPart("xl/drawings/drawing1.xml");
            text[`xl/drawings/_rels/drawing${n}.xml.rels`] = rels([["rId1", REL.chart, `../charts/chart${n}.xml`]]);
            text[`xl/charts/chart${n}.xml`] = excelPart("xl/charts/chart1.xml");
            overrides +=
                `<Override PartName="/xl/chartsheets/sheet${n}.xml" ContentType="${CT.chartsheet}"/>` +
                `<Override PartName="/xl/drawings/drawing${n}.xml" ContentType="${CT.drawing}"/>` +
                `<Override PartName="/xl/charts/chart${n}.xml" ContentType="${CT.chart}"/>`;
        });
        text["xl/_rels/workbook.xml.rels"] = rels(workbookRels);
        const workbook = strFromU8(parts["xl/workbook.xml"]).replace(/<sheets>.*<\/sheets>/, `<sheets>${sheetElements.join("")}</sheets>`);
        text["xl/workbook.xml"] = workbookXml(workbook);
        text["[Content_Types].xml"] = strFromU8(parts["[Content_Types].xml"]).replace("</Types>", `${overrides}</Types>`);
        for (const [name, value] of Object.entries(text)) {
            parts[name] = strToU8(value);
        }
        if (edit) {
            for (const [name, value] of Object.entries(parts)) {
                parts[name] = strFromU8(value);
            }
            edit(parts);
            for (const [name, value] of Object.entries(parts)) {
                if (typeof value === "string") {
                    parts[name] = strToU8(value);
                }
            }
        }
    });
}

// sets the workbook view (ExcelJS writes none for a workbook without views)
const bookView = (attributes) => (xml) =>
    xml.replace(
        /(<bookViews>.*<\/bookViews>)?<sheets>/,
        `<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="100" windowHeight="100" ${attributes}/></bookViews><sheets>`
    );

// a file written, for a manual check in Excel (spec/out is ignored by git)
function writeSample(name, buffer) {
    const dir = path.join(__dirname, "../out/chartsheet-samples");
    fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(path.join(dir, name), buffer);
}

async function load(buffer) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    return wb;
}

// the text parts of a file written
function textParts(buffer) {
    const parts = unzipParts(buffer);
    const text = {};
    for (const [name, data] of Object.entries(parts)) {
        text[name] = strFromU8(data);
    }
    return text;
}

// the <sheet> elements of a workbook part, as [name, sheetId, state, rel type]
function writtenSheets(parts) {
    const rIds = new Map();
    for (const [, id, type] of parts["xl/_rels/workbook.xml.rels"].matchAll(/<Relationship Id="([^"]+)" Type="([^"]+)"/g)) {
        rIds.set(id, type.slice(type.lastIndexOf("/") + 1));
    }
    return [...parts["xl/workbook.xml"].matchAll(/<sheet sheetId="(\d+)" name="([^"]*)"(?: state="(\w+)")? r:id="([^"]+)"\/>/g)].map(
        ([, id, name, state, rId]) => [name, Number(id), state ?? "visible", rIds.get(rId)]
    );
}

describe("Chartsheets", () => {
    describe("chart-sheet.xlsx (made by Excel)", () => {
        it("keeps the chartsheet, in its place before the worksheet", async () => {
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.readFile(EXCEL_FILE);
            expect(wb.worksheets.map((ws) => ws.name)).to.deep.equal(["Sheet1"]);
            expect(wb.chartsheets).to.deep.equal([{ id: 9, name: "Chart1", state: "visible" }]);

            const buffer = await wb.xlsx.writeBuffer();
            writeSample("chart-sheet.xlsx", buffer);

            const parts = textParts(buffer);
            expect(writtenSheets(parts)).to.deep.equal([
                ["Chart1", 9, "visible", "chartsheet"],
                ["Sheet1", 1, "visible", "worksheet"],
            ]);
            // the chartsheet as Excel wrote it (but the line end of the XML declaration)
            const body = (xml) => xml.slice(xml.indexOf("<chartsheet"));
            expect(body(parts["xl/chartsheets/sheet1.xml"])).to.equal(body(excelPart("xl/chartsheets/sheet1.xml")));
            expect(parts["xl/chartsheets/_rels/sheet1.xml.rels"]).to.equal(rels([["rId1", REL.drawing, "../drawings/drawing1.xml"]]));
            expect(parts["xl/drawings/drawing1.xml"]).to.include("<xdr:absoluteAnchor>").and.include('r:id="rId1"');
            expect(parts["xl/drawings/_rels/drawing1.xml.rels"]).to.equal(rels([["rId1", REL.chart, "../charts/chart1.xml"]]));
            // the parts of the chart, as read
            for (const name of ["xl/charts/chart1.xml", "xl/charts/style1.xml", "xl/charts/colors1.xml"]) {
                expect(parts[name]).to.equal(excelPart(name));
            }
            const contentTypes = parts["[Content_Types].xml"];
            expect(contentTypes).to.include(`<Override PartName="/xl/chartsheets/sheet1.xml" ContentType="${CT.chartsheet}"/>`);
            expect(contentTypes).to.include(`<Override PartName="/xl/drawings/drawing1.xml" ContentType="${CT.drawing}"/>`);
            expect(contentTypes).to.include(`<Override PartName="/xl/charts/chart1.xml" ContentType="${CT.chart}"/>`);
            // docProps/app.xml lists the worksheets, then the charts (as Excel)
            expect(parts["docProps/app.xml"]).to.include(
                '<HeadingPairs><vt:vector size="4" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant><vt:variant><vt:lpstr>Charts</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="2" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr><vt:lpstr>Chart1</vt:lpstr></vt:vector></TitlesOfParts>'
            );
        });

        it("writes the same parts again after a second round trip", async () => {
            const first = await (await load(fs.readFileSync(EXCEL_FILE))).xlsx.writeBuffer();
            const second = await (await load(first)).xlsx.writeBuffer();
            const [a, b] = [textParts(first), textParts(second)];
            delete a["docProps/core.xml"];
            delete b["docProps/core.xml"];
            expect(b).to.deep.equal(a);
        });

        it("is read by the streaming reader (which skips the chartsheet)", async () => {
            const reader = new ExcelJS.stream.xlsx.WorkbookReader(EXCEL_FILE, {});
            const names = [];
            for await (const worksheet of reader) {
                names.push(worksheet.name);
                for await (const row of worksheet) {
                    expect(row.number).to.be.a("number");
                }
            }
            expect(names).to.deep.equal(["Sheet1"]);
        });

        it("is kept by a ModelContainer and by workbook.model", async () => {
            const container = new ExcelJS.ModelContainer();
            await container.xlsx.load(fs.readFileSync(EXCEL_FILE));
            let parts = textParts(await container.xlsx.writeBuffer());
            expect(writtenSheets(parts).map(([name, , , type]) => [name, type])).to.deep.equal([
                ["Chart1", "chartsheet"],
                ["Sheet1", "worksheet"],
            ]);

            const copy = new ExcelJS.Workbook();
            copy.model = (await load(fs.readFileSync(EXCEL_FILE))).model;
            expect(copy.chartsheets.map((chartsheet) => chartsheet.name)).to.deep.equal(["Chart1"]);
            copy.addWorksheet("Added");
            parts = textParts(await copy.xlsx.writeBuffer());
            expect(writtenSheets(parts).map(([name]) => name)).to.deep.equal(["Chart1", "Sheet1", "Added"]);
            expect(parts["xl/charts/chart1.xml"]).to.equal(excelPart("xl/charts/chart1.xml"));
        });
    });

    describe("worksheet API", () => {
        it("never returns a chartsheet", async () => {
            const wb = await load(fs.readFileSync(EXCEL_FILE));
            expect(wb.getWorksheet("Chart1")).to.equal(undefined);
            expect(wb.getWorksheet(9)).to.equal(undefined);
            const visited = [];
            wb.eachSheet((ws) => visited.push(ws.name));
            expect(visited).to.deep.equal(["Sheet1"]);
            expect(wb.getWorksheet().name).to.equal("Sheet1");
        });

        it("refuses a worksheet named as a chartsheet (case-insensitive)", async () => {
            const wb = await load(fs.readFileSync(EXCEL_FILE));
            expect(() => wb.addWorksheet("chart1")).to.throw("Worksheet name already exists: chart1");
            expect(() => {
                wb.getWorksheet("Sheet1").name = "CHART1";
            }).to.throw("Worksheet name already exists: CHART1");
            expect(wb.worksheets.map((ws) => ws.name)).to.deep.equal(["Sheet1"]);
        });

        it("chartsheets is a copy", async () => {
            const wb = await load(fs.readFileSync(EXCEL_FILE));
            const [chartsheet] = wb.chartsheets;
            chartsheet.name = "Changed";
            wb.chartsheets.pop();
            expect(wb.chartsheets).to.deep.equal([{ id: 9, name: "Chart1", state: "visible" }]);
            expect(new ExcelJS.Workbook().chartsheets).to.deep.equal([]);
        });

        it("adds a worksheet after the chartsheets, with a sheet id of its own", async () => {
            const buffer = await chartsheetWorkbook([{ chartsheet: "Chart1", sheetId: 2 }, { worksheet: "Sheet1" }]);
            const wb = await load(buffer);
            // the id of the chartsheet is not free (nextId skips it)
            expect(wb.addWorksheet("Two").id).to.equal(3);
            const sheets = writtenSheets(textParts(await wb.xlsx.writeBuffer()));
            expect(sheets.map(([name, , , type]) => [name, type])).to.deep.equal([
                ["Chart1", "chartsheet"],
                ["Sheet1", "worksheet"],
                ["Two", "worksheet"],
            ]);
            // the sheet ids are unique, as Excel requires: the chartsheet keeps its own
            expect(sheets.map(([, id]) => id)).to.deep.equal([2, 1, 3]);
        });

        it("a chartsheet with the sheet id of a worksheet gets a free one when written", async () => {
            const buffer = await chartsheetWorkbook([
                { worksheet: "Sheet1" },
                { chartsheet: "Chart1", sheetId: 1 },
                { worksheet: "Sheet2" },
            ]);
            const sheets = writtenSheets(textParts(await (await load(buffer)).xlsx.writeBuffer()));
            expect(sheets.map(([name, id]) => [name, id])).to.deep.equal([
                ["Sheet1", 1],
                ["Chart1", 3],
                ["Sheet2", 2],
            ]);
        });

        it("removeChartsheet: the chartsheet is not written", async () => {
            const wb = await load(fs.readFileSync(EXCEL_FILE));
            wb.removeChartsheet("Unknown");
            expect(wb.chartsheets).to.have.length(1);
            wb.removeChartsheet("Chart1");
            expect(wb.chartsheets).to.deep.equal([]);
            wb.addWorksheet("Chart1");
            const parts = textParts(await wb.xlsx.writeBuffer());
            expect(Object.keys(parts).filter((name) => name.includes("chart"))).to.deep.equal([]);
            expect(writtenSheets(parts).map(([name]) => name)).to.deep.equal(["Sheet1", "Chart1"]);
            expect(parts["docProps/app.xml"]).to.not.include("Charts");
        });
    });

    describe("order, state and settings", () => {
        it("keeps worksheets and chartsheets interleaved, their states and the active tab", async () => {
            const buffer = await chartsheetWorkbook(
                [
                    { worksheet: "Sheet1" },
                    { chartsheet: "Chart A", state: "hidden" },
                    { worksheet: "Sheet2" },
                    { chartsheet: "Chart B", state: "veryHidden" },
                    { chartsheet: "Chart & <C>" },
                ],
                { workbookXml: bookView('firstSheet="1" activeTab="4"') }
            );
            const wb = await load(buffer);
            expect(wb.chartsheets.map(({ name, state }) => [name, state])).to.deep.equal([
                ["Chart A", "hidden"],
                ["Chart B", "veryHidden"],
                ["Chart & <C>", "visible"],
            ]);
            expect(wb.views[0]).to.include({ firstSheet: 1, activeTab: 4 });
            const written = await wb.xlsx.writeBuffer();
            writeSample("interleaved.xlsx", written);
            const parts = textParts(written);
            expect(writtenSheets(parts)).to.deep.equal([
                ["Sheet1", 1, "visible", "worksheet"],
                ["Chart A", 101, "hidden", "chartsheet"],
                ["Sheet2", 2, "visible", "worksheet"],
                ["Chart B", 102, "veryHidden", "chartsheet"],
                ["Chart &amp; &lt;C&gt;", 103, "visible", "chartsheet"],
            ]);
            expect(parts["xl/workbook.xml"]).to.include('firstSheet="1" activeTab="4"');
            for (const n of [1, 2, 3]) {
                // the active tab (4) is the third chartsheet: its tab is selected
                const selected = n === 3 ? '<sheetView tabSelected="1" workbookViewId="0" zoomToFit="1"/>' : undefined;
                expect(parts[`xl/chartsheets/sheet${n}.xml`]).to.equal(
                    selected ? chartsheetXml({ inner: `<sheetViews>${selected}</sheetViews>` }) : chartsheetXml()
                );
                // numbered after the drawings of the worksheets (none here)
                expect(parts[`xl/chartsheets/_rels/sheet${n}.xml.rels`]).to.include(`Target="../drawings/drawing${n}.xml"`);
            }
            expect(parts["docProps/app.xml"]).to.include(
                "<vt:lpstr>Sheet1</vt:lpstr><vt:lpstr>Sheet2</vt:lpstr><vt:lpstr>Chart A</vt:lpstr><vt:lpstr>Chart B</vt:lpstr><vt:lpstr>Chart &amp; &lt;C&gt;</vt:lpstr>"
            );
        });

        it("keeps the tab color, protection, page setup and header of the chartsheet", async () => {
            const inner =
                '<sheetPr codeName="Chart9"><tabColor rgb="FFFF0000"/></sheetPr><sheetViews><sheetView zoomScale="80" workbookViewId="0"/></sheetViews>' +
                '<sheetProtection content="1" objects="1"/><pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5" header="0.2" footer="0.2"/>' +
                '<pageSetup orientation="landscape" r:id="rId2"/><headerFooter><oddHeader>&amp;CQ &amp;&amp; A &lt;1&gt;</oddHeader></headerFooter>';
            const xml = chartsheetXml({ inner });
            const buffer = await chartsheetWorkbook([{ worksheet: "Sheet1" }, { chartsheet: "Chart1", xml }]);
            const written = await (await load(buffer)).xlsx.writeBuffer();
            writeSample("chartsheet-settings.xlsx", written);
            const parts = textParts(written);
            // the relationship id of the printer settings (not carried) is left out
            expect(parts["xl/chartsheets/sheet1.xml"]).to.equal(chartsheetXml({ inner: inner.replace(' r:id="rId2"', "") }));
        });

        it("drawings of the chartsheets are numbered after those of the worksheets", async () => {
            const png = fs.readFileSync(path.join(__dirname, "data/image.png"));
            const buffer = await chartsheetWorkbook([{ chartsheet: "Chart1" }, { worksheet: "Sheet1" }]);
            const wb = await load(buffer);
            wb.getWorksheet("Sheet1").addImage(wb.addImage({ buffer: png, extension: "png" }), "D2:E4");
            const parts = textParts(await wb.xlsx.writeBuffer());
            expect(parts["xl/worksheets/_rels/sheet1.xml.rels"]).to.include("../drawings/drawing1.xml");
            expect(parts["xl/chartsheets/_rels/sheet1.xml.rels"]).to.include("../drawings/drawing2.xml");
            expect(parts["xl/drawings/drawing2.xml"]).to.include("<xdr:absoluteAnchor>");
            expect(parts["xl/drawings/_rels/drawing2.xml.rels"]).to.equal(rels([["rId1", REL.chart, "../charts/chart1.xml"]]));
        });
    });

    describe("localSheetId and views count the chartsheets", () => {
        const printAreas =
            (...areas) =>
            (xml) =>
                xml.replace(
                    "</sheets>",
                    `</sheets><definedNames>${areas.map(([id, ref]) => `<definedName name="_xlnm.Print_Area" localSheetId="${id}">${ref}</definedName>`).join("")}</definedNames>`
                );

        it("reads a print area after a chartsheet on its worksheet", async () => {
            // Excel's localSheetId: the index of the <sheet> (Chart1 is 0): it used to give SheetA's area to SheetB
            const buffer = await chartsheetWorkbook([{ chartsheet: "Chart1" }, { worksheet: "SheetA" }, { worksheet: "SheetB" }], {
                workbookXml: printAreas([1, "SheetA!$A$1:$B$2"]),
            });
            const wb = await load(buffer);
            expect(wb.getWorksheet("SheetA").pageSetup.printArea).to.equal("A1:B2");
            expect(wb.getWorksheet("SheetB").pageSetup.printArea).to.equal(undefined);
        });

        it("writes the localSheetId of a print area counting the chartsheets", async () => {
            const buffer = await chartsheetWorkbook([{ worksheet: "Sheet1" }, { chartsheet: "Chart1" }, { worksheet: "Sheet2" }], {
                workbookXml: printAreas([0, "Sheet1!$A$1:$B$2"], [2, "Sheet2!$C$3:$D$4"]),
            });
            const wb = await load(buffer);
            expect(wb.getWorksheet("Sheet2").pageSetup.printArea).to.equal("C3:D4");
            wb.getWorksheet("Sheet2").pageSetup.printTitlesRow = "1:2";
            const xml = textParts(await wb.xlsx.writeBuffer())["xl/workbook.xml"];
            expect(xml).to.include(`<definedName name="_xlnm.Print_Area" localSheetId="0">&apos;Sheet1&apos;!$A$1:$B$2</definedName>`);
            expect(xml).to.include(`<definedName name="_xlnm.Print_Area" localSheetId="2">&apos;Sheet2&apos;!$C$3:$D$4</definedName>`);
            expect(xml).to.include(`<definedName name="_xlnm.Print_Titles" localSheetId="2">&apos;Sheet2&apos;!$1:$2</definedName>`);

            const again = await load(await wb.xlsx.writeBuffer());
            expect(again.getWorksheet("Sheet2").pageSetup.printArea).to.equal("C3:D4");
            expect(again.getWorksheet("Sheet1").pageSetup.printArea).to.equal("A1:B2");
        });

        it("drops a print area of a sheet that is not read", async () => {
            const buffer = await chartsheetWorkbook([{ chartsheet: "Chart1" }, { worksheet: "Sheet1" }], {
                workbookXml: printAreas([0, "Chart1!$A$1:$B$2"], [5, "Sheet1!$A$1:$B$2"]),
            });
            const wb = await load(buffer);
            expect(wb.getWorksheet("Sheet1").pageSetup.printArea).to.equal(undefined);
        });

        it("moves activeTab / firstSheet when a sheet is dropped", async () => {
            const buffer = await chartsheetWorkbook(
                [{ chartsheet: "Broken" }, { worksheet: "Sheet1" }, { chartsheet: "Chart1" }, { worksheet: "Sheet2" }],
                {
                    workbookXml: bookView('firstSheet="1" activeTab="3"'),
                    edit(parts) {
                        delete parts["xl/chartsheets/sheet1.xml"];
                    },
                }
            );
            const wb = await load(buffer);
            expect(wb.chartsheets.map((chartsheet) => chartsheet.name)).to.deep.equal(["Chart1"]);
            expect(wb.views[0]).to.include({ firstSheet: 0, activeTab: 2 });
            const parts = textParts(await wb.xlsx.writeBuffer());
            expect(writtenSheets(parts).map(([name]) => name)).to.deep.equal(["Sheet1", "Chart1", "Sheet2"]);
            expect(parts["xl/workbook.xml"]).to.include('firstSheet="0" activeTab="2"');
        });

        it("an active tab on a dropped sheet moves to the next sheet kept (the last one at the end)", async () => {
            const view = (activeTab) => bookView(`activeTab="${activeTab}"`);
            const sheets = [{ worksheet: "Sheet1" }, { chartsheet: "Broken" }, { worksheet: "Sheet2" }, { chartsheet: "Broken2" }];
            const edit = (parts) => {
                delete parts["xl/chartsheets/sheet1.xml"];
                delete parts["xl/chartsheets/sheet2.xml"];
            };
            expect((await load(await chartsheetWorkbook(sheets, { workbookXml: view(1), edit }))).views[0].activeTab).to.equal(1);
            expect((await load(await chartsheetWorkbook(sheets, { workbookXml: view(3), edit }))).views[0].activeTab).to.equal(1);
            // beyond the sheets: as read
            expect((await load(await chartsheetWorkbook(sheets, { workbookXml: view(7), edit }))).views[0].activeTab).to.equal(7);
        });
    });

    describe("hostile and malformed files: the chartsheet is dropped, the file loads", () => {
        async function dropped(sheet, options) {
            const wb = await load(await chartsheetWorkbook([{ worksheet: "Sheet1" }, { chartsheet: "Chart1", ...sheet }], options));
            expect(wb.chartsheets).to.deep.equal([]);
            expect(wb.getWorksheet("Sheet1").getCell("B3").value).to.equal(3);
            const parts = textParts(await wb.xlsx.writeBuffer());
            expect(Object.keys(parts).filter((name) => name.includes("chart"))).to.deep.equal([]);
            expect(writtenSheets(parts).map(([name]) => name)).to.deep.equal(["Sheet1"]);
            return parts;
        }

        it("a relationship to a missing part", async () => {
            await dropped({}, { edit: (parts) => delete parts["xl/chartsheets/sheet1.xml"] });
            await dropped({ target: "chartsheets/missing.xml" });
            await dropped({}, { edit: (parts) => delete parts["xl/chartsheets/_rels/sheet1.xml.rels"] });
            await dropped({}, { edit: (parts) => delete parts["xl/drawings/drawing1.xml"] });
            await dropped({}, { edit: (parts) => delete parts["xl/charts/chart1.xml"] });
        });

        it("targets out of the zip, with a scheme, or to another kind of part", async () => {
            await dropped({ target: "../../chartsheets/sheet1.xml" });
            await dropped({ target: "http://example.com/chartsheets/sheet1.xml" });
            await dropped({ target: "chartsheets/sheet1.xml#x" });
            await dropped({ rels: rels([["rId1", REL.drawing, "../../../drawings/drawing1.xml"]]) });
            await dropped({ rels: rels([["rId1", REL.drawing, "../drawings/../../drawing1.xml"]]) });
            await dropped({ rels: rels([["rId1", REL.image, "../drawings/drawing1.xml"]]) });
            await dropped({ rels: rels([["rId1", REL.drawing, "../worksheets/sheet1.xml"]]) });
            await dropped({ rels: rels([["rId1", REL.drawing, "../drawings/_rels/drawing1.xml.rels"]]) });
            await dropped({
                rels: `${XML_DECL}<Relationships xmlns="${NS.pkg}"><Relationship Id="rId1" Type="${REL.drawing}" Target="https://example.com/d.xml" TargetMode="External"/></Relationships>`,
            });
        });

        it("an absolute target is resolved", async () => {
            const wb = await load(
                await chartsheetWorkbook([
                    { worksheet: "Sheet1" },
                    {
                        chartsheet: "Chart1",
                        target: "/xl/chartsheets/sheet1.xml",
                        rels: rels([["rId1", REL.drawing, "/xl/drawings/drawing1.xml"]]),
                    },
                ])
            );
            expect(wb.chartsheets.map((chartsheet) => chartsheet.name)).to.deep.equal(["Chart1"]);
        });

        it("cycles: a drawing relationship to the chartsheet itself, a sheet or a drawing used twice", async () => {
            await dropped({ rels: rels([["rId1", REL.drawing, "sheet1.xml"]]) });
            // two <sheet> elements to one chartsheet part, two chartsheets to one drawing: the first one only
            const buffer = await chartsheetWorkbook(
                [{ worksheet: "Sheet1" }, { chartsheet: "Chart1" }, { chartsheet: "Chart2" }, { chartsheet: "Chart3" }],
                {
                    edit(parts) {
                        parts["xl/_rels/workbook.xml.rels"] = parts["xl/_rels/workbook.xml.rels"].replace(
                            "chartsheets/sheet2.xml",
                            "chartsheets/sheet1.xml"
                        );
                        parts["xl/chartsheets/_rels/sheet3.xml.rels"] = rels([["rId1", REL.drawing, "../drawings/drawing1.xml"]]);
                    },
                }
            );
            const wb = await load(buffer);
            expect(wb.chartsheets.map((chartsheet) => chartsheet.name)).to.deep.equal(["Chart1"]);
        });

        it("a drawing of a worksheet stays with the worksheet", async () => {
            const buffer = await chartsheetWorkbook([{ worksheet: "Sheet1" }, { chartsheet: "Chart1" }], {
                edit(parts) {
                    parts["xl/worksheets/sheet1.xml"] = parts["xl/worksheets/sheet1.xml"].replace(
                        "</worksheet>",
                        '<drawing r:id="rId1"/></worksheet>'
                    );
                    parts["xl/worksheets/_rels/sheet1.xml.rels"] = rels([["rId1", REL.drawing, "../drawings/drawing1.xml"]]);
                },
            });
            const wb = await load(buffer);
            expect(wb.chartsheets).to.deep.equal([]);
            const parts = textParts(await wb.xlsx.writeBuffer());
            expect(parts["xl/worksheets/_rels/sheet1.xml.rels"]).to.include("../drawings/drawing1.xml");
            expect(parts["xl/drawings/drawing1.xml"]).to.include("<xdr:absoluteAnchor>");
        });

        it("a drawing without chart, or a chart refused", async () => {
            await dropped({
                drawing: `${XML_DECL}<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>`,
            });
            await dropped(
                {},
                { edit: (parts) => (parts["xl/drawings/_rels/drawing1.xml.rels"] = rels([["rId1", REL.chart, "../../../chart1.xml"]])) }
            );
        });

        it("names already used by a sheet (case-insensitive)", async () => {
            await dropped({ chartsheet: "SHEET1" });
            const wb = await load(await chartsheetWorkbook([{ worksheet: "Sheet1" }, { chartsheet: "Chart" }, { chartsheet: "chart" }]));
            expect(wb.chartsheets.map((chartsheet) => chartsheet.name)).to.deep.equal(["Chart"]);
        });

        it("malformed or refused chartsheet XML", async () => {
            await dropped({ xml: `${XML_DECL}<chartsheet xmlns="${NS.main}"><sheetViews>` });
            await dropped({}, { edit: (parts) => (parts["xl/chartsheets/_rels/sheet1.xml.rels"] = "<Relationships><") });
            // not a chartsheet of SpreadsheetML
            await dropped({ xml: chartsheetXml().replace(/<(\/?)chartsheet/g, "<$1worksheet") });
            await dropped({ xml: chartsheetXml().replace(`xmlns="${NS.main}"`, 'xmlns="urn:other"') });
            // no drawing, two drawings, a drawing without relationship id
            await dropped({ xml: chartsheetXml({ drawing: "" }) });
            await dropped({ xml: chartsheetXml({ drawing: '<drawing r:id="rId1"/><drawing r:id="rId1"/>' }) });
            await dropped({ xml: chartsheetXml({ drawing: "<drawing/>" }) });
            // an undeclared prefix
            await dropped({ xml: chartsheetXml({ after: '<x:extLst><x:ext uri="u"/></x:extLst>' }) });
            await dropped({ xml: chartsheetXml({ inner: '<sheetViews><sheetView q:a="1" workbookViewId="0"/></sheetViews>' }) });
        });

        it("leaves out the elements that reference parts not carried", async () => {
            const after =
                '<legacyDrawing r:id="rId2"/><legacyDrawingHF r:id="rId3"/><drawingHF r:id="rId4"/><picture r:id="rId5"/>' +
                '<webPublishItems count="1"><webPublishItem id="1" divId="d" sourceType="chart" destinationFile="f.htm"/></webPublishItems>' +
                `<extLst><ext uri="{X}" xmlns:p="urn:p"><p:thing xmlns:rr="${NS.r}" rr:id="rId6"/><p:kept a="1"/></ext></extLst>`;
            const xml = chartsheetXml({ after });
            const wb = await load(await chartsheetWorkbook([{ worksheet: "Sheet1" }, { chartsheet: "Chart1", xml }]));
            const parts = textParts(await wb.xlsx.writeBuffer());
            expect(parts["xl/chartsheets/sheet1.xml"]).to.equal(
                chartsheetXml({
                    after:
                        '<webPublishItems count="1"><webPublishItem id="1" divId="d" sourceType="chart" destinationFile="f.htm"/></webPublishItems>' +
                        '<extLst><ext uri="{X}" xmlns:p="urn:p"><p:kept a="1"/></ext></extLst>',
                })
            );
            expect(parts["xl/chartsheets/_rels/sheet1.xml.rels"]).to.equal(rels([["rId1", REL.drawing, "../drawings/drawing1.xml"]]));
        });

        it("a drawing relationship id of any prefix, renumbered", async () => {
            const xml = `${XML_DECL}<x:chartsheet xmlns:x="${NS.main}" xmlns:rel="${NS.r}"><x:sheetViews><x:sheetView workbookViewId="0"/></x:sheetViews><x:drawing rel:id="d&amp;1"/></x:chartsheet>`;
            const wb = await load(
                await chartsheetWorkbook([
                    { worksheet: "Sheet1" },
                    { chartsheet: "Chart1", xml, rels: rels([["d&amp;1", REL.drawing, "../drawings/drawing1.xml"]]) },
                ])
            );
            const parts = textParts(await wb.xlsx.writeBuffer());
            expect(parts["xl/chartsheets/sheet1.xml"]).to.equal(xml.replace("d&amp;1", "rId1"));
        });

        it("many chartsheets: bounded, and read in linear time", async function () {
            this.timeout(60000);
            const count = 1500;
            const sheets = [{ worksheet: "Sheet1" }];
            for (let i = 1; i <= count; i++) {
                sheets.push({ chartsheet: `Chart${i}` });
            }
            const buffer = await chartsheetWorkbook(sheets);
            const start = Date.now();
            const wb = await load(buffer);
            const parts = textParts(await wb.xlsx.writeBuffer());
            const elapsed = Date.now() - start;
            expect(wb.chartsheets).to.have.length(count);
            expect(parts[`xl/chartsheets/sheet${count}.xml`]).to.equal(chartsheetXml());
            expect(elapsed).to.be.below(20000);

            // the bound on the chartsheet parts read
            const entries = new Map(Object.entries(unzipParts(buffer)).map(([name, data]) => [name, { data: async () => data }]));
            const wbParts = textParts(buffer);
            const sheetDefs = [...wbParts["xl/workbook.xml"].matchAll(/<sheet name="([^"]*)" sheetId="(\d+)" r:id="([^"]+)"\/>/g)].map(
                ([, name, id, rId]) => ({
                    name,
                    id: Number(id),
                    rId,
                })
            );
            const workbookRels = [
                ...wbParts["xl/_rels/workbook.xml.rels"].matchAll(/<Relationship Id="([^"]+)" Type="([^"]+)" Target="([^"]+)"\/>/g),
            ].map(([, Id, Type, Target]) => ({ Id, Type, Target }));
            expect((await loadChartsheets(sheetDefs, workbookRels, entries)).size).to.equal(count);
            expect((await loadChartsheets(sheetDefs, workbookRels, entries, 10)).size).to.equal(10);
            expect(MAX_CHARTSHEETS).to.equal(10000);
        });

        it("namespace declarations × elements: linear", async () => {
            const n = 20000;
            let inner = "<sheetViews>";
            for (let i = 0; i < n; i++) {
                inner += `<sheetView xmlns:p${i % 50}="urn:${i}" workbookViewId="0"/>`;
            }
            inner += "</sheetViews>";
            let declarations = "";
            for (let i = 0; i < n; i++) {
                declarations += ` xmlns:d${i}="urn:d${i}"`;
            }
            const xml = `${XML_DECL}<chartsheet xmlns="${NS.main}" xmlns:r="${NS.r}"${declarations}>${inner}<drawing r:id="rId1"/></chartsheet>`;
            const start = Date.now();
            const model = await new ChartsheetXform().parseStream([xml]);
            expect(Date.now() - start).to.be.below(2000);
            expect(model.drawingRelId).to.equal("rId1");
            expect(model.sheet.children[0].children).to.have.length(n);
        });
    });

    describe("review fixes", () => {
        const pie = { type: "pie", series: [{ values: "Data!B1:B3" }] };
        const bookViews = (view) => [{ x: 0, y: 0, width: 100, height: 100, visibility: "visible", ...view }];
        const writtenView = async (wb) => textParts(await wb.xlsx.writeBuffer())["xl/workbook.xml"].match(/<workbookView [^>]*\/>/)[0];

        it("replacing the workbook by a model or a file forgets the former chartsheets first (their names do not clash)", async () => {
            const source = new ExcelJS.Workbook();
            source.addWorksheet("Sheet1").getCell("A1").value = 1;
            const buffer = await source.xlsx.writeBuffer();

            const created = () => {
                const wb = new ExcelJS.Workbook();
                wb.addWorksheet("Data").getCell("B1").value = 1;
                wb.addChartsheet("Sheet1", pie);
                return wb;
            };
            const loaded = created();
            await loaded.xlsx.load(buffer);
            expect(loaded.worksheets.map((ws) => ws.name)).to.deep.equal(["Sheet1"]);
            expect(loaded.chartsheets).to.deep.equal([]);

            const replaced = created();
            replaced.model = source.model;
            expect(replaced.worksheets.map((ws) => ws.name)).to.deep.equal(["Sheet1"]);
            expect(replaced.chartsheets).to.deep.equal([]);
        });

        it("removeChartsheet takes the exact name, as getWorksheet / removeWorksheet", async () => {
            const wb = await load(fs.readFileSync(EXCEL_FILE));
            wb.removeChartsheet("chart1");
            expect(wb.chartsheets.map((chartsheet) => chartsheet.name)).to.deep.equal(["Chart1"]);
            expect(wb.getWorksheet("sheet1")).to.equal(undefined);
        });

        it("a chartsheet tab is selected only when it is the active tab", async () => {
            const tabSelected = async (wb) => textParts(await wb.xlsx.writeBuffer())["xl/chartsheets/sheet1.xml"].includes("tabSelected");
            // Chart1 (tab 0) is selected in chart-sheet.xlsx, which has no activeTab (0)
            const wb = await load(fs.readFileSync(EXCEL_FILE));
            expect(await tabSelected(wb)).to.equal(true);
            // the worksheet made active: the chartsheet is no longer selected (else Excel opens both tabs grouped)
            wb.views = bookViews({ activeTab: 1 });
            expect(await tabSelected(wb)).to.equal(false);
            const written = textParts(await wb.xlsx.writeBuffer())["xl/chartsheets/sheet1.xml"];
            expect(written).to.include('<sheetView zoomScale="121" workbookViewId="0" zoomToFit="1"/>');
            wb.views = bookViews({ activeTab: 0 });
            expect(await tabSelected(wb)).to.equal(true);

            // a chartsheet created, made active
            const created = new ExcelJS.Workbook();
            created.addWorksheet("Data").getCell("B1").value = 1;
            created.addChartsheet("Pie", pie);
            expect(await tabSelected(created)).to.equal(false);
            created.views = bookViews({ activeTab: 1 });
            expect(await tabSelected(created)).to.equal(true);
        });

        it("removing a sheet moves activeTab / firstSheet to the same sheets, or the next one", async () => {
            // Chart1, Sheet1, Added, Last
            const wb = await load(fs.readFileSync(EXCEL_FILE));
            wb.addWorksheet("Added");
            wb.addWorksheet("Last");
            wb.views = bookViews({ activeTab: 3, firstSheet: 1 });
            wb.removeChartsheet("Chart1");
            expect(wb.views[0]).to.include({ activeTab: 2, firstSheet: 0 });
            expect(await writtenView(wb)).to.include('firstSheet="0" activeTab="2"');
            // the active (last) sheet removed: the last one left
            wb.removeWorksheet("Last");
            expect(wb.views[0]).to.include({ activeTab: 1, firstSheet: 0 });
            // a sheet after the active one: unchanged
            wb.views = bookViews({ activeTab: 0 });
            wb.removeWorksheet("Added");
            expect(wb.views[0]).to.include({ activeTab: 0 });
            // a view without activeTab keeps none
            wb.views = bookViews({});
            wb.removeWorksheet("Sheet1");
            expect(wb.views[0]).to.not.have.property("activeTab");
        });
    });

    describe("mergeSheets", () => {
        it("inserts the chartsheets at their indexes", () => {
            const c = (index, name = `c${index}`) => ({ index, name });
            expect(mergeSheets(["a", "b"], undefined)).to.deep.equal(["a", "b"]);
            expect(mergeSheets(["a", "b"], [c(2), c(0)]).map((s) => s.name ?? s)).to.deep.equal(["c0", "a", "c2", "b"]);
            expect(mergeSheets(["a"], [c(5), c(NaN, "x"), c(1)]).map((s) => s.name ?? s)).to.deep.equal(["a", "c1", "c5", "x"]);
            expect(mergeSheets([], [c(3), c(1)]).map((s) => s.name)).to.deep.equal(["c1", "c3"]);
        });
    });

    it("a workbook without chartsheets writes no chartsheet part (as before)", async () => {
        const wb = await load(await baseWorkbookBuffer());
        expect(wb.chartsheets).to.deep.equal([]);
        expect(wb.model).to.not.have.property("chartsheets");
        const parts = textParts(await wb.xlsx.writeBuffer());
        expect(parts["docProps/app.xml"]).to.include('<vt:vector size="2" baseType="variant">');
        expect(parts["[Content_Types].xml"]).to.not.include("chartsheet");
    });
});
