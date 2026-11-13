// Generates workbooks for a manual check in Microsoft Excel: bun scripts/excel-check/generate.js
//
// Every file is written with the public API of the package, under spec/out/excel-check/ (ignored by git)
// with a README.md that says what to look at in each file. A file must open without a repair prompt.
const fs = require("fs");
const path = require("path");

const ExcelJS = require("../../excel");

const OUT_DIR = path.resolve(__dirname, "../../spec/out/excel-check");
const IMAGE = path.resolve(__dirname, "../../spec/integration/data/image.png");

// the only place this script writes to
function outFile(name) {
    const file = path.resolve(OUT_DIR, name);
    if (path.dirname(file) !== OUT_DIR) {
        throw new Error(`refusing to write outside ${OUT_DIR}: ${name}`);
    }
    return file;
}

const bold = { font: { bold: true } };
const yellow = { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFFFF00" } } };
const red = { font: { color: { argb: "FFC00000" }, bold: true } };

async function reload(wb) {
    const back = new ExcelJS.Workbook();
    await back.xlsx.load(await wb.xlsx.writeBuffer());
    return back;
}

const files = [];

function file(name, checks, generate) {
    files.push({ name, checks, generate });
}

// ---------------------------------------------------------------------------------------------------
file(
    "x14-cross-sheet-cf.xlsx",
    [
        "Sheet `Data`, A1:A10 hold 1..10. Sheet `Rules` holds the thresholds B1=7, B2=2, B3=4.",
        "Data!A1:A10: values >= Rules!B1 (7..10) have a yellow fill (expression rule referencing 'Rules'!$B$1).",
        "Data!A1:A10: values between Rules!B2 and Rules!B3 (2..4) are bold red (cellIs rule referencing Rules!$B$2 and Rules!$B$3).",
        "Data!B1:B10 (=A*10): values > 50 are bold (same-sheet rule, written in the main section).",
        "Home > Conditional Formatting > Manage Rules (show formatting rules for: This Worksheet) lists the 3 rules, the cross-sheet ones with their formulas. Change Rules!B1 to 9: the yellow fill follows.",
    ],
    async () => {
        const wb = new ExcelJS.Workbook();
        const data = wb.addWorksheet("Data");
        const rules = wb.addWorksheet("Rules");
        rules.getCell("A1").value = "yellow from";
        rules.getCell("B1").value = 7;
        rules.getCell("A2").value = "red between";
        rules.getCell("B2").value = 2;
        rules.getCell("B3").value = 4;
        for (let i = 1; i <= 10; i++) {
            data.getCell(`A${i}`).value = i;
            data.getCell(`B${i}`).value = { formula: `A${i}*10`, result: i * 10 };
        }
        data.addConditionalFormatting({
            ref: "A1:A10",
            rules: [
                { type: "expression", formulae: ["A1>='Rules'!$B$1"], style: yellow },
                { type: "cellIs", operator: "between", formulae: ["Rules!$B$2", "Rules!$B$3"], style: red },
            ],
        });
        data.addConditionalFormatting({
            ref: "B1:B10",
            rules: [{ type: "cellIs", operator: "greaterThan", formulae: ["50"], style: bold }],
        });
        await wb.xlsx.writeFile(outFile("x14-cross-sheet-cf.xlsx"));
        // the same workbook read back and written again (rules read from the x14 extension)
        await (await reload(wb)).xlsx.writeFile(outFile("x14-cross-sheet-cf-reread.xlsx"));
    }
);

// ---------------------------------------------------------------------------------------------------
file(
    "internal-hyperlinks.xlsx",
    [
        "Sheet `Links`: A1 jumps to cell B2 of sheet `Other Sheet`; A2 jumps to the defined name `Target` (Other Sheet!D4:E5, selected after the click).",
        "A3 opens https://example.com/page#section in the browser (the #section anchor is kept: check with Edit Hyperlink).",
        "A4 jumps to Links!C10 (same sheet).",
    ],
    async () => {
        const wb = new ExcelJS.Workbook();
        const links = wb.addWorksheet("Links");
        const other = wb.addWorksheet("Other Sheet");
        other.getCell("B2").value = "B2 of Other Sheet";
        other.getCell("D4").value = "Target";
        wb.definedNames.add("'Other Sheet'!$D$4:$E$5", "Target");
        links.getCell("A1").value = { text: "to 'Other Sheet'!B2", hyperlink: "#'Other Sheet'!B2" };
        links.getCell("A2").value = { text: "to the name Target", hyperlink: "#Target" };
        links.getCell("A3").value = { text: "to example.com#section", hyperlink: "https://example.com/page#section" };
        links.getCell("A4").value = { text: "to C10", hyperlink: "#Links!C10" };
        links.getCell("C10").value = "C10";
        links.getColumn(1).width = 28;
        await wb.xlsx.writeFile(outFile("internal-hyperlinks.xlsx"));
    }
);

// ---------------------------------------------------------------------------------------------------
file(
    "print-areas.xlsx",
    [
        "Sheet `Bob's sheet` (apostrophe in the name): Page Layout > Print Area shows 2 areas, A1:C10 and E1:G10 (File > Print previews them on separate pages).",
        "Page Layout > Print Titles: rows to repeat $1:$2, columns to repeat $A:$A.",
        "Formulas > Name Manager lists Print_Area and Print_Titles scoped to `Bob's sheet`. ExcelJS writes the areas as $A1:$C10 (rows not absolute): check they still cover A1:C10 and E1:G10.",
    ],
    async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet("Bob's sheet");
        for (let r = 1; r <= 40; r++) {
            for (let c = 1; c <= 8; c++) {
                ws.getCell(r, c).value = r <= 2 ? `title ${r}/${c}` : r * c;
            }
        }
        ws.pageSetup.printArea = "A1:C10&&E1:G10";
        ws.pageSetup.printTitlesRow = "1:2";
        ws.pageSetup.printTitlesColumn = "A:A";
        await wb.xlsx.writeFile(outFile("print-areas.xlsx"));
    }
);

// ---------------------------------------------------------------------------------------------------
function addNotes(wb) {
    ["One", "Two", "Three"].forEach((name, index) => {
        const ws = wb.addWorksheet(name);
        ws.getCell("A1").value = `sheet ${name}`;
        ws.getCell("B2").value = "note here";
        ws.getCell("B2").note = `Note on ${name}!B2`;
        ws.getCell("D5").value = "rich note";
        ws.getCell("D5").note = {
            texts: [{ font: { bold: true }, text: `${name}: ` }, { text: `second note (${index + 1})` }],
        };
        if (ws.commit) {
            ws.getRow(1).commit();
            ws.getRow(2).commit();
            ws.getRow(5).commit();
            ws.commit();
        }
    });
}

file(
    "comments-multi-sheet.xlsx",
    [
        "Sheets One, Two, Three: B2 and D5 have a red note marker; hovering shows `Note on <sheet>!B2` and `<sheet>: second note (n)` (the prefix bold).",
        "Each note belongs to its own sheet (no note shown on the wrong sheet). Review > Notes > Show All Notes shows 2 notes per sheet.",
        "Same checks for comments-multi-sheet-stream.xlsx (written with the streaming WorkbookWriter).",
    ],
    async () => {
        const wb = new ExcelJS.Workbook();
        addNotes(wb);
        await wb.xlsx.writeFile(outFile("comments-multi-sheet.xlsx"));

        const streamed = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: outFile("comments-multi-sheet-stream.xlsx"), useStyles: true });
        addNotes(streamed);
        await streamed.commit();
    }
);

// ---------------------------------------------------------------------------------------------------
file(
    "tables-filter-buttons.xlsx",
    [
        "Sheet `Tables` (the file was written, read back and written again): table `Fruits` at B2:D6 with a header row and a totals row.",
        "The headers Name and Qty have filter drop-down buttons, Price has none (hidden button). Filtering on Name works.",
        "Table Design > Table Style is TableStyleMedium9 with banded rows.",
    ],
    async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet("Tables");
        ws.addTable({
            name: "Fruits",
            ref: "B2",
            headerRow: true,
            totalsRow: true,
            style: { theme: "TableStyleMedium9", showRowStripes: true },
            columns: [
                { name: "Name", filterButton: true, totalsRowLabel: "Total" },
                { name: "Qty", filterButton: true, totalsRowFunction: "sum" },
                { name: "Price", filterButton: false, totalsRowFunction: "average" },
            ],
            rows: [
                ["Apple", 3, 1.2],
                ["Banana", 5, 0.5],
                ["Cherry", 12, 4],
            ],
        });
        await (await reload(wb)).xlsx.writeFile(outFile("tables-filter-buttons.xlsx"));
    }
);

// ---------------------------------------------------------------------------------------------------
file(
    "images-repeat.xlsx",
    [
        "Sheet `First`: the same image at B2:C4 and at E2:F6 (stretched). Sheet `Second`: the same image at B2:C4.",
        "All 3 pictures display (no red cross / `picture can't be displayed`). The file holds the image once (xl/media/ has one file).",
    ],
    async () => {
        const wb = new ExcelJS.Workbook();
        const imageId = wb.addImage({ filename: IMAGE, extension: "png" });
        const first = wb.addWorksheet("First");
        const second = wb.addWorksheet("Second");
        first.addImage(imageId, "B2:C4");
        first.addImage(imageId, "E2:F6");
        second.addImage(imageId, "B2:C4");
        await wb.xlsx.writeFile(outFile("images-repeat.xlsx"));
    }
);

// ---------------------------------------------------------------------------------------------------
file(
    "data-validations-whole-column.xlsx",
    [
        "Sheets `S1` and `S2`: any cell of column A (try A1, A500, A100000) offers a drop-down list Yes/No/Maybe; typing another value is refused.",
        "On S2, A5 instead accepts only a whole number between 1 and 10 (input message `1 to 10`).",
        "Data > Data Validation on column B shows no validation.",
    ],
    async () => {
        const wb = new ExcelJS.Workbook();
        const list = {
            type: "list",
            allowBlank: true,
            formulae: ['"Yes,No,Maybe"'],
            showErrorMessage: true,
            errorTitle: "Invalid",
            error: "Pick Yes, No or Maybe",
        };
        ["S1", "S2"].forEach((name) => {
            const ws = wb.addWorksheet(name);
            ws.getCell("B1").value = `${name}: column A has a list validation`;
            // a whole column is keyed by its full range ("A:A" is not a supported key)
            ws.dataValidations.add("A1:A1048576", { ...list });
        });
        wb.getWorksheet("S2").getCell("A5").dataValidation = {
            type: "whole",
            operator: "between",
            formulae: [1, 10],
            showInputMessage: true,
            promptTitle: "Number",
            prompt: "1 to 10",
            showErrorMessage: true,
        };
        await wb.xlsx.writeFile(outFile("data-validations-whole-column.xlsx"));
    }
);

// ---------------------------------------------------------------------------------------------------
file(
    "merged-and-styled.xlsx",
    [
        "Sheet `Merged`: 2 rows were inserted above the merge, so the merged block is B4:D5 (one cell, text `merged`, centred, bold white on blue, thick border around).",
        "Rows 2 and 3 hold the inserted values `inserted 1` / `inserted 2` in A2:A3; A6 holds `below` (was A4).",
        "Selecting B4 selects the whole merge; no stray merge remains at B2:D3.",
    ],
    async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet("Merged");
        ws.getCell("A1").value = "header";
        ws.mergeCells("B2:D3");
        const master = ws.getCell("B2");
        master.value = "merged";
        master.alignment = { horizontal: "center", vertical: "middle" };
        master.font = { bold: true, color: { argb: "FFFFFFFF" } };
        master.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E79" } };
        master.border = {
            top: { style: "thick" },
            left: { style: "thick" },
            bottom: { style: "thick" },
            right: { style: "thick" },
        };
        ws.getCell("A4").value = "below";
        ws.insertRows(2, [["inserted 1"], ["inserted 2"]]);
        await wb.xlsx.writeFile(outFile("merged-and-styled.xlsx"));
    }
);

// ---------------------------------------------------------------------------------------------------
file(
    "stream-writer-order.xlsx",
    [
        "Written with the streaming WorkbookWriter; sheet `Report` holds every part whose order the schema fixes:",
        "Sheet protected (Review > Unprotect Sheet, no password); autofilter buttons on A1:C1; B10:C11 merged (`merged`).",
        "Conditional formatting: A2:A20 values > 10 yellow; C2:C20 cells equal to Lookup!A1 (=5) bold red (cross-sheet, x14).",
        "D2 has a list validation (a/b/c); D3 links to https://example.com/ and D4 to Lookup!A1.",
        "File > Print: landscape, A4, fit to 1 page wide; header `Report` centred, footer `Page n of m`.",
        "A background image tiles the sheet (Page Layout > Delete Background is enabled).",
    ],
    async () => {
        const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: outFile("stream-writer-order.xlsx"), useStyles: true });
        const imageId = wb.addImage({ filename: IMAGE, extension: "png" });
        const ws = wb.addWorksheet("Report", {
            pageSetup: { orientation: "landscape", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
            headerFooter: { oddHeader: "&CReport", oddFooter: "&CPage &P of &N" },
        });
        ws.columns = [
            { header: "Value", key: "value", width: 10 },
            { header: "Name", key: "name", width: 12 },
            { header: "Code", key: "code", width: 10 },
            { header: "Extra", key: "extra", width: 30 },
        ];
        await ws.protect("", { selectLockedCells: true, selectUnlockedCells: true, autoFilter: true });
        ws.autoFilter = "A1:C1";
        ws.addConditionalFormatting({
            ref: "A2:A20",
            rules: [{ type: "cellIs", operator: "greaterThan", formulae: ["10"], style: yellow }],
        });
        ws.addConditionalFormatting({
            ref: "C2:C20",
            rules: [{ type: "cellIs", operator: "equal", formulae: ["Lookup!$A$1"], style: red }],
        });
        ws.dataValidations.add("D2", { type: "list", allowBlank: true, formulae: ['"a,b,c"'] });
        ws.addBackgroundImage(imageId);
        for (let i = 1; i <= 19; i++) {
            const row = ws.addRow({ value: i * 2, name: `row ${i}`, code: i % 7 });
            if (i === 2) {
                row.getCell("extra").value = { text: "example.com", hyperlink: "https://example.com/" };
            }
            if (i === 3) {
                row.getCell("extra").value = { text: "to Lookup!A1", hyperlink: "#Lookup!A1" };
            }
        }
        ws.mergeCells("B10:C11");
        ws.getCell("B10").value = "merged";
        ws.commit();

        const lookup = wb.addWorksheet("Lookup");
        lookup.getCell("A1").value = 5;
        lookup.getRow(1).commit();
        lookup.commit();
        await wb.commit();
    }
);

// ---------------------------------------------------------------------------------------------------
// tables, page breaks, headers and footers, note authors

file(
    "tables-styles-totals.xlsx",
    [
        "Sheet `Sales`: table `Sales` (A1:C6) with a totals row. In the data rows, the `Amount` column is bold blue with 2 decimals; its header and its total are NOT blue.",
        "Total row: A6 shows `Total`; B6 counts the products (5); C6 holds the custom formula `SUM(Sales[Amount])*2` = 2 × the sum of C2:C5.",
        "Click C6: the formula bar shows `=SUM(Sales[Amount])*2`. Table Design > Total Row: unticking and re-ticking it keeps the custom formula.",
    ],
    async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet("Sales");
        ws.addTable({
            name: "Sales",
            ref: "A1",
            totalsRow: true,
            columns: [
                { name: "Product", totalsRowLabel: "Total" },
                { name: "Count", totalsRowFunction: "count" },
                {
                    name: "Amount",
                    totalsRowFunction: "custom",
                    totalsRowFormula: "SUM(Sales[Amount])*2",
                    style: { font: { bold: true, color: { argb: "FF1F4E79" } }, numFmt: "0.00" },
                },
            ],
            rows: [
                ["Apples", 1, 10.5],
                ["Pears", 2, 20.25],
                ["Plums", 3, 5],
                ["Kiwis", 4, 7.75],
            ],
        });
        await wb.xlsx.writeFile(outFile("tables-styles-totals.xlsx"));
    }
);

file(
    "page-breaks-header-footer.xlsx",
    [
        "Sheet `Report`: View > Page Break Preview shows manual breaks after column C and after column F (blue lines), and after row 20.",
        "File > Print preview: header `Confidential` (left, bold) / sheet name `Report` (center) / `Page 1 of N` (right); footer `R&D – printed <date>` in the center (the `&` shown as is, the date filled in).",
    ],
    async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet("Report");
        for (let r = 1; r <= 40; r++) {
            for (let c = 1; c <= 9; c++) {
                ws.getCell(r, c).value = r * 100 + c;
            }
        }
        ws.getColumn(3).addPageBreak();
        ws.getColumn(6).addPageBreak();
        ws.getRow(20).addPageBreak();
        ws.headerFooter.oddHeader = ExcelJS.headerFooter.build({
            left: "&BConfidential&B",
            center: "&[Tab]",
            right: "Page &[Page] of &[Pages]",
        });
        ws.headerFooter.oddFooter = ExcelJS.headerFooter.build({
            center: `${ExcelJS.headerFooter.escape("R&D")} – printed &[Date]`,
        });
        await wb.xlsx.writeFile(outFile("page-breaks-header-footer.xlsx"));
    }
);

file(
    "note-authors.xlsx",
    [
        "Sheet `Notes`: A1, A2, A3 have notes (red triangle). Hover each: A1 by `Alice`, A2 by `Bob`, A3 by `Alice` (Review > Notes > Edit Note shows the author as the first bold line only if Excel adds it; the author is also in the status bar on hover).",
        "File > Info > Inspect Workbook > Document Inspector lists comments/notes by Alice and Bob.",
    ],
    async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet("Notes");
        ws.getCell("A1").value = "first";
        ws.getCell("A1").note = { texts: [{ text: "Alice:", font: { bold: true } }, { text: " check this" }], author: "Alice" };
        ws.getCell("A2").value = "second";
        ws.getCell("A2").note = { texts: [{ text: "Bob:", font: { bold: true } }, { text: " looks fine" }], author: "Bob" };
        ws.getCell("A3").value = "third";
        ws.getCell("A3").note = { texts: [{ text: "Alice:", font: { bold: true } }, { text: " done" }], author: "Alice" };
        await wb.xlsx.writeFile(outFile("note-authors.xlsx"));
    }
);

// ---------------------------------------------------------------------------------------------------
// charts

function salesSheet(wb, name = "Sales") {
    const ws = wb.addWorksheet(name);
    ws.addRow(["Month", "North", "South", "West"]);
    const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"];
    months.forEach((month, i) => ws.addRow([month, 10 + i * 3, 7 + ((i * 5) % 9), 12 - i]));
    return ws;
}

const SALES_SERIES = [
    { name: { ref: "B1" }, categories: "A2:A7", values: "B2:B7" },
    { name: { ref: "C1" }, categories: "A2:A7", values: "C2:C7" },
    { name: { ref: "D1" }, categories: "A2:A7", values: "D2:D7" },
];

file(
    "charts-all-types.xlsx",
    [
        "Sheet `Sales` holds Month/North/South/West for Jan..Jun (A1:D7) and 8 charts to its right, each titled with its type:",
        "`column (clustered)`, `bar (stacked)`, `line (markers)`, `area (percent stacked)`: 3 series North/South/West, months on the category axis, legend at the bottom.",
        "`pie` and `doughnut`: one series (North), 6 slices labelled by month in the legend.",
        "`scatter`: North (x) against South (y), points joined by lines, axis titles `North` and `South`.",
        "`column (colors, axis)`: North in orange, South in green, value axis titled `Units` from 0 to 30.",
        "Change Sales!B2 to 40: every chart using North updates.",
    ],
    async () => {
        const wb = new ExcelJS.Workbook();
        const ws = salesSheet(wb);
        const at = (i) => `${i % 2 ? "N" : "F"}${2 + Math.floor(i / 2) * 18}:${i % 2 ? "U" : "M"}${17 + Math.floor(i / 2) * 18}`;
        const legend = { position: "b" };
        ws.addChart({ type: "column", title: "column (clustered)", anchor: at(0), series: SALES_SERIES, legend });
        ws.addChart({ type: "bar", grouping: "stacked", title: "bar (stacked)", anchor: at(1), series: SALES_SERIES, legend });
        ws.addChart({ type: "line", title: "line (markers)", anchor: at(2), series: SALES_SERIES, legend });
        ws.addChart({
            type: "area",
            grouping: "percentStacked",
            title: "area (percent stacked)",
            anchor: at(3),
            series: SALES_SERIES,
            legend,
        });
        ws.addChart({ type: "pie", title: "pie", anchor: at(4), series: [SALES_SERIES[0]] });
        ws.addChart({ type: "doughnut", title: "doughnut", anchor: at(5), series: [SALES_SERIES[0]] });
        ws.addChart({
            type: "scatter",
            line: true,
            title: "scatter",
            anchor: at(6),
            series: [{ name: "North/South", xValues: "B2:B7", yValues: "C2:C7" }],
            axes: { x: { title: "North" }, y: { title: "South" } },
        });
        ws.addChart({
            type: "column",
            title: "column (colors, axis)",
            anchor: at(7),
            series: [
                { ...SALES_SERIES[0], color: "ED7D31" },
                { ...SALES_SERIES[1], color: "#70AD47" },
            ],
            axes: { value: { title: "Units", min: 0, max: 30 } },
        });
        await wb.xlsx.writeFile(outFile("charts-all-types.xlsx"));
    }
);

file(
    "charts-mixed-sheets.xlsx",
    [
        "Sheet `Dashboard`: an image (top left), a line chart of `'Sales Data'` (another sheet whose name has a space) and a pie chart, both titled.",
        "Sheet `Sales Data` holds the data. Renaming `Sales Data` in Excel updates the series formulas of the charts (Select Data).",
        "Chartsheet `Chart Sales` (a tab that is only a chart, after `Sales Data`): a column chart of North/South/West filling the window.",
        "Chartsheet `Hidden chart` is hidden: right-click a tab > Unhide lists it.",
    ],
    async () => {
        const wb = new ExcelJS.Workbook();
        const dashboard = wb.addWorksheet("Dashboard");
        salesSheet(wb, "Sales Data");
        const imageId = wb.addImage({ filename: IMAGE, extension: "png" });
        dashboard.addImage(imageId, "A1:C8");
        const ranges = (col) => ({
            name: { ref: `'Sales Data'!${col}1` },
            categories: "'Sales Data'!A2:A7",
            values: `'Sales Data'!${col}2:${col}7`,
        });
        dashboard.addChart({ type: "line", title: "Trend", anchor: "E1:L16", series: [ranges("B"), ranges("C")] });
        dashboard.addChart({ type: "pie", title: "North share", anchor: "E18:L33", series: [ranges("B")] });
        wb.addChartsheet("Chart Sales", {
            type: "column",
            title: "All regions",
            series: [ranges("B"), ranges("C"), ranges("D")],
        });
        wb.addChartsheet("Hidden chart", { type: "line", series: [ranges("D")] }, { state: "hidden" });
        await wb.xlsx.writeFile(outFile("charts-mixed-sheets.xlsx"));
    }
);

file(
    "charts-read-back.xlsx",
    [
        "charts-mixed-sheets.xlsx read back by ExcelJS and written again, with a worksheet `Added` (holding `added` in A1) appended: same checks as charts-mixed-sheets.xlsx, plus the `Added` tab at the end.",
    ],
    async () => {
        const back = new ExcelJS.Workbook();
        await back.xlsx.readFile(outFile("charts-mixed-sheets.xlsx"));
        back.addWorksheet("Added").getCell("A1").value = "added";
        await back.xlsx.writeFile(outFile("charts-read-back.xlsx"));
    }
);

file(
    "chartsheet-excel-read-back.xlsx",
    [
        "spec/integration/data/chart-sheet.xlsx (made by Excel: a chartsheet `Chart1` before the worksheet `Sheet1`) read by ExcelJS and written again: the tab `Chart1` shows the same chart as the original, `Sheet1` holds its data.",
    ],
    async () => {
        const back = new ExcelJS.Workbook();
        await back.xlsx.readFile(path.resolve(__dirname, "../../spec/integration/data/chart-sheet.xlsx"));
        await back.xlsx.writeFile(outFile("chartsheet-excel-read-back.xlsx"));
    }
);

file(
    "pivot-read-back.xlsx",
    [
        "A pivot table and a pivot chart (built by hand, as Excel stores them) read by ExcelJS and written again, plus a pivot table added with `addPivotTable`.",
        "Sheet `Pivot`: the pivot table `PivotTable1` (sum of Sales by Region) and its pivot chart (field buttons on the chart). PivotTable Analyze > Refresh works.",
        "Sheet `Added`: the pivot table added by ExcelJS (Region by Product, sum of Sales). Refresh works too.",
    ],
    async () => {
        global.verquire = global.verquire || require("../../spec/utils/verquire");
        const { pivotWorkbook } = require("../../spec/utils/pivot-fixtures");
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(await pivotWorkbook({ chart: true }));
        wb.addWorksheet("Added").addPivotTable({
            sourceSheet: wb.getWorksheet("Data"),
            rows: ["Region"],
            columns: ["Product"],
            values: ["Sales"],
        });
        await wb.xlsx.writeFile(outFile("pivot-read-back.xlsx"));
    }
);

file(
    "charts-streaming.xlsx",
    [
        "Written by the streaming writer. Sheet `Data`: 200 rows (Item 1..200, values 2..400) and a column chart `Cached` to the right of the data (added before the rows: its values are cached).",
        "Chartsheet `Line (no cache)`: a line chart of Data!B2:B201, added after the data was written (no cached values: Excel draws it from the cells; the Finder / Quick Look preview may show it empty).",
    ],
    async () => {
        const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: outFile("charts-streaming.xlsx") });
        const data = wb.addWorksheet("Data");
        data.addRow(["Item", "Value"]).commit();
        data.addChart({
            type: "column",
            title: "Cached",
            anchor: "D2:L18",
            series: [{ name: { ref: "B1" }, categories: "A2:A201", values: "B2:B201" }],
        });
        for (let i = 1; i <= 200; i++) {
            data.addRow([`Item ${i}`, i * 2]).commit();
        }
        await data.commit();
        wb.addChartsheet("Line (no cache)", { type: "line", title: "Line (no cache)", series: [{ values: "Data!B2:B201" }] });
        await wb.commit();
    }
);

// ---------------------------------------------------------------------------------------------------
function readme() {
    const lines = [
        "# Manual Excel check",
        "",
        "Generated by `bun scripts/excel-check/generate.js` (do not commit). Open every file in Microsoft Excel.",
        "",
        "For **every** file: it opens with **no repair prompt** (`We found a problem with some content...`) and no",
        "`PROTECTED VIEW` error beyond the usual downloaded-file banner. Then check what is listed for it.",
        "",
    ];
    files.forEach(({ name, checks }) => {
        lines.push(`## ${name}`, "", ...checks.map((check) => `- ${check}`), "");
    });
    lines.push(
        "## x14-cross-sheet-cf-reread.xlsx",
        "",
        "- x14-cross-sheet-cf.xlsx read back by ExcelJS and written again: same checks as x14-cross-sheet-cf.xlsx.",
        ""
    );
    return lines.join("\n");
}

async function main() {
    fs.mkdirSync(OUT_DIR, { recursive: true });
    for (const { generate } of files) {
        await generate();
    }
    fs.writeFileSync(outFile("README.md"), readme());
    process.stdout.write(`${files.length} checks written to ${OUT_DIR}\n`);
}

main().catch((error) => {
    process.stderr.write(`${error.stack}\n`);
    process.exitCode = 1;
});
