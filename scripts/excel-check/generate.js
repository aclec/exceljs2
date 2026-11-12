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
