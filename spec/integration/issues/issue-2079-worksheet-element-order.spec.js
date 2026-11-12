const { readPart } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

// ECMA-376 Part 1, CT_Worksheet: children must appear in this order.
const CT_WORKSHEET = [
    "sheetPr",
    "dimension",
    "sheetViews",
    "sheetFormatPr",
    "cols",
    "sheetData",
    "sheetCalcPr",
    "sheetProtection",
    "protectedRanges",
    "scenarios",
    "autoFilter",
    "sortState",
    "dataConsolidate",
    "customSheetViews",
    "mergeCells",
    "phoneticPr",
    "conditionalFormatting",
    "dataValidations",
    "hyperlinks",
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "rowBreaks",
    "colBreaks",
    "customProperties",
    "cellWatches",
    "ignoredErrors",
    "smartTags",
    "drawing",
    "legacyDrawing",
    "legacyDrawingHF",
    "drawingHF",
    "picture",
    "oleObjects",
    "controls",
    "webPublishItems",
    "tableParts",
    "extLst",
];

// 1x1 transparent PNG
const PNG = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==", "base64");

function topLevelElements(xml) {
    const body = xml.slice(xml.indexOf("<worksheet"));
    const names = [];
    let depth = 0;
    const re = /<(\/?)([A-Za-z0-9:]+)[^>]*?(\/?)>/g;
    let match;
    while ((match = re.exec(body))) {
        const [, closing, name, selfClosing] = match;
        if (closing) {
            depth--;
        } else {
            if (depth === 1) {
                names.push(name);
            }
            if (!selfClosing) {
                depth++;
            }
        }
    }
    return names;
}

describe("github issues", () => {
    it("issue 2079 - worksheet children are written in CT_Worksheet order", async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet("Sheet1", { properties: { tabColor: { argb: "FFFF0000" } } });
        ws.columns = [{ header: "A", key: "a", width: 12 }];
        ws.getCell("A1").value = "hello";
        ws.getCell("A2").value = { text: "link", hyperlink: "https://example.com" };
        ws.getCell("A1").note = "a comment";
        ws.mergeCells("C1:D1");
        ws.autoFilter = "A1:A3";
        await ws.protect("", {});
        ws.getCell("B1").dataValidation = { type: "list", formulae: ['"a,b"'] };
        ws.addConditionalFormatting({
            ref: "A1:A3",
            rules: [{ type: "dataBar", cfvo: [{ type: "min" }, { type: "max" }], color: { argb: "FF638EC6" } }],
        });
        ws.getRow(3).addPageBreak();
        ws.headerFooter.oddFooter = "page &P";
        const imageId = wb.addImage({ buffer: PNG, extension: "png" });
        ws.addImage(imageId, "F2:G3");
        ws.addBackgroundImage(imageId);
        ws.addTable({
            name: "T1",
            ref: "H1",
            headerRow: true,
            columns: [{ name: "x" }],
            rows: [[1], [2]],
        });

        const xml = await readPart(await wb.xlsx.writeBuffer(), "xl/worksheets/sheet1.xml");
        const names = topLevelElements(xml);
        ["legacyDrawing", "drawing", "picture", "tableParts", "extLst"].forEach((name) => expect(names).to.include(name));
        names.forEach((name) => expect(CT_WORKSHEET).to.include(name));
        const positions = names.map((name) => CT_WORKSHEET.indexOf(name));
        const sorted = [...positions].sort((a, b) => a - b);
        expect(names).to.deep.equal(sorted.map((i) => CT_WORKSHEET[i]));
    });
});
