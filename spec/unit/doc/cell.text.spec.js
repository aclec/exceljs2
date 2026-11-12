// The text forms of a cell (`cell.text`, `cell.toCsvString()`) for the values that used to throw or give
// "[object Object]".
const Excel = verquire("exceljs");

describe("Cell text", () => {
    let ws;
    beforeEach(() => {
        ws = new Excel.Workbook().addWorksheet("sheet");
    });

    describe("merged cell", () => {
        it("of an empty master is empty", () => {
            ws.mergeCells("A1:B1");
            expect(ws.getCell("B1").text).to.equal("");
        });

        it("is the text of its master", () => {
            ws.mergeCells("A1:B1");
            ws.getCell("A1").value = { richText: [{ text: "rich " }, { text: "text" }] };
            expect(ws.getCell("B1").text).to.equal("rich text");
            ws.getCell("A1").value = { formula: "1+1", result: 2 };
            expect(ws.getCell("B1").text).to.equal("2");
            ws.getCell("A1").value = 3;
            expect(ws.getCell("B1").text).to.equal("3");
        });
    });

    describe("formula with an error result", () => {
        it("is the error code", () => {
            const cell = ws.getCell("A1");
            cell.value = { formula: "1/0", result: { error: "#DIV/0!" } };
            expect(cell.text).to.equal("#DIV/0!");
            expect(cell.toCsvString()).to.equal("#DIV/0!");
        });
    });

    describe("rich text", () => {
        it("has a CSV string: its quoted text", () => {
            const cell = ws.getCell("A1");
            cell.value = { richText: [{ text: 'a "quoted" ' }, { text: "word" }] };
            expect(cell.toCsvString()).to.equal('"a ""quoted"" word"');
        });
    });
});
