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

    describe("hyperlink", () => {
        it("is the text of a string text", () => {
            const cell = ws.getCell("A1");
            cell.value = { text: "<link>", hyperlink: "https://x" };
            expect(cell.text).to.equal("<link>");
            expect(cell.html).to.equal("&lt;link&gt;");
        });

        it("of a number is its text, as for a number cell", () => {
            const cell = ws.getCell("C1");
            cell.value = { text: 42, hyperlink: "https://x" };
            expect(cell.text).to.equal("42");
            expect(cell.html).to.equal("42");
            expect(cell.toString()).to.equal("42");
            ws.getCell("C2").value = 42;
            expect(cell.text).to.equal(ws.getCell("C2").text);
        });

        it("of a boolean, a date or an error is its text, as for such a cell", () => {
            const date = new Date(Date.UTC(2020, 0, 2));
            const texts = [true, date, { error: "#N/A" }].map((text) => {
                ws.getCell("A1").value = { text, hyperlink: "https://x" };
                ws.getCell("A2").value = text;
                return [ws.getCell("A1").text, ws.getCell("A2").text];
            });
            expect(texts).to.deep.equal([
                ["true", "true"],
                [date.toString(), date.toString()],
                ["#N/A", "#N/A"],
            ]);
        });

        it("of a rich text is the text of its runs", () => {
            const cell = ws.getCell("A1");
            cell.value = { text: { richText: [{ text: "a" }, { text: "<b>", font: { bold: true } }] }, hyperlink: "https://x" };
            expect(cell.text).to.equal("a<b>");
            expect(cell.html).to.equal("a&lt;b&gt;");
        });

        it("of a merged cell is the text of its master", () => {
            ws.mergeCells("A1:B1");
            ws.getCell("A1").value = { text: 7, hyperlink: "https://x" };
            expect(ws.getCell("B1").text).to.equal("7");
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
