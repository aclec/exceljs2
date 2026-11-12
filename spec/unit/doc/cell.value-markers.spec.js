// The kind of an object value is told by its marker properties: their presence and type, not their truthiness.
const Excel = verquire("exceljs");

const { ValueType, FormulaType } = Excel;

describe("Cell value markers", () => {
    let ws;
    beforeEach(() => {
        ws = new Excel.Workbook().addWorksheet("s");
    });

    describe("a hyperlink", () => {
        it("with an empty text is a hyperlink", () => {
            const cell = ws.getCell("A1");
            cell.value = { text: "", hyperlink: "https://example.com" };
            expect(cell.type).to.equal(ValueType.Hyperlink);
            expect(cell.hyperlink).to.equal("https://example.com");
            expect(cell.text).to.equal("");
            expect(cell.value).to.deep.equal({ text: "", hyperlink: "https://example.com" });
        });

        it("with a zero or false text is a hyperlink", () => {
            ws.getCell("A1").value = { text: 0, hyperlink: "https://example.com" };
            ws.getCell("A2").value = { text: false, hyperlink: "https://example.com" };
            expect(ws.getCell("A1").type).to.equal(ValueType.Hyperlink);
            expect(ws.getCell("A1").text).to.equal("0");
            expect(ws.getCell("A2").type).to.equal(ValueType.Hyperlink);
            expect(ws.getCell("A2").text).to.equal("false");
        });

        it("without text, or with a target that is not a string, is a JSON value", () => {
            ws.getCell("A1").value = { hyperlink: "https://example.com" };
            ws.getCell("A2").value = { text: "a", hyperlink: 1 };
            ws.getCell("A3").value = { text: null, hyperlink: "https://example.com" };
            for (const address of ["A1", "A2", "A3"]) {
                expect(ws.getCell(address).type).to.equal(ValueType.String);
            }
            expect(ws.getCell("A1").text).to.equal('{"hyperlink":"https://example.com"}');
        });

        it("is written and read back with an empty text", async () => {
            ws.getCell("A1").value = { text: "", hyperlink: "https://example.com" };
            const wb = new Excel.Workbook();
            await wb.xlsx.load(await ws.workbook.xlsx.writeBuffer());
            const cell = wb.getWorksheet("s").getCell("A1");
            expect(cell.type).to.equal(ValueType.Hyperlink);
            expect(cell.hyperlink).to.equal("https://example.com");
        });
    });

    describe("a shared string", () => {
        it("of index 0 is a shared string", () => {
            const cell = ws.getCell("A1");
            cell.value = { sharedString: 0 };
            expect(cell.type).to.equal(ValueType.SharedString);
            expect(cell.value).to.deep.equal({ sharedString: 0 });
        });

        it("of an index that is not a number is a JSON value", () => {
            const cell = ws.getCell("A1");
            cell.value = { sharedString: "0" };
            expect(cell.type).to.equal(ValueType.String);
        });
    });

    describe("a formula", () => {
        it("that is empty is a formula cell without formula (its result is its text)", () => {
            const cell = ws.getCell("A1");
            cell.value = { formula: "", result: 3 };
            expect(cell.type).to.equal(ValueType.Formula);
            expect(cell.formulaType).to.equal(FormulaType.None);
            expect(cell.formula).to.equal(undefined);
            expect(cell.result).to.equal(3);
            expect(cell.text).to.equal("3");
        });

        it("with a falsy result keeps it", () => {
            const cell = ws.getCell("A1");
            cell.value = { formula: "A2", result: 0 };
            expect(cell.type).to.equal(ValueType.Formula);
            expect(cell.result).to.equal(0);
            expect(cell.text).to.equal("0");
        });

        it("that is not a string is a JSON value", () => {
            const cell = ws.getCell("A1");
            cell.value = { formula: 1 };
            expect(cell.type).to.equal(ValueType.String);
        });
    });

    describe("an error", () => {
        it("with a code that is not a string is a JSON value", () => {
            const cell = ws.getCell("A1");
            cell.value = { error: 0 };
            expect(cell.type).to.equal(ValueType.String);
        });
    });

    describe("a rich text", () => {
        it("with no run is a rich text", () => {
            const cell = ws.getCell("A1");
            cell.value = { richText: [] };
            expect(cell.type).to.equal(ValueType.RichText);
            expect(cell.text).to.equal("");
        });

        it("whose runs are not an array is a JSON value", () => {
            const cell = ws.getCell("A1");
            cell.value = { richText: "abc" };
            expect(cell.type).to.equal(ValueType.String);
            expect(cell.text).to.equal('{"richText":"abc"}');
        });
    });
});
