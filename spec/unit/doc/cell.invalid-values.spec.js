// Values and models of a cell that do not fit their kind: a clear Error, or a clean value (never a broken holder).
const Excel = verquire("exceljs");

const { ValueType, FormulaType } = Excel;

describe("Cell invalid values", () => {
    let ws;
    beforeEach(() => {
        ws = new Excel.Workbook().addWorksheet("s");
    });

    describe("setting a value", () => {
        it("throws for a bigint, a symbol or a function, and leaves the cell as it was", () => {
            const cell = ws.getCell("A1");
            cell.value = 7;
            expect(() => {
                cell.value = 1n;
            }).to.throw("Invalid cell value: a bigint is not a cell value");
            expect(() => {
                cell.value = Symbol("s");
            }).to.throw("Invalid cell value: a symbol is not a cell value");
            expect(() => {
                cell.value = () => 1;
            }).to.throw("Invalid cell value: a function is not a cell value");
            expect(cell.value).to.equal(7);
        });

        it("holds a String object as its string", () => {
            ws.getCell("A1").value = new String("abc");
            expect(ws.getCell("A1").type).to.equal(ValueType.String);
            expect(ws.getCell("A1").value).to.equal("abc");
        });

        it("throws for a hyperlink whose text or tooltip is not of a cell value type", () => {
            expect(() => {
                ws.getCell("A1").value = { text: { a: 1 }, hyperlink: "https://example.com" };
            }).to.throw("Invalid hyperlink value: its text must be");
            expect(() => {
                ws.getCell("A1").value = { text: "a", hyperlink: "https://example.com", tooltip: 5 };
            }).to.throw("Invalid hyperlink value: its tooltip must be a string");
            expect(ws.getCell("A1").type).to.equal(ValueType.Null);
        });

        it("takes a null tooltip as none", () => {
            ws.getCell("A1").value = { text: "a", hyperlink: "https://example.com", tooltip: null };
            expect(ws.getCell("A1").value).to.deep.equal({ text: "a", hyperlink: "https://example.com" });
        });

        it("throws for a formula field of the wrong type", () => {
            expect(() => {
                ws.getCell("A1").value = { formula: "A2", result: { a: 1 } };
            }).to.throw("Invalid formula value: its result must be a string, a number, a boolean, a Date or an error value");
            expect(() => {
                ws.getCell("A1").value = { formula: "A2", shareType: "other" };
            }).to.throw('Invalid formula value: its shareType must be "shared" or "array"');
            expect(() => {
                ws.getCell("A1").value = { formula: 1, sharedFormula: "B1" };
            }).to.throw("Invalid formula value: its formula must be a string");
            expect(() => {
                ws.getCell("A1").value = { formula: "A2", ref: 1 };
            }).to.throw("Invalid formula value: its ref must be a string");
        });

        it("takes a null formula result as none", () => {
            ws.getCell("A1").value = { formula: "A2", result: null };
            expect(ws.getCell("A1").result).to.equal(undefined);
            expect(ws.getCell("A1").value).to.deep.equal({ formula: "A2" });
        });

        it("throws for a rich text whose runs are not runs", () => {
            expect(() => {
                ws.getCell("A1").value = { richText: [{ text: 5 }] };
            }).to.throw("Invalid rich text value: each run must be an object with a string text");
            expect(() => {
                ws.getCell("A1").value = { richText: [null] };
            }).to.throw("Invalid rich text value");
            expect(() => {
                ws.getCell("A1").value = { richText: [{ text: "a", font: 1 }] };
            }).to.throw("Invalid rich text value");
        });

        it("throws for an error value of an unknown code", () => {
            expect(() => {
                ws.getCell("A1").value = { error: "#FOO" };
            }).to.throw("Invalid error value: #FOO is not an error code");
            ws.getCell("A1").value = { error: "#REF!" };
            expect(ws.getCell("A1").type).to.equal(ValueType.Error);
        });
    });

    describe("setting a model", () => {
        it("throws for a merged cell without master, and leaves the cell as it was", () => {
            const cell = ws.getCell("B2");
            cell.value = 3;
            expect(() => {
                cell.model = { address: "B2", type: ValueType.Merge };
            }).to.throw("The model of the merged cell B2 has no master");
            expect(cell.value).to.equal(3);
        });

        it("holds a rich text whose model has the type String (the model of a rich text value) as a rich text", () => {
            ws.getCell("A1").value = { richText: [{ text: "a" }, { text: "b" }] };
            const copy = ws.getCell("C1");
            copy.model = ws.getCell("A1").model;
            expect(copy.type).to.equal(ValueType.RichText);
            expect(copy.text).to.equal("ab");
            expect(copy.html).to.equal("ab");
        });

        it("empties a String cell whose value is neither a string nor a rich text", () => {
            const cell = ws.getCell("A1");
            cell.model = { address: "A1", type: ValueType.String, value: 0 };
            expect(cell.type).to.equal(ValueType.Null);
            expect(cell.value).to.equal(null);
        });

        it("keeps the address and type of a hyperlink or formula model when its holder value is set", () => {
            const link = ws.getCell("A1");
            link.value = { text: "a", hyperlink: "https://example.com" };
            link._value.value = { text: "b", hyperlink: "https://example.org" };
            expect(link.model).to.include({ address: "A1", type: ValueType.Hyperlink, text: "b" });

            const formula = ws.getCell("A2");
            formula.value = { formula: "B1" };
            formula._value.value = { formula: "B2", result: 1 };
            expect(formula.model).to.include({ address: "A2", type: ValueType.Formula, formula: "B2", result: 1 });
        });
    });

    describe("a shared formula clone", () => {
        it("whose master has no formula has no formula and no dependencies", () => {
            ws.getCell("A1").value = 1;
            const clone = ws.getCell("A2");
            clone.value = { sharedFormula: "A1", result: 2 };
            expect(clone.formulaType).to.equal(FormulaType.Shared);
            expect(clone.formula).to.equal(undefined);
            expect(clone._value.dependencies).to.deep.equal({ ranges: null, cells: null });
        });

        it("whose master is missing has no dependencies", () => {
            const clone = ws.getCell("A2");
            clone.value = { sharedFormula: "A1", result: 2 };
            expect(clone._value.dependencies).to.deep.equal({ ranges: null, cells: null });
        });
    });
});
