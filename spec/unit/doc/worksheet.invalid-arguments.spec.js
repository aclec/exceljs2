// Invalid arguments of the worksheet and row methods: a clear Error before anything changes (not a TypeError deep inside).
const Excel = verquire("exceljs");

const { ValueType } = Excel;

describe("Worksheet and Row invalid arguments", () => {
    let ws;
    beforeEach(() => {
        ws = new Excel.Workbook().addWorksheet("s");
    });

    it("eachRow and eachCell throw without an iteratee function", () => {
        ws.getCell("A1").value = 1;
        expect(() => ws.eachRow()).to.throw(Error, "eachRow needs an iteratee function");
        expect(() => ws.eachRow({ includeEmpty: true })).to.throw(Error, "eachRow needs an iteratee function");
        expect(() => ws.getRow(1).eachCell()).to.throw(Error, "eachCell needs an iteratee function");
        expect(() => ws.getRow(1).eachCell({ includeEmpty: true })).to.throw(Error, "eachCell needs an iteratee function");
    });

    it("spliceColumns throws for a start that is not a column number or a count that is not an integer >= 0", () => {
        ws.getCell("A1").value = 1;
        ws.getCell("B1").value = 2;
        expect(() => ws.spliceColumns(0, 1)).to.throw(Error, "Invalid column number: 0");
        expect(() => ws.spliceColumns(1.5, 1)).to.throw(Error, "Invalid column number: 1.5");
        expect(() => ws.spliceColumns(1, -1)).to.throw(Error, "Invalid count of columns: -1");
        expect(ws.getRow(1).values).to.deep.equal([, 1, 2]);
    });

    it("spliceColumns works on a worksheet without any cell", () => {
        ws.spliceColumns(1, 1);
        expect(ws.columns).to.equal(null);
        // a row without cells
        ws.getRow(1);
        ws.spliceColumns(1, 1);
        expect(ws.columns).to.equal(null);
    });

    it("a row model whose first cell has no address throws", () => {
        expect(() => {
            ws.getRow(2).model = { number: 2, cells: [{ type: ValueType.Number, value: 1 }] };
        }).to.throw(Error, "The first cell of row 2 has no address");
    });
});
