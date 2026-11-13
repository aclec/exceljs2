// The accessors of a table that read or wrote the wrong field of its model.
const Excel = verquire("exceljs");

function addTable(props) {
    const ws = new Excel.Workbook().addWorksheet("sheet");
    return ws.addTable({ name: "T1", ref: "A1", columns: [{ name: "a" }, { name: "b" }], rows: [[1, 2]], ...props });
}

describe("Table accessors", () => {
    describe("displayName", () => {
        it("is the display name of the model", () => {
            const table = addTable({ displayName: "Shown" });
            expect(table.displayName).to.equal("Shown");
        });

        it("is the name when there is no display name", () => {
            expect(addTable({}).displayName).to.equal("T1");
        });

        it("sets the display name of the model", () => {
            const table = addTable({});
            table.displayName = "Other";
            expect(table.model.displayName).to.equal("Other");
            expect(table.displayName).to.equal("Other");
        });
    });

    describe("theme", () => {
        it("is the theme of the table style", () => {
            expect(addTable({ style: { theme: "TableStyleDark3" } }).theme).to.equal("TableStyleDark3");
            expect(addTable({}).theme).to.equal("TableStyleMedium2");
        });

        it("sets the theme of the table style", () => {
            const table = addTable({});
            table.theme = "TableStyleLight1";
            expect(table.model.style.theme).to.equal("TableStyleLight1");
        });
    });
});
