// CSV reading must stay fast on text fields.
const { Readable } = require("stream");

const ExcelJS = verquire("exceljs");

async function readCsv(text, options) {
    const wb = new ExcelJS.Workbook();
    return wb.csv.read(Readable.from([text]), options);
}

describe("CSV robustness", function () {
    this.timeout(20000);

    describe("read", () => {
        it("maps text fields without trying to parse them as dates", async () => {
            const start = Date.now();
            const ws = await readCsv("a,b,c,d,e,f,g,h,i,j\n".repeat(2000));
            expect(Date.now() - start).to.be.below(1500);
            expect(ws.getCell("A1").value).to.equal("a");
        });

        it("still reads the default date formats, numbers and special values", async () => {
            const ws = await readCsv("2020-01-05,01-05-2020,2020-01-05T10:20:30,2020-1-5,1,2.5,true,#N/A,x2020-01-05,2020-01-05 ,  ,abc\n");
            const values = ws.getRow(1).values.slice(1);
            expect(values.slice(0, 3).map((v) => v instanceof Date && v.toISOString())).to.deep.equal([
                new Date(2020, 0, 5).toISOString(),
                new Date(2020, 0, 5).toISOString(),
                new Date(2020, 0, 5, 10, 20, 30).toISOString(),
            ]);
            expect(values.slice(3)).to.deep.equal(["2020-1-5", 1, 2.5, true, { error: "#N/A" }, "x2020-01-05", "2020-01-05 ", "  ", "abc"]);
        });

        it("still tries the date formats given in the options on any field", async () => {
            const ws = await readCsv("Jan 5 2020,abc\n", { dateFormats: ["MMM D YYYY"] });
            expect(ws.getCell("A1").value).to.be.an.instanceof(Date);
            expect(ws.getCell("B1").value).to.equal("abc");
        });
    });
});
