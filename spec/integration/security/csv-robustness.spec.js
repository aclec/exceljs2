// CSV reading must stay fast on text fields, and CSV writing must report the errors of its destination.
const fs = require("fs");
const os = require("os");
const path = require("path");
const { Readable, Writable } = require("stream");

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

    describe("write", () => {
        function workbook() {
            const wb = new ExcelJS.Workbook();
            wb.addWorksheet("s").addRow(["a", "b"]);
            return wb;
        }

        it("rejects when the destination stream fails", async () => {
            const destination = new Writable({
                write(chunk, encoding, callback) {
                    callback(new Error("disk full"));
                },
            });
            let error;
            await workbook()
                .csv.write(destination)
                .catch((e) => {
                    error = e;
                });
            expect(error).to.be.an("error");
            expect(error.message).to.equal("disk full");
        });

        it("rejects when the file cannot be created", async () => {
            const dir = await fs.promises.mkdtemp(path.join(os.tmpdir(), "exceljs-csv-"));
            try {
                let error;
                await workbook()
                    .csv.writeFile(path.join(dir, "missing", "out.csv"))
                    .catch((e) => {
                        error = e;
                    });
                expect(error).to.be.an("error");
                expect(error.code).to.equal("ENOENT");
            } finally {
                await fs.promises.rm(dir, { recursive: true, force: true });
            }
        });

        it("still writes a file", async () => {
            const dir = await fs.promises.mkdtemp(path.join(os.tmpdir(), "exceljs-csv-"));
            try {
                const file = path.join(dir, "out.csv");
                await workbook().csv.writeFile(file);
                expect(await fs.promises.readFile(file, "utf8")).to.equal("a,b");
            } finally {
                await fs.promises.rm(dir, { recursive: true, force: true });
            }
        });
    });
});
