const http = require("http");
const { Readable } = require("stream");
const testutils = require("../utils/index");

const Excel = verquire("exceljs");

describe("HTTP streaming", () => {
    let server;
    let port;
    before((done) => {
        server = http.createServer((req, res) => {
            const wb = testutils.createTestBook(new Excel.Workbook(), "xlsx");
            res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            res.setHeader("Content-Disposition", "attachment; filename=Report.xlsx");
            wb.xlsx.write(res).then(() => res.end());
        });
        server.listen(0, "127.0.0.1", () => {
            ({ port } = server.address());
            done();
        });
    });

    after(() => server.close());

    it("downloads a workbook", async function () {
        this.timeout(5000);
        const res = await fetch(`http://127.0.0.1:${port}/workbook`);
        const wb2 = new Excel.Workbook();
        await wb2.xlsx.read(Readable.fromWeb(res.body));
        testutils.checkTestBook(wb2, "xlsx");
    });
});
