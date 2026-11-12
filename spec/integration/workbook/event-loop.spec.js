// xlsx.load / read / readFile and xlsx.write / writeBuffer / writeFile give the event loop a turn now and then, so that a
// server stays responsive while a large workbook is read or written.
const ExcelJS = verquire("exceljs");

// runs fn while a 5 ms interval timer counts its ticks
async function ticksDuring(fn) {
    let ticks = 0;
    const timer = setInterval(() => {
        ticks++;
    }, 5);
    const start = Date.now();
    try {
        await fn();
    } finally {
        clearInterval(timer);
    }
    return { ticks, ms: Date.now() - start };
}

describe("Event loop during large reads and writes", function () {
    this.timeout(60000);

    let buffer;
    before(async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet("data");
        for (let row = 1; row <= 40000; row++) {
            ws.addRow([row, `text ${row}`, row / 7, new Date(Date.UTC(2020, 0, 1 + (row % 365))), row % 2 === 0]);
        }
        wb.addWorksheet("small").getCell("A1").value = 1;
        buffer = await wb.xlsx.writeBuffer();
    });

    it("lets timers run while a workbook is loaded", async () => {
        let wb;
        const { ticks, ms } = await ticksDuring(async () => {
            wb = new ExcelJS.Workbook();
            await wb.xlsx.load(buffer);
        });
        expect(wb.getWorksheet("data").getCell("B40000").value).to.equal("text 40000");
        // one turn about every 25 ms of work
        expect(ticks, `${ticks} ticks in ${ms} ms`).to.be.at.least(Math.min(3, Math.floor(ms / 100)));
        expect(ticks).to.be.at.least(1);
    });

    it("lets timers run while a workbook is written", async () => {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buffer);
        let written;
        const { ticks, ms } = await ticksDuring(async () => {
            written = await wb.xlsx.writeBuffer();
        });
        const reloaded = new ExcelJS.Workbook();
        await reloaded.xlsx.load(written);
        expect(reloaded.getWorksheet("data").getCell("A40000").value).to.equal(40000);
        expect(ticks, `${ticks} ticks in ${ms} ms`).to.be.at.least(Math.min(3, Math.floor(ms / 100)));
        expect(ticks).to.be.at.least(1);
    });
});
