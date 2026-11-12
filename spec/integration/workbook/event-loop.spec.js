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

// runs fn with fake timers, as Vitest, Jest and sinon install them: setImmediate and setTimeout callbacks are queued until the
// test runs them, which it never does here. Rejects when fn has not settled after a while (real timer).
async function withFakeTimers(fn) {
    const real = { setImmediate: global.setImmediate, setTimeout: global.setTimeout, clearTimeout: global.clearTimeout };
    const queued = [];
    global.setImmediate = (callback) => queued.push(callback);
    global.setTimeout = (callback) => queued.push(callback);
    let guard;
    try {
        return await Promise.race([
            fn(),
            new Promise((resolve, reject) => {
                guard = real.setTimeout(() => reject(new Error(`hangs under fake timers (${queued.length} callbacks queued)`)), 3000);
            }),
        ]);
    } finally {
        real.clearTimeout(guard);
        Object.assign(global, real);
    }
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

    describe("small workbooks under fake timers", () => {
        const small = () => {
            const wb = new ExcelJS.Workbook();
            const ws = wb.addWorksheet("s");
            for (let row = 1; row <= 100; row++) {
                ws.addRow([row, `text ${row}`]);
            }
            return wb;
        };

        it("writes a buffer", async () => {
            const buffer = await withFakeTimers(() => small().xlsx.writeBuffer());
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.load(buffer);
            expect(wb.getWorksheet("s").getCell("B100").value).to.equal("text 100");
        });

        it("writes to an output that does not report errors", async () => {
            const chunks = [];
            const output = { write: (chunk) => chunks.push(chunk), end() {} };
            await withFakeTimers(() => small().xlsx.write(output));
            expect(Buffer.concat(chunks).length).to.be.above(1000);
        });

        it("loads a buffer", async () => {
            const buffer = await small().xlsx.writeBuffer();
            const wb = await withFakeTimers(async () => {
                const loaded = new ExcelJS.Workbook();
                await loaded.xlsx.load(buffer);
                return loaded;
            });
            expect(wb.getWorksheet("s").getCell("B100").value).to.equal("text 100");
        });
    });
});
