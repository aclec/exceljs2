// xlsx.load(data, { base64: true }): the base64 text of the file, given as a string or as its bytes.
const ExcelJS = verquire("exceljs");

async function base64Text() {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet("s").getCell("A1").value = "hello";
    return Buffer.from(await wb.xlsx.writeBuffer()).toString("base64");
}

async function loadedA1(data) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(data, { base64: true });
    return wb.getWorksheet("s").getCell("A1").value;
}

describe("xlsx.load with base64", () => {
    it("reads the base64 text given as a string, a Buffer, a Uint8Array, an ArrayBuffer, a DataView or a Blob", async () => {
        const text = await base64Text();
        const bytes = new TextEncoder().encode(text);
        const inputs = {
            string: text,
            Buffer: Buffer.from(text),
            Uint8Array: bytes,
            ArrayBuffer: bytes.buffer,
            DataView: new DataView(bytes.buffer),
            Blob: new Blob([text]),
        };
        for (const [kind, data] of Object.entries(inputs)) {
            expect(await loadedA1(data), kind).to.equal("hello");
        }
    });

    it("rejects data that is not base64 text with an error", async () => {
        let thrown;
        try {
            await loadedA1(12);
        } catch (e) {
            thrown = e;
        }
        expect(thrown).to.be.an.instanceOf(Error);
    });
});
