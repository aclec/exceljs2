// Builds small, readable xlsx fixtures for robustness/security tests:
// a valid workbook is written with ExcelJS, then individual XML parts are patched.
// Keeping fixtures as code (instead of binary files) makes them reviewable.
const JSZip = require("jszip");

const ExcelJS = verquire("exceljs");

async function baseWorkbookBuffer(build) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "hello";
    if (build) {
        await build(wb, ws);
    }
    return wb.xlsx.writeBuffer();
}

// patches: {'xl/worksheets/sheet1.xml': xml => newXml, ...}
async function tamper(buffer, patches) {
    const zip = await JSZip.loadAsync(buffer);
    for (const [path, patch] of Object.entries(patches)) {
        const file = zip.file(path);
        if (!file) {
            throw new Error(`xlsx-tamper: no part ${path}`);
        }
        zip.file(path, patch(await file.async("string")));
    }
    return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

async function readPart(buffer, path) {
    const zip = await JSZip.loadAsync(buffer);
    return zip.file(path).async("string");
}

module.exports = { baseWorkbookBuffer, tamper, readPart };
