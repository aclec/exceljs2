// Builds small, readable xlsx fixtures for robustness/security tests:
// a valid workbook is written with ExcelJS, then individual XML parts are patched.
// Keeping fixtures as code (instead of binary files) makes them reviewable.
// The zip is handled with fflate directly, independently of ExcelJS's own zip code.
const { strFromU8, strToU8, unzipSync, zipSync } = require("fflate");

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

// the files of a zip, in archive order: {name: Uint8Array}
function unzipParts(buffer) {
    const parts = unzipSync(new Uint8Array(buffer));
    for (const name of Object.keys(parts)) {
        if (name.endsWith("/")) {
            delete parts[name];
        }
    }
    return parts;
}

// a zip of parts ({name: string | Uint8Array}), in key order
function zipParts(parts) {
    const files = {};
    for (const [name, data] of Object.entries(parts)) {
        files[name] = typeof data === "string" ? strToU8(data) : data;
    }
    return Buffer.from(zipSync(files, { level: 6 }));
}

// rebuilds the zip after edit(parts) changed, added or deleted parts
function editParts(buffer, edit) {
    const parts = unzipParts(buffer);
    edit(parts);
    return zipParts(parts);
}

// patches: {'xl/worksheets/sheet1.xml': xml => newXml, ...}
async function tamper(buffer, patches) {
    return editParts(buffer, (parts) => {
        for (const [path, patch] of Object.entries(patches)) {
            if (!parts[path]) {
                throw new Error(`xlsx-tamper: no part ${path}`);
            }
            parts[path] = patch(strFromU8(parts[path]));
        }
    });
}

async function readPart(buffer, path) {
    const part = unzipParts(buffer)[path];
    if (!part) {
        throw new Error(`xlsx-tamper: no part ${path}`);
    }
    return strFromU8(part);
}

module.exports = { baseWorkbookBuffer, tamper, readPart, unzipParts, zipParts, editParts };
