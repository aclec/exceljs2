/* eslint-disable no-console -- command-line report */
// Performance suite: median wall time of the main code paths, to compare two builds of the library.
//
//   node --expose-gc scripts/bench-suite.mjs                      # this checkout (lib/, run `bun run build:lib` first)
//   node --expose-gc scripts/bench-suite.mjs --lib ../other/lib   # another build, e.g. a worktree of the base commit
//   node --expose-gc scripts/bench-suite.mjs --only doc-write,csv --runs 9
//
// Prints one line per scenario and a JSON summary on the last line ({ scenario: medianMs }).
import { createRequire } from "node:module";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { Writable } from "node:stream";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const args = process.argv.slice(2);
const arg = (name, fallback) => {
    const i = args.indexOf(`--${name}`);
    return i === -1 ? fallback : args[i + 1];
};
const libDir = path.resolve(arg("lib", path.join(root, "lib")));
const runs = Number(arg("runs", 5));
const only = arg("only", "").split(",").filter(Boolean);

const require = createRequire(import.meta.url);
const ExcelJS = require(path.join(libDir, "exceljs.nodejs.js"));

const ROWS = 30000;
const COLS = 10;
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "exceljs-bench-"));

class NullSink extends Writable {
    _write(_chunk, _enc, cb) {
        cb();
    }
}

function fillSheet(ws) {
    ws.columns = Array.from({ length: COLS }, (_, i) => ({ header: `H${i}`, key: `k${i}`, width: 12 }));
    const date = new Date(Date.UTC(2024, 0, 1));
    for (let r = 0; r < ROWS; r++) {
        const row = ws.addRow([
            r,
            `text ${r % 1000}`,
            r * 1.5,
            date,
            r % 2 === 0,
            `s${r}`,
            { formula: `A${r + 2}*2`, result: r * 2 },
            r,
            "x",
            null,
        ]);
        if (r % 10 === 0) row.font = { bold: true };
    }
}

let docBuffer;
const scenarios = {
    "doc-write": async () => {
        const wb = new ExcelJS.Workbook();
        fillSheet(wb.addWorksheet("data"));
        docBuffer = await wb.xlsx.writeBuffer();
    },
    "doc-read": async () => {
        if (!docBuffer) await scenarios["doc-write"]();
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(docBuffer);
    },
    "stream-write": async () => {
        const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ stream: new NullSink(), useSharedStrings: true, useStyles: true });
        const ws = wb.addWorksheet("data");
        ws.columns = Array.from({ length: COLS }, (_, i) => ({ header: `H${i}`, key: `k${i}` }));
        for (let r = 0; r < ROWS * 2; r++) ws.addRow([r, `text ${r % 1000}`, r * 1.5, new Date(r), r % 2 === 0, `s${r}`]).commit();
        ws.commit();
        await wb.commit();
    },
    "stream-read": async () => {
        const reader = new ExcelJS.stream.xlsx.WorkbookReader(path.join(root, "spec/integration/data/huge.xlsx"));
        let rows = 0;
        for await (const ws of reader) for await (const row of ws) rows += row ? 1 : 0;
        if (rows === 0) throw new Error("no rows read");
    },
    csv: async () => {
        const wb = new ExcelJS.Workbook();
        fillSheet(wb.addWorksheet("data"));
        const file = path.join(tmp, "bench.csv");
        await wb.csv.writeFile(file);
        await new ExcelJS.Workbook().csv.readFile(file);
    },
    "cell-access": async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet("data");
        for (let r = 1; r <= 2000; r++) for (let c = 1; c <= 50; c++) ws.getCell(r, c).value = r * c;
        let sum = 0;
        for (let i = 0; i < 5; i++) ws.eachRow((row) => row.eachCell((cell) => (sum += cell.value)));
        for (let r = 1; r <= 2000; r++) sum += ws.getCell(`B${r}`).value;
        if (!sum) throw new Error("unexpected");
    },
};

const median = (xs) => [...xs].sort((a, b) => a - b)[Math.floor(xs.length / 2)];
const summary = {};
try {
    for (const [name, run] of Object.entries(scenarios)) {
        if (only.length && !only.includes(name)) continue;
        await run(); // warm-up
        const times = [];
        for (let i = 0; i < runs; i++) {
            globalThis.gc?.();
            const t0 = performance.now();
            await run();
            times.push(performance.now() - t0);
        }
        summary[name] = Math.round(median(times));
        console.log(`${name.padEnd(14)} median ${String(summary[name]).padStart(6)} ms  (${times.map((t) => Math.round(t)).join(", ")})`);
    }
    console.log(JSON.stringify(summary));
} finally {
    fs.rmSync(tmp, { recursive: true, force: true });
}
