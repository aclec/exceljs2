"use strict";

// Checks the output of `bun run build` (run it first).
const fs = require("fs");
const path = require("path");
const { spawnSync } = require("child_process");

const DIST = path.resolve(__dirname, "../../dist");
const BUNDLES = ["exceljs.js", "exceljs.min.js", "exceljs.bare.js", "exceljs.bare.min.js"];
const { version } = require("../../package.json");

describe("ExcelJS", () => {
    describe("dist folder", () => {
        it("contains exactly the bundles, their source maps and LICENSE", () => {
            const expected = ["LICENSE", ...BUNDLES, ...BUNDLES.map((name) => `${name}.map`)].sort();
            expect(fs.readdirSync(DIST).sort()).to.deep.equal(expected);
        });

        it("no longer contains the ES5 build (dropped in v5)", () => {
            expect(fs.existsSync(path.join(DIST, "es5"))).to.be.false();
        });

        BUNDLES.forEach((name) => {
            describe(name, () => {
                const file = path.join(DIST, name);

                it("has a version banner and links its source map", () => {
                    const code = fs.readFileSync(file, "utf8");
                    expect(code.startsWith(`/*! ExcelJS v${version} |`)).to.be.true();
                    expect(code.trimEnd().endsWith(`//# sourceMappingURL=${name}.map`)).to.be.true();
                    const map = JSON.parse(fs.readFileSync(`${file}.map`, "utf8"));
                    expect(map.sources).to.include("../src/doc/workbook.ts");
                });

                it("exports ExcelJS through module.exports (UMD) and works", async () => {
                    const ExcelJS = require(file);
                    expect(ExcelJS.Workbook).to.be.a("function");
                    expect(ExcelJS.ValueType.Number).to.equal(2);

                    const wb = new ExcelJS.Workbook();
                    wb.addWorksheet("blort").getCell("A1").value = "Hello, World!";
                    const buffer = await wb.xlsx.writeBuffer();
                    const wb2 = new ExcelJS.Workbook();
                    await wb2.xlsx.load(buffer);
                    expect(wb2.getWorksheet("blort").getCell("A1").value).to.equal("Hello, World!");
                });

                it("writes a chart created by addChart and a chartsheet", async () => {
                    const ExcelJS = require(file);
                    const wb = new ExcelJS.Workbook();
                    const ws = wb.addWorksheet("My Data");
                    ws.getCell("A1").value = 1;
                    ws.addChart({ type: "pie", anchor: "C2:H12", series: [{ values: "A1" }] });
                    wb.addChartsheet("Chart", { type: "line", series: [{ values: "'My Data'!A1" }] });
                    const wb2 = new ExcelJS.Workbook();
                    await wb2.xlsx.load(await wb.xlsx.writeBuffer());
                    expect(wb2.getWorksheet("My Data").model.charts).to.have.length(1);
                    expect(wb2.chartsheets).to.deep.equal([{ id: 2, name: "Chart", state: "visible" }]);
                });

                it("keeps merges when rows are inserted above them (class names are minified)", () => {
                    const ExcelJS = require(file);
                    const ws = new ExcelJS.Workbook().addWorksheet("s");
                    ws.getCell("A2").value = "merged";
                    ws.mergeCells("A2:B2");
                    ws.insertRow(1, ["new"]);
                    expect(ws.getCell("B3").isMerged).to.equal(true);
                    expect(ws.getCell("B3").master.address).to.equal("A3");
                });

                it("loads without a warning (no access to globalThis.localStorage, which Node.js 25+ warns about)", () => {
                    const { stderr, status } = spawnSync(
                        process.execPath,
                        ["-e", `const ExcelJS = require(${JSON.stringify(file)}); new ExcelJS.Workbook().addWorksheet("s");`],
                        { encoding: "utf8" }
                    );
                    expect(status).to.equal(0);
                    expect(stderr).to.equal("");
                    expect(fs.readFileSync(file, "utf8")).not.to.include("localStorage");
                });

                it("compresses with the bundled fflate, not node:zlib", () => {
                    // zlib is an empty module in the bundle, so src/utils/deflate.ts selects fflate
                    const code = fs.readFileSync(file, "utf8");
                    expect(code).not.to.match(/require\(["'](node:)?zlib["']\)/);
                    expect(code).to.include("deflateRawSync");
                });
            });
        });
    });
});
