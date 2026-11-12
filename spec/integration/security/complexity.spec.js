// Regression tests for hostile input that used to cost super-linear time or memory
// (ReDoS, quadratic lookups, per-index traversals of sparse arrays).
const { baseWorkbookBuffer, tamper, zipParts } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

const SHEET = "xl/worksheets/sheet1.xml";

async function load(buffer) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    return wb;
}

async function timed(fn) {
    const start = Date.now();
    const result = await fn();
    return { result, ms: Date.now() - start };
}

describe("hostile input: algorithmic complexity", function () {
    this.timeout(20000);

    describe("conditional formatting thresholds", () => {
        const LONG = `${"1".repeat(100000)}x`;

        it("parses a long non-numeric <cfvo val> in linear time", async () => {
            const buffer = await tamper(
                await baseWorkbookBuffer((wb, ws) => {
                    ws.addConditionalFormatting({
                        ref: "A1:A2",
                        rules: [
                            {
                                type: "colorScale",
                                priority: 1,
                                cfvo: [{ type: "num", value: 5 }, { type: "max" }],
                                color: [{ argb: "FFFF0000" }, { argb: "FF00FF00" }],
                            },
                        ],
                    });
                }),
                { [SHEET]: (xml) => xml.replace('<cfvo type="num" val="5"', `<cfvo type="num" val="${LONG}"`) }
            );
            const { result: wb, ms } = await timed(() => load(buffer));
            expect(wb.getWorksheet(1).conditionalFormattings[0].rules[0].cfvo[0].value).to.equal(LONG);
            expect(ms).to.be.below(1000);
        });

        it("parses a long non-numeric x14 <xm:f> threshold in linear time", async () => {
            const buffer = await tamper(
                await baseWorkbookBuffer((wb, ws) => {
                    ws.addConditionalFormatting({
                        ref: "A1:A2",
                        rules: [{ type: "dataBar", cfvo: [{ type: "num", value: 5 }, { type: "max" }], color: { argb: "FF00FF00" } }],
                    });
                }),
                {
                    [SHEET]: (xml) => {
                        expect(xml).to.include("<xm:f>5</xm:f>");
                        return xml.replace("<xm:f>5</xm:f>", `<xm:f>${LONG}</xm:f>`);
                    },
                }
            );
            const { ms } = await timed(() => load(buffer));
            expect(ms).to.be.below(1000);
        });
    });

    it("matches worksheets to many <sheet> elements in linear time", async () => {
        const parts = { "xl/workbook.xml": `<workbook><sheets>${'<sheet name=""/>'.repeat(400000)}</sheets></workbook>` };
        for (let i = 1; i <= 2000; i++) {
            parts[`xl/worksheets/sheet${i}.xml`] = "<worksheet/>";
        }
        const { ms } = await timed(() => load(zipParts(parts)));
        expect(ms).to.be.below(1500);
    });
});
