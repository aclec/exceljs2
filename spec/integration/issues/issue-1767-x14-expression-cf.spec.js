const fs = require("fs");
const os = require("os");
const path = require("path");

const { baseWorkbookBuffer, tamper, readPart } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

const SHEET = "xl/worksheets/sheet1.xml";

// x14 conditional formats as written by LibreOffice / Excel for expression rules
const x14Rule = (attributes, formula) =>
    `<x14:cfRule ${attributes}><xm:f>${formula}</xm:f>` +
    '<x14:dxf><font><b/><color rgb="FFFF0000"/></font>' +
    '<fill><patternFill><bgColor rgb="FFFFFF00"/></patternFill></fill></x14:dxf></x14:cfRule>';

const extLst = (...conditionalFormattings) =>
    '<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" ' +
    'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
    `<x14:conditionalFormattings>${conditionalFormattings.join("")}</x14:conditionalFormattings></ext></extLst>`;

const x14Cf = (sqref, ...rules) =>
    '<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">' +
    `${rules.join("")}<xm:sqref>${sqref}</xm:sqref></x14:conditionalFormatting>`;

const EXT_LST = extLst(
    x14Cf(
        "A1:A5",
        x14Rule('type="expression" priority="1" id="{00000000-0000-0000-0000-000000000001}"', "A1=1"),
        x14Rule('type="cellIs" priority="2" operator="greaterThan" id="{00000000-0000-0000-0000-000000000002}"', "5")
    )
);

async function load(buffer) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    return wb;
}

// comparable models: writing assigns dxfId and ref to the rules, and the x14 ids are regenerated
const snapshot = (conditionalFormattings) =>
    JSON.parse(JSON.stringify(conditionalFormattings.map((cf) => ({ ...cf, rules: cf.rules.map(({ x14Id, ...rule }) => rule) }))));

const mainSection = (xml) => (xml.match(/<conditionalFormatting [\s\S]*<\/conditionalFormatting>/) || [""])[0];
const extSection = (xml) => (xml.match(/<x14:conditionalFormattings>[\s\S]*<\/x14:conditionalFormattings>/) || [""])[0];

function checkRules(wb) {
    const [cf] = wb.getWorksheet("Sheet1").conditionalFormattings;
    expect(cf.ref).to.equal("A1:A5");
    const [expression, cellIs] = cf.rules;
    expect(expression.type).to.equal("expression");
    expect(expression.formulae).to.deep.equal(["A1=1"]);
    expect(cellIs.type).to.equal("cellIs");
    expect(cellIs.operator).to.equal("greaterThan");
    expect(cellIs.formulae).to.deep.equal(["5"]);
    cf.rules.forEach((rule) => {
        expect(rule.x14).to.equal(true);
        expect(rule.style.font.bold).to.equal(true);
        expect(rule.style.font.color).to.deep.equal({ argb: "FFFF0000" });
        expect(rule.style.fill.bgColor).to.deep.equal({ argb: "FFFFFF00" });
    });
}

describe("github issues", () => {
    describe("issue 1767 - x14 expression conditional formats", () => {
        it("keep their formula and style, and are written back in the x14 extension", async () => {
            const source = await tamper(await baseWorkbookBuffer(), {
                [SHEET]: (xml) => xml.replace("</worksheet>", `${EXT_LST}</worksheet>`),
            });
            const wb = await load(source);
            checkRules(wb);
            const before = snapshot(wb.getWorksheet("Sheet1").conditionalFormattings);

            const buffer = await wb.xlsx.writeBuffer();
            const xml = await readPart(buffer, SHEET);
            expect(xml).not.to.include("<conditionalFormatting");
            expect(xml).not.to.include("<cfRule");
            expect(extSection(xml)).to.match(
                /<x14:cfRule type="expression" priority="1" id="\{[0-9A-F-]{36}\}"><xm:f>A1=1<\/xm:f><x14:dxf><font><b\/>/
            );
            expect(extSection(xml)).to.match(
                /<x14:cfRule type="cellIs" priority="2" operator="greaterThan" id="\{[0-9A-F-]{36}\}"><xm:f>5<\/xm:f>/
            );
            const back = await load(buffer);
            checkRules(back);
            expect(snapshot(back.getWorksheet("Sheet1").conditionalFormattings)).to.deep.equal(before);
        });

        it("rules referencing another sheet found only in the extension stay there, beside the main rules", async () => {
            const base = await baseWorkbookBuffer((wb, ws) => {
                wb.addWorksheet("Other sheet").getCell("A1").value = 1;
                ws.addConditionalFormatting({
                    ref: "A1:A5",
                    rules: [
                        { type: "expression", formulae: ["A1>2"], style: { font: { italic: true } } },
                        {
                            type: "iconSet",
                            iconSet: "3Triangles",
                            cfvo: [
                                { type: "percent", value: 0 },
                                { type: "percent", value: 33 },
                                { type: "percent", value: 67 },
                            ],
                        },
                    ],
                });
            });
            const crossSheet = x14Rule(
                'type="expression" priority="9" id="{00000000-0000-0000-0000-00000000000A}"',
                "'Other sheet'!$A$1=A1"
            );
            const source = await tamper(base, {
                [SHEET]: (xml) =>
                    xml.replace(
                        "</x14:conditionalFormatting></x14:conditionalFormattings>",
                        `</x14:conditionalFormatting>${x14Cf("B1:B5", crossSheet)}</x14:conditionalFormattings>`
                    ),
            });
            expect(await readPart(source, SHEET)).to.include("00000000-0000-0000-0000-00000000000A");

            const wb = await load(source);
            const cfs = wb.getWorksheet("Sheet1").conditionalFormattings;
            expect(cfs.map((cf) => [cf.ref, cf.rules.map((rule) => [rule.type, rule.x14])])).to.deep.equal([
                [
                    "A1:A5",
                    [
                        ["expression", undefined],
                        ["iconSet", undefined],
                    ],
                ],
                ["B1:B5", [["expression", true]]],
            ]);
            const before = snapshot(cfs);

            const buffer = await wb.xlsx.writeBuffer();
            const xml = await readPart(buffer, SHEET);
            expect(mainSection(xml)).to.include("<formula>A1&gt;2</formula>");
            expect(extSection(xml)).to.include('<x14:cfRule type="iconSet" priority="2"');
            expect(mainSection(xml)).not.to.include("Other sheet");
            expect(mainSection(xml)).not.to.include('sqref="B1:B5"');
            expect(extSection(xml)).to.match(
                /<x14:cfRule type="expression" priority="9" id="\{[0-9A-F-]{36}\}"><xm:f>&apos;Other sheet&apos;!\$A\$1=A1<\/xm:f><x14:dxf>[\s\S]*?<\/x14:dxf><\/x14:cfRule><xm:sqref>B1:B5<\/xm:sqref>/
            );

            const back = await load(buffer);
            expect(snapshot(back.getWorksheet("Sheet1").conditionalFormattings)).to.deep.equal(before);
        });

        it("rules created with a formula referencing another sheet are written in the extension", async () => {
            const build = (wb) => {
                const ws = wb.addWorksheet("Sheet1");
                wb.addWorksheet("Other");
                ws.addConditionalFormatting({
                    ref: "A1:A5",
                    rules: [
                        { type: "cellIs", operator: "equal", formulae: ["Other!$A$1"], style: { font: { bold: true } } },
                        { type: "expression", formulae: ['A1="Hi!"'], style: { font: { bold: true } } },
                    ],
                });
                return ws;
            };
            const check = (xml) => {
                expect(mainSection(xml)).to.match(
                    /<cfRule type="expression" dxfId="\d+" priority="2"><formula>A1=&quot;Hi!&quot;<\/formula>/
                );
                expect(mainSection(xml)).not.to.include("Other!");
                expect(extSection(xml)).to.match(
                    /<x14:cfRule type="cellIs" priority="1" operator="equal" id="\{[0-9A-F-]{36}\}"><xm:f>Other!\$A\$1<\/xm:f><x14:dxf><font><b\/><\/font><\/x14:dxf><\/x14:cfRule><xm:sqref>A1:A5<\/xm:sqref>/
                );
            };

            const wb = new ExcelJS.Workbook();
            build(wb);
            const buffer = await wb.xlsx.writeBuffer();
            check(await readPart(buffer, SHEET));
            const [cf] = (await load(buffer)).getWorksheet("Sheet1").conditionalFormattings;
            expect(cf.rules.map((rule) => [rule.type, rule.formulae, rule.x14])).to.deep.equal([
                ["expression", ['A1="Hi!"'], undefined],
                ["cellIs", ["Other!$A$1"], true],
            ]);

            const dir = fs.mkdtempSync(path.join(os.tmpdir(), "exceljs-x14-"));
            try {
                const filename = path.join(dir, "stream.xlsx");
                const streamed = new ExcelJS.stream.xlsx.WorkbookWriter({ filename, useStyles: true });
                build(streamed).commit();
                await streamed.commit();
                check(await readPart(fs.readFileSync(filename), SHEET));
            } finally {
                fs.rmSync(dir, { recursive: true, force: true });
            }
        });

        it("a workbook written again after a cross-sheet rule is added gets the extension", async () => {
            const wb = new ExcelJS.Workbook();
            const ws = wb.addWorksheet("Sheet1");
            ws.addConditionalFormatting({
                ref: "A1",
                rules: [{ type: "expression", formulae: ["A1=1"], style: { font: { bold: true } } }],
            });
            expect(await readPart(await wb.xlsx.writeBuffer(), SHEET)).not.to.include("extLst");

            ws.addConditionalFormatting({
                ref: "B1",
                rules: [{ type: "expression", formulae: ["Other!A1=1"], style: { font: { bold: true } } }],
            });
            const xml = await readPart(await wb.xlsx.writeBuffer(), SHEET);
            expect(extSection(xml)).to.include("<xm:f>Other!A1=1</xm:f>");
            expect(ws.conditionalFormattings).not.to.have.property("hasExtContent");
        });
    });
});
