const { baseWorkbookBuffer, tamper, readPart } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

const SHEET = "xl/worksheets/sheet1.xml";

// x14 conditional formats as written by LibreOffice / Excel for expression rules
const x14Rule = (attributes, formula) =>
    `<x14:cfRule ${attributes}><xm:f>${formula}</xm:f>` +
    '<x14:dxf><font><b/><color rgb="FFFF0000"/></font>' +
    '<fill><patternFill><bgColor rgb="FFFFFF00"/></patternFill></fill></x14:dxf></x14:cfRule>';

const EXT_LST =
    '<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" ' +
    'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
    "<x14:conditionalFormattings>" +
    '<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">' +
    `${x14Rule('type="expression" priority="1" id="{00000000-0000-0000-0000-000000000001}"', "A1=1")}` +
    `${x14Rule('type="cellIs" priority="2" operator="greaterThan" id="{00000000-0000-0000-0000-000000000002}"', "5")}` +
    "<xm:sqref>A1:A5</xm:sqref></x14:conditionalFormatting>" +
    "</x14:conditionalFormattings></ext></extLst>";

async function load(buffer) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    return wb;
}

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
        expect(rule.style.font.bold).to.equal(true);
        expect(rule.style.font.color).to.deep.equal({ argb: "FFFF0000" });
        expect(rule.style.fill.bgColor).to.deep.equal({ argb: "FFFFFF00" });
    });
}

describe("github issues", () => {
    it("issue 1767 - x14 expression conditional formats keep formula and style", async () => {
        const source = await tamper(await baseWorkbookBuffer(), {
            [SHEET]: (xml) => xml.replace("</worksheet>", `${EXT_LST}</worksheet>`),
        });
        const wb = await load(source);
        checkRules(wb);

        const buffer = await wb.xlsx.writeBuffer();
        const xml = await readPart(buffer, SHEET);
        expect(xml).to.match(/<cfRule type="expression" dxfId="\d+" priority="1"><formula>A1=1<\/formula><\/cfRule>/);
        checkRules(await load(buffer));
    });
});
