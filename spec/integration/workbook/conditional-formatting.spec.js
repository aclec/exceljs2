// Conditional formats round-tripped through their x14 extension (<extLst> of the worksheet).
const { baseWorkbookBuffer, tamper, readPart } = require("../../utils/xlsx-tamper");

const ExcelJS = verquire("exceljs");

const SHEET = "xl/worksheets/sheet1.xml";

const extLst = (...conditionalFormattings) =>
    '<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" ' +
    'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
    `<x14:conditionalFormattings>${conditionalFormattings.join("")}</x14:conditionalFormattings></ext></extLst>`;

const x14Cf = (sqref, ...rules) =>
    '<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">' +
    `${rules.join("")}<xm:sqref>${sqref}</xm:sqref></x14:conditionalFormatting>`;

const extSection = (xml) => (xml.match(/<x14:conditionalFormattings>[\s\S]*<\/x14:conditionalFormattings>/) || [""])[0];

async function load(buffer) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    return wb;
}

// comparable models: writing assigns ref to the rules, and the x14 ids are regenerated
const snapshot = (conditionalFormattings) =>
    JSON.parse(JSON.stringify(conditionalFormattings.map((cf) => ({ ...cf, rules: cf.rules.map(({ x14Id, ...rule }) => rule) }))));

// a source sheet with the given main <conditionalFormatting>s and x14 extension
async function sourceWith(conditionalFormattings, ext) {
    return tamper(await baseWorkbookBuffer(), {
        [SHEET]: (xml) =>
            xml.replace("<pageMargins", `${conditionalFormattings}<pageMargins`).replace("</worksheet>", `${ext}</worksheet>`),
    });
}

describe("conditional formatting x14 extension", () => {
    describe("dataBar", () => {
        // as written by Excel: the main rule references its x14 counterpart by id
        const mainDataBar = (id, priority) =>
            `<cfRule type="dataBar" priority="${priority}"><dataBar><cfvo type="min"/><cfvo type="max"/>` +
            '<color rgb="FF638EC6"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" ' +
            `xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:id>${id}</x14:id></ext></extLst></cfRule>`;
        const ID1 = "{00000000-0000-0000-0000-0000000000D1}";
        const ID2 = "{00000000-0000-0000-0000-0000000000D2}";
        const MAIN = `<conditionalFormatting sqref="A1:A5">${mainDataBar(ID1, 1)}</conditionalFormatting><conditionalFormatting sqref="B1:B5">${mainDataBar(ID2, 2)}</conditionalFormatting>`;
        const EXT = extLst(
            x14Cf(
                "A1:A5",
                // gradient left to its default (true)
                `<x14:cfRule type="dataBar" id="${ID1}"><x14:dataBar minLength="0" maxLength="100" border="1" ` +
                    'negativeBarColorSameAsPositive="0" negativeBarBorderColorSameAsPositive="0" axisPosition="middle">' +
                    '<x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:borderColor rgb="FF638EC6"/>' +
                    '<x14:negativeFillColor rgb="FFFF0000"/><x14:negativeBorderColor rgb="FFFF0000"/>' +
                    '<x14:axisColor rgb="FF000000"/></x14:dataBar></x14:cfRule>'
            ),
            x14Cf(
                "B1:B5",
                `<x14:cfRule type="dataBar" id="${ID2}"><x14:dataBar minLength="10" maxLength="90" gradient="0" ` +
                    'direction="rightToLeft"><x14:cfvo type="num"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="max"/>' +
                    '<x14:negativeFillColor rgb="FF00FF00"/><x14:axisColor rgb="FF0000FF"/></x14:dataBar></x14:cfRule>'
            )
        );

        it("keeps its x14 settings when written back", async () => {
            const wb = await load(await sourceWith(MAIN, EXT));
            const [a, b] = wb.getWorksheet("Sheet1").conditionalFormattings;
            expect(a.rules[0]).to.include({ type: "dataBar", gradient: true, border: true, axisPosition: "middle" });
            expect(a.rules[0].cfvo).to.deep.equal([{ type: "autoMin" }, { type: "autoMax" }]);
            expect(a.rules[0].negativeFillColor).to.deep.equal({ argb: "FFFF0000" });
            expect(b.rules[0]).to.include({ gradient: false, minLength: 10, maxLength: 90, direction: "rightToLeft" });
            const before = snapshot(wb.getWorksheet("Sheet1").conditionalFormattings);

            const buffer = await wb.xlsx.writeBuffer();
            const xml = await readPart(buffer, SHEET);
            // main rules: Excel 2007 cfvo types, and a reference to their extension
            expect(xml).to.match(
                /<conditionalFormatting sqref="A1:A5"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"\/><cfvo type="max"\/><color rgb="FF638EC6"\/><\/dataBar><extLst><ext [^>]*><x14:id>\{[0-9A-F-]{36}\}<\/x14:id>/
            );
            expect(xml).to.include('<dataBar><cfvo type="num" val="0"/><cfvo type="max"/>');
            const ext = extSection(xml);
            expect(ext).to.include(
                '<x14:dataBar minLength="0" maxLength="100" border="1" negativeBarColorSameAsPositive="0" ' +
                    'negativeBarBorderColorSameAsPositive="0" axisPosition="middle"><x14:cfvo type="autoMin"/>' +
                    '<x14:cfvo type="autoMax"/><x14:borderColor rgb="FF638EC6"/><x14:negativeFillColor rgb="FFFF0000"/>' +
                    '<x14:negativeBorderColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/></x14:dataBar>'
            );
            expect(ext).to.include(
                '<x14:dataBar minLength="10" maxLength="90" gradient="0" direction="rightToLeft">' +
                    '<x14:cfvo type="num"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="max"/>' +
                    '<x14:negativeFillColor rgb="FF00FF00"/><x14:axisColor rgb="FF0000FF"/></x14:dataBar>'
            );

            const back = await load(buffer);
            expect(snapshot(back.getWorksheet("Sheet1").conditionalFormattings)).to.deep.equal(before);
        });

        it("created with x14 settings is written with its extension, even with a gradient", async () => {
            const wb = new ExcelJS.Workbook();
            const ws = wb.addWorksheet("Sheet1");
            ws.addConditionalFormatting({
                ref: "A1:A5",
                rules: [
                    { type: "dataBar", gradient: true, axisColor: { argb: "FF000000" }, cfvo: [{ type: "autoMin" }, { type: "autoMax" }] },
                ],
            });
            ws.addConditionalFormatting({ ref: "B1:B5", rules: [{ type: "dataBar", gradient: true }] });
            const xml = await readPart(await wb.xlsx.writeBuffer(), SHEET);
            // autoMin / autoMax only exist in the extension
            expect(xml).to.include('<dataBar><cfvo type="min"/><cfvo type="max"/>');
            expect(extSection(xml)).to.include('<x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:axisColor rgb="FF000000"/>');
            expect(extSection(xml)).not.to.include("B1:B5");
            // nor a reference to a missing extension rule
            expect(xml).to.include(
                '<conditionalFormatting sqref="B1:B5"><cfRule type="dataBar" priority="2"><dataBar><cfvo type="min"/>' +
                    '<cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule></conditionalFormatting>'
            );
            // autoMin alone is an x14 setting
            ws.addConditionalFormatting({
                ref: "C1:C5",
                rules: [{ type: "dataBar", gradient: true, cfvo: [{ type: "autoMin" }, { type: "max" }] }],
            });
            expect(extSection(await readPart(await wb.xlsx.writeBuffer(), SHEET))).to.include("<xm:sqref>C1:C5</xm:sqref>");
        });
    });
});
