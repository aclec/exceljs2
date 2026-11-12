// Conditional formats written and read back, in the main section and in the x14 extension (<extLst>)
// of the worksheet.
const { PassThrough } = require("stream");
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

const mainSection = (xml) => (xml.match(/<conditionalFormatting [\s\S]*<\/conditionalFormatting>/) || [""])[0];
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

describe("conditional formatting", () => {
    describe("dataBar x14 extension", () => {
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
                '<x14:dataBar minLength="0" maxLength="100" border="1" ' +
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

        it("keeps the settings Excel writes when they differ from the schema defaults", async () => {
            const main =
                '<conditionalFormatting sqref="C1:C5"><cfRule type="dataBar" priority="1"><dataBar showValue="0">' +
                '<cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst>' +
                '<ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
                `<x14:id>${ID1}</x14:id></ext></extLst></cfRule></conditionalFormatting>`;
            // [MS-XLSX] CT_DataBar defaults: minLength 10, maxLength 90, negativeBarColorSameAsPositive false,
            // direction context
            const ext = extLst(
                x14Cf(
                    "C1:C5",
                    `<x14:cfRule type="dataBar" id="${ID1}"><x14:dataBar negativeBarColorSameAsPositive="1" ` +
                        'direction="leftToRight"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/></x14:dataBar></x14:cfRule>'
                )
            );
            const wb = await load(await sourceWith(main, ext));
            const [rule] = wb.getWorksheet("Sheet1").conditionalFormattings[0].rules;
            expect(rule).to.include({
                showValue: false,
                minLength: 10,
                maxLength: 90,
                negativeBarColorSameAsPositive: true,
                direction: "leftToRight",
            });

            const xml = await readPart(await wb.xlsx.writeBuffer(), SHEET);
            expect(xml).to.include('<dataBar showValue="0"><cfvo type="min"/>');
            expect(extSection(xml)).to.include(
                '<x14:dataBar minLength="10" maxLength="90" negativeBarColorSameAsPositive="1" direction="leftToRight">'
            );
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

    describe("text rules", () => {
        const bold = { font: { bold: true } };

        it("are written with their type, operator, text and formula, and read back", async () => {
            const wb = new ExcelJS.Workbook();
            const ws = wb.addWorksheet("Sheet1");
            ws.addConditionalFormatting({
                ref: "A1:A5",
                rules: [
                    { type: "containsText", operator: "containsText", text: "ab", style: bold },
                    { type: "containsText", operator: "notContainsText", text: "ab", style: bold },
                    { type: "containsText", operator: "beginsWith", text: 'say "hi"', style: bold },
                    { type: "containsText", operator: "endsWith", text: "ab", style: bold },
                    { type: "containsText", operator: "containsBlanks", style: bold },
                    { type: "containsText", text: "x", style: bold },
                ],
            });
            const buffer = await wb.xlsx.writeBuffer();
            const xml = await readPart(buffer, SHEET);
            [
                '<cfRule type="containsText" dxfId="0" priority="1" operator="containsText" text="ab"><formula>NOT(ISERROR(SEARCH(&quot;ab&quot;,A1)))</formula></cfRule>',
                '<cfRule type="notContainsText" dxfId="1" priority="2" operator="notContains" text="ab"><formula>ISERROR(SEARCH(&quot;ab&quot;,A1))</formula></cfRule>',
                '<cfRule type="beginsWith" dxfId="2" priority="3" operator="beginsWith" text="say &quot;hi&quot;"><formula>LEFT(A1,LEN(&quot;say &quot;&quot;hi&quot;&quot;&quot;))=&quot;say &quot;&quot;hi&quot;&quot;&quot;</formula></cfRule>',
                '<cfRule type="endsWith" dxfId="3" priority="4" operator="endsWith" text="ab"><formula>RIGHT(A1,LEN(&quot;ab&quot;))=&quot;ab&quot;</formula></cfRule>',
                // containsBlanks is not an ST_ConditionalFormattingOperator
                '<cfRule type="containsBlanks" dxfId="4" priority="5"><formula>LEN(TRIM(A1))=0</formula></cfRule>',
                '<cfRule type="containsText" dxfId="5" priority="6" operator="containsText" text="x"><formula>NOT(ISERROR(SEARCH(&quot;x&quot;,A1)))</formula></cfRule>',
            ].forEach((rule) => expect(xml).to.include(rule));

            const [cf] = (await load(buffer)).getWorksheet("Sheet1").conditionalFormattings;
            expect(cf.rules.map(({ type, operator, text }) => [type, operator, text])).to.deep.equal([
                ["containsText", "containsText", "ab"],
                ["containsText", "notContainsText", "ab"],
                ["containsText", "beginsWith", 'say "hi"'],
                ["containsText", "endsWith", "ab"],
                ["containsText", "containsBlanks", undefined],
                ["containsText", "containsText", "x"],
            ]);
            expect(cf.rules[2].formulae).to.deep.equal(['LEFT(A1,LEN("say ""hi"""))="say ""hi"""']);
        });
    });

    describe("x14 text and time period rules", () => {
        const DXF = "<x14:dxf><font><b/></font></x14:dxf>";
        const id = (n) => `{00000000-0000-0000-0000-0000000000E${n}}`;
        // as Excel writes a text rule comparing with a cell of another sheet: the text is a second formula
        const textRule = (type, operator, priority, formula) =>
            `<x14:cfRule type="${type}" priority="${priority}" operator="${operator}" id="${id(priority)}">` +
            `<xm:f>${formula}</xm:f><xm:f>Other!$A$1</xm:f>${DXF}</x14:cfRule>`;
        const EXT = extLst(
            x14Cf(
                "A1:A5",
                textRule("containsText", "containsText", 1, "NOT(ISERROR(SEARCH(Other!$A$1,A1)))"),
                textRule("beginsWith", "beginsWith", 2, "LEFT(A1,LEN(Other!$A$1))=Other!$A$1"),
                textRule("endsWith", "endsWith", 3, "RIGHT(A1,LEN(Other!$A$1))=Other!$A$1"),
                textRule("notContainsText", "notContains", 4, "ISERROR(SEARCH(Other!$A$1,A1))"),
                `<x14:cfRule type="timePeriod" priority="5" timePeriod="today" id="${id(5)}"><xm:f>FLOOR(A1,1)=TODAY()</xm:f>${DXF}</x14:cfRule>`,
                `<x14:cfRule type="containsBlanks" priority="6" id="${id(6)}"><xm:f>LEN(TRIM(A1))=0</xm:f>${DXF}</x14:cfRule>`
            )
        );

        it("found only in the extension keep their type and are written back there", async () => {
            const wb = await load(await sourceWith("", EXT));
            const [cf] = wb.getWorksheet("Sheet1").conditionalFormattings;
            expect(cf.rules.map(({ type, operator, timePeriod, x14 }) => [type, operator || timePeriod, x14])).to.deep.equal([
                ["containsText", "containsText", true],
                ["containsText", "beginsWith", true],
                ["containsText", "endsWith", true],
                ["containsText", "notContainsText", true],
                ["timePeriod", "today", true],
                ["containsText", "containsBlanks", true],
            ]);
            expect(cf.rules[1].formulae).to.deep.equal(["LEFT(A1,LEN(Other!$A$1))=Other!$A$1", "Other!$A$1"]);
            cf.rules.forEach((rule) => expect(rule.style.font.bold).to.equal(true));
            const before = snapshot(wb.getWorksheet("Sheet1").conditionalFormattings);

            const buffer = await wb.xlsx.writeBuffer();
            const xml = await readPart(buffer, SHEET);
            expect(xml).not.to.include("<conditionalFormatting");
            expect(xml).not.to.include("<cfRule");
            const ext = extSection(xml);
            const written = (start, formulae) =>
                new RegExp(`${start} id="\\{[0-9A-F-]{36}\\}">${formulae}<x14:dxf><font><b/></font></x14:dxf></x14:cfRule>`);
            const text = (formula) => `<xm:f>${formula.replace(/[$()]/g, "\\$&")}</xm:f><xm:f>Other!\\$A\\$1</xm:f>`;
            expect(ext).to.match(
                written('<x14:cfRule type="containsText" priority="1" operator="containsText"', text("NOT(ISERROR(SEARCH(Other!$A$1,A1)))"))
            );
            expect(ext).to.match(
                written('<x14:cfRule type="beginsWith" priority="2" operator="beginsWith"', text("LEFT(A1,LEN(Other!$A$1))=Other!$A$1"))
            );
            expect(ext).to.match(
                written('<x14:cfRule type="endsWith" priority="3" operator="endsWith"', text("RIGHT(A1,LEN(Other!$A$1))=Other!$A$1"))
            );
            expect(ext).to.match(
                written('<x14:cfRule type="notContainsText" priority="4" operator="notContains"', text("ISERROR(SEARCH(Other!$A$1,A1))"))
            );
            expect(ext).to.match(
                written('<x14:cfRule type="timePeriod" priority="5" timePeriod="today"', "<xm:f>FLOOR\\(A1,1\\)=TODAY\\(\\)</xm:f>")
            );
            expect(ext).to.match(written('<x14:cfRule type="containsBlanks" priority="6"', "<xm:f>LEN\\(TRIM\\(A1\\)\\)=0</xm:f>"));

            expect(snapshot((await load(buffer)).getWorksheet("Sheet1").conditionalFormattings)).to.deep.equal(before);
        });

        it("created with a formula referencing another sheet are written in the extension", async () => {
            const build = (wb) => {
                const ws = wb.addWorksheet("Sheet1");
                ws.addConditionalFormatting({
                    ref: "A1:A5",
                    rules: [
                        {
                            type: "containsText",
                            operator: "beginsWith",
                            formulae: ["LEFT(A1,LEN(Other!$A$1))=Other!$A$1", "Other!$A$1"],
                            style: { font: { bold: true } },
                        },
                        { type: "containsText", operator: "endsWith", text: "z!", style: { font: { bold: true } } },
                    ],
                });
                return ws;
            };
            const check = (xml) => {
                expect(mainSection(xml)).not.to.include("Other!");
                expect(mainSection(xml)).to.include('<cfRule type="endsWith" dxfId="1" priority="2" operator="endsWith" text="z!">');
                expect(extSection(xml)).to.match(
                    /<x14:cfRule type="beginsWith" priority="1" operator="beginsWith" id="\{[0-9A-F-]{36}\}"><xm:f>LEFT\(A1,LEN\(Other!\$A\$1\)\)=Other!\$A\$1<\/xm:f><xm:f>Other!\$A\$1<\/xm:f><x14:dxf><font><b\/><\/font><\/x14:dxf><\/x14:cfRule><xm:sqref>A1:A5<\/xm:sqref>/
                );
            };

            const wb = new ExcelJS.Workbook();
            build(wb);
            check(await readPart(await wb.xlsx.writeBuffer(), SHEET));

            const stream = new PassThrough();
            const chunks = [];
            stream.on("data", (chunk) => chunks.push(chunk));
            const streamed = new ExcelJS.stream.xlsx.WorkbookWriter({ stream, useStyles: true });
            build(streamed).commit();
            await streamed.commit();
            check(await readPart(Buffer.concat(chunks), SHEET));
        });
    });
});
