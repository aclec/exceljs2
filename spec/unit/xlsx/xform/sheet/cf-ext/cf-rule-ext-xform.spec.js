const testXformHelper = require("../../test-xform-helper");

const XmlStream = verquire("utils/xml-stream");
const CfRuleExtXform = verquire("xlsx/xform/sheet/cf-ext/cf-rule-ext-xform");

const expectations = [
    {
        title: "Icon Set",
        create() {
            return new CfRuleExtXform();
        },
        preparedModel: {
            type: "iconSet",
            priority: 1,
            x14Id: "x14-id",
            iconSet: "3Triangles",
            cfvo: [
                { type: "percent", value: 0 },
                { type: "percent", value: 33 },
                { type: "percent", value: 67 },
            ],
        },
        xml: `
      <x14:cfRule type="iconSet" priority="1" id="x14-id">
        <x14:iconSet iconSet="3Triangles">
          <x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo>
          <x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo>
          <x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo>
        </x14:iconSet>
      </x14:cfRule>
    `,
        parsedModel: {
            type: "iconSet",
            priority: 1,
            x14Id: "x14-id",
            iconSet: "3Triangles",
            reverse: false,
            showValue: true,
            cfvo: [
                { type: "percent", value: 0 },
                { type: "percent", value: 33 },
                { type: "percent", value: 67 },
            ],
        },
        tests: ["render", "parse"],
    },
    {
        title: "Databar",
        create() {
            return new CfRuleExtXform();
        },
        preparedModel: {
            type: "dataBar",
            x14Id: "x14-id",
            cfvo: [
                { type: "num", value: 5 },
                { type: "num", value: 20 },
            ],
        },
        xml: `
      <x14:cfRule type="dataBar" id="x14-id">
        <x14:dataBar minLength="0" maxLength="100">
          <x14:cfvo type="num"><xm:f>5</xm:f></x14:cfvo>
          <x14:cfvo type="num"><xm:f>20</xm:f></x14:cfvo>
        </x14:dataBar>
      </x14:cfRule>
    `,
        parsedModel: {
            type: "dataBar",
            x14Id: "x14-id",
            minLength: 0,
            maxLength: 100,
            border: false,
            gradient: true,
            // reverse: false,
            // showValue: true,
            negativeBarColorSameAsPositive: true,
            negativeBarBorderColorSameAsPositive: true,
            axisPosition: "auto",
            direction: "leftToRight",
            cfvo: [
                { type: "num", value: 5 },
                { type: "num", value: 20 },
            ],
        },
        tests: ["render", "parse"],
    },
    {
        title: "Expression (x14, formula referencing another sheet)",
        create() {
            return new CfRuleExtXform();
        },
        preparedModel: {
            type: "expression",
            priority: 2,
            x14Id: "x14-id",
            formulae: ["Other!$A$1=1"],
            style: { font: { bold: true }, fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFFFF00" } } },
        },
        xml: `
      <x14:cfRule type="expression" priority="2" id="x14-id">
        <xm:f>Other!$A$1=1</xm:f>
        <x14:dxf>
          <font><b/></font>
          <fill><patternFill patternType="solid"><bgColor rgb="FFFFFF00"/></patternFill></fill>
        </x14:dxf>
      </x14:cfRule>
    `,
        parsedModel: {
            type: "expression",
            priority: 2,
            x14Id: "x14-id",
            formulae: ["Other!$A$1=1"],
            // DxfXform reports every part of the format
            style: {
                font: { bold: true },
                fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFFFF00" } },
                alignment: null,
                border: null,
                numFmt: null,
                protection: null,
            },
        },
        tests: ["render", "parse"],
    },
    {
        title: "Cell Is (x14)",
        create() {
            return new CfRuleExtXform();
        },
        preparedModel: {
            type: "cellIs",
            priority: 3,
            stopIfTrue: true,
            operator: "between",
            x14: true,
            x14Id: "x14-id",
            formulae: ["'My Sheet'!$A$1", "'My Sheet'!$A$2"],
        },
        xml: `
      <x14:cfRule type="cellIs" priority="3" stopIfTrue="1" operator="between" id="x14-id">
        <xm:f>'My Sheet'!$A$1</xm:f>
        <xm:f>'My Sheet'!$A$2</xm:f>
      </x14:cfRule>
    `,
        parsedModel: {
            type: "cellIs",
            priority: 3,
            stopIfTrue: true,
            operator: "between",
            x14Id: "x14-id",
            formulae: ["'My Sheet'!$A$1", "'My Sheet'!$A$2"],
        },
        tests: ["render", "parse"],
    },
];

const render = (model) => {
    const xmlStream = new XmlStream();
    new CfRuleExtXform().render(xmlStream, model);
    return xmlStream.xml;
};

describe("CfRuleExtXform", () => {
    testXformHelper(expectations);

    describe("formula rules (expression, cellIs)", () => {
        it("are extension rules when flagged x14 or when a formula references another sheet", () => {
            const isExt = (type, formulae, x14) => CfRuleExtXform.isExt({ type, formulae, x14 });
            expect(isExt("expression", ["A1=1"])).to.equal(false);
            expect(isExt("expression", ["A1=1"], true)).to.equal(true);
            expect(isExt("cellIs", ["5"], true)).to.equal(true);
            expect(isExt("expression", ["Other!A1=1"])).to.equal(true);
            expect(isExt("expression", ["'Other sheet'!A1=1"])).to.equal(true);
            expect(isExt("expression", ["'It''s!'!A1=1"])).to.equal(true);
            expect(isExt("cellIs", ["Sheet1:Sheet3!A1"])).to.equal(true);
            expect(isExt("cellIs", ["[1]Data!A1"])).to.equal(true);
            // a "!" in a string literal is not a sheet reference
            expect(isExt("expression", ['A1="Hi!"'])).to.equal(false);
            expect(isExt("expression", ['A1="say ""Hi!"""'])).to.equal(false);
            // nor is the "!" of an error literal
            expect(isExt("expression", ["ISERROR(#REF!)", "A1=#DIV/0!"])).to.equal(false);
            expect(isExt("expression", ["AND(A1<>#NULL!,Other!A1)"])).to.equal(true);
            expect(isExt("cellIs", [5])).to.equal(false);
            expect(isExt("cellIs", undefined)).to.equal(false);
            // other rule types keep their own rules
            expect(isExt("containsText", ["Other!A1"], true)).to.equal(false);
            expect(isExt("top10", undefined, true)).to.equal(false);
        });

        it("render type, priority, stopIfTrue, operator and id in schema order, then the formulae and the inline dxf", () => {
            expect(
                render({
                    type: "cellIs",
                    priority: 4,
                    operator: "equal",
                    x14Id: "{ID}",
                    formulae: ["Other!$B$1"],
                    style: { font: { color: { argb: "FFFF0000" } } },
                })
            ).to.equal(
                '<x14:cfRule type="cellIs" priority="4" operator="equal" id="{ID}">' +
                    "<xm:f>Other!$B$1</xm:f>" +
                    '<x14:dxf><font><color rgb="FFFF0000"/></font></x14:dxf>' +
                    "</x14:cfRule>"
            );
        });

        it("render only the first formula of an expression and escape it", () => {
            expect(render({ type: "expression", priority: 1, x14Id: "{ID}", formulae: ['Other!A1<>"a&b"', "ignored"] })).to.equal(
                '<x14:cfRule type="expression" priority="1" id="{ID}"><xm:f>Other!A1&lt;&gt;&quot;a&amp;b&quot;</xm:f></x14:cfRule>'
            );
        });
    });
});
