const testXformHelper = require("../../test-xform-helper");

const XmlStream = verquire("utils/xml-stream");
const DatabarExtXform = verquire("xlsx/xform/sheet/cf-ext/databar-ext-xform");

const expectations = [
    {
        title: "Default Set",
        create() {
            return new DatabarExtXform();
        },
        preparedModel: {
            cfvo: [
                { type: "num", value: 5 },
                { type: "num", value: 20 },
            ],
        },
        xml: `
      <x14:dataBar minLength="0" maxLength="100">
        <x14:cfvo type="num"><xm:f>5</xm:f></x14:cfvo>
        <x14:cfvo type="num"><xm:f>20</xm:f></x14:cfvo>
      </x14:dataBar>
    `,
        parsedModel: {
            minLength: 0,
            maxLength: 100,
            border: false,
            gradient: true,
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
        title: "Non Default Set",
        create() {
            return new DatabarExtXform();
        },
        preparedModel: {
            minLength: 5,
            maxLength: 95,
            gradient: false,
            border: true,
            negativeBarColorSameAsPositive: false,
            negativeBarBorderColorSameAsPositive: false,
            axisPosition: "middle",
            borderColor: { argb: "FFFF0000" },
            negativeFillColor: { argb: "FF00FF00" },
            axisColor: { argb: "FF0000FF" },
            direction: "rightToLeft",
            cfvo: [{ type: "autoMin" }, { type: "autoMax" }],
        },
        xml: `
      <x14:dataBar
         minLength="5"
         maxLength="95"
         gradient="0"
         border="1"
         negativeBarColorSameAsPositive="0"
         negativeBarBorderColorSameAsPositive="0"
         axisPosition="middle"
         direction="rightToLeft"
       >
        <x14:cfvo type="autoMin"/>
        <x14:cfvo type="autoMax"/>
        <x14:borderColor rgb="FFFF0000"/>
        <x14:negativeFillColor rgb="FF00FF00"/>
        <x14:axisColor rgb="FF0000FF"/>
      </x14:dataBar>
    `,
        get parsedModel() {
            return this.preparedModel;
        },
        tests: ["render", "parse"],
    },
];

describe("DatabarExtXform", () => {
    testXformHelper(expectations);

    // chai-xml ignores sibling order, so check the CT_DataBar sequence on the raw string
    it("renders colours in CT_DataBar order", () => {
        const xmlStream = new XmlStream();
        new DatabarExtXform().render(xmlStream, {
            cfvo: [{ type: "autoMin" }, { type: "autoMax" }],
            axisColor: { argb: "FF000004" },
            negativeBorderColor: { argb: "FF000003" },
            negativeFillColor: { argb: "FF000002" },
            borderColor: { argb: "FF000001" },
        });
        const order = ["borderColor", "negativeFillColor", "negativeBorderColor", "axisColor"].map((name) =>
            xmlStream.xml.indexOf(`<x14:${name} `)
        );
        order.forEach((index) => expect(index).to.be.above(0));
        expect(order).to.deep.equal([...order].sort((a, b) => a - b));
    });
});
