const testXformHelper = require("../test-xform-helper");

const IntegerXform = verquire("xlsx/xform/simple/integer-xform");
const XmlStream = verquire("utils/xml-stream");

const expectations = [
    {
        title: "five",
        create() {
            return new IntegerXform({ tag: "integer", attr: "val" });
        },
        preparedModel: 5,
        xml: '<integer val="5"/>',
        parsedModel: 5,
        tests: ["render", "renderIn", "parse"],
    },
    {
        title: "zero",
        create() {
            return new IntegerXform({ tag: "integer", attr: "val" });
        },
        preparedModel: 0,
        xml: "",
        tests: ["render", "renderIn"],
    },
    {
        title: "undefined",
        create() {
            return new IntegerXform({ tag: "integer", attr: "val" });
        },
        preparedModel: undefined,
        xml: "",
        tests: ["render", "renderIn"],
    },
];

describe("IntegerXform", () => {
    testXformHelper(expectations);

    it("throws a clear Error for no value in text mode with the zero option", () => {
        const xform = new IntegerXform({ tag: "xdr:col", zero: true });
        const xmlStream = new XmlStream();
        expect(() => xform.render(xmlStream, undefined)).to.throw(Error, "No integer to write in <xdr:col>");
    });
});
