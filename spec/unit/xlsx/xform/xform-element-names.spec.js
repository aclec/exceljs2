const { Readable } = require("stream");

const parseSax = verquire("utils/parse-sax");
const CoreXform = verquire("xlsx/xform/core/core-xform");
const AppXform = verquire("xlsx/xform/core/app-xform");
const SharedStringXform = verquire("xlsx/xform/strings/shared-string-xform");
const PhoneticTextXform = verquire("xlsx/xform/strings/phonetic-text-xform");
const RichTextXform = verquire("xlsx/xform/strings/rich-text-xform");
const DatabarXform = verquire("xlsx/xform/sheet/cf/databar-xform");

const parse = (xform, xml) => xform.parse(parseSax(Readable.from([xml])));

// The child xforms are looked up by element name, which comes from the file: an element named after an Object.prototype
// member (constructor, toString, __proto__...) must be an unknown element, not the inherited member (which made the parse
// fail with "this.parser.parseOpen is not a function").
describe("xforms and element names of Object.prototype members", () => {
    const names = ["constructor", "toString", "hasOwnProperty", "__proto__", "valueOf"];

    it("CoreXform skips them", async () => {
        for (const name of names) {
            const model = await parse(new CoreXform(), `<cp:coreProperties><${name}>x</${name}><dc:title>T</dc:title></cp:coreProperties>`);
            expect(model.title, name).to.equal("T");
        }
    });

    it("AppXform skips them", async () => {
        for (const name of names) {
            const model = await parse(new AppXform(), `<Properties><${name}/><Company>C</Company></Properties>`);
            expect(model.company, name).to.equal("C");
        }
    });

    it("SharedStringXform and PhoneticTextXform skip them", async () => {
        for (const name of names) {
            expect(await parse(new SharedStringXform(), `<si><${name}/><t>a</t></si>`), name).to.equal("a");
            const phonetic = await parse(new PhoneticTextXform(), `<rPh sb="0" eb="1"><${name}/><t>b</t></rPh>`);
            expect(phonetic, name).to.deep.equal({ sb: 0, eb: 1, text: "b" });
        }
    });

    it("CompositeXform skips them", async () => {
        for (const name of names) {
            const model = await parse(new DatabarXform(), `<dataBar><${name}/><cfvo type="min"/><cfvo type="max"/></dataBar>`);
            expect(model.cfvo, name).to.deep.equal([
                { type: "min", value: undefined },
                { type: "max", value: undefined },
            ]);
        }
    });
});

describe("RichTextXform", () => {
    it("reads the text and the font of a run", async () => {
        const model = await parse(new RichTextXform(), '<r><rPr><b/><sz val="11"/></rPr><t xml:space="preserve"> x </t></r>');
        expect(model).to.deep.equal({ font: { bold: true, size: 11 }, text: " x " });
    });

    it("ignores a <t> misplaced inside <rPr> and an <rPr> misplaced inside <t>", async () => {
        expect(await parse(new RichTextXform(), "<r><rPr><t>x</t></rPr><t>y</t></r>")).to.deep.equal({ text: "y" });
        expect(await parse(new RichTextXform(), "<r><t>y<rPr/></t></r>")).to.deep.equal({});
    });
});
