const XmlStream = verquire("utils/xml-stream");

// leafNode() writes the element in one fragment: it must give the same document as openNode + writeText + closeNode
describe("XmlStream leafNode", () => {
    const longhand = (xmlStream, name, attributes, text) => {
        xmlStream.openNode(name, attributes);
        if (text !== undefined) xmlStream.writeText(text);
        xmlStream.closeNode();
    };
    const cases = [
        ["v", undefined, undefined],
        ["v", undefined, 0],
        ["v", undefined, ""],
        ["v", undefined, false],
        ["v", { a: 1, b: undefined, c: "x<y" }, "a & b"],
        ["v", {}, undefined],
        ["x:v", { "xml:space": "preserve" }, " text "],
    ];

    for (const [name, attributes, text] of cases) {
        it(`${name} ${JSON.stringify(attributes)} ${JSON.stringify(text)}`, () => {
            const build = (leaf) => {
                const xmlStream = new XmlStream();
                // first child of an open element (whose start tag is still open), then after a sibling, then at top level
                xmlStream.openNode("root", { r: 1 });
                leaf(xmlStream);
                leaf(xmlStream);
                xmlStream.addRollback();
                leaf(xmlStream);
                xmlStream.rollback();
                xmlStream.openNode("inner");
                xmlStream.writeText("t");
                leaf(xmlStream);
                xmlStream.closeNode();
                xmlStream.closeNode();
                leaf(xmlStream);
                return xmlStream.xml;
            };
            const expected = build((xmlStream) => longhand(xmlStream, name, attributes, text));
            expect(build((xmlStream) => xmlStream.leafNode(name, attributes, text))).to.equal(expected);
        });
    }

    it("leaves the parent open for attributes only until it is written", () => {
        const xmlStream = new XmlStream();
        xmlStream.openNode("root");
        xmlStream.leafNode("a");
        expect(() => xmlStream.addAttribute("x", 1)).to.throw("Cannot write attributes to node if it is not open");
        expect(xmlStream.xml).to.equal("<root><a/></root>");
    });
});
