const fs = require("fs");

const testXformHelper = require("../test-xform-helper");

const RelationshipsXform = verquire("xlsx/xform/core/relationships-xform");

const expectations = [
    {
        title: "worksheet.rels",
        create() {
            return new RelationshipsXform();
        },
        preparedModel: require("./data/worksheet.rels.1.json"),
        xml: fs.readFileSync(`${__dirname}/data/worksheet.rels.xml`).toString().replace(/\r\n/g, "\n"),
        get parsedModel() {
            return this.preparedModel;
        },
        tests: ["render", "renderIn", "parse"],
    },
];

describe("RelationshipsXform", () => {
    testXformHelper(expectations);

    const parseSax = verquire("utils/parse-sax");
    const parse = (xml) => new RelationshipsXform().parse(parseSax([`<Relationships>${xml}</Relationships>`]));

    it("skips a relationship without Id, Type or Target (all required)", async () => {
        const model = await parse(
            '<Relationship Type="t" Target="a.xml"/>' +
                '<Relationship Id="rId2" Target="b.xml"/>' +
                '<Relationship Id="rId3" Type="t"/>' +
                '<Relationship Id="rId4" Type="t" Target="d.xml"/>'
        );
        expect(model).to.deep.equal([{ Id: "rId4", Type: "t", Target: "d.xml" }]);
    });

    it("keeps a valid TargetMode and drops an invalid one", async () => {
        const model = await parse(
            '<Relationship Id="rId1" Type="t" Target="http://x/" TargetMode="External"/>' +
                '<Relationship Id="rId2" Type="t" Target="b.xml" TargetMode="Internal"/>' +
                '<Relationship Id="rId3" Type="t" Target="c.xml" TargetMode="external" Other="o"/>'
        );
        expect(model).to.deep.equal([
            { Id: "rId1", Type: "t", Target: "http://x/", TargetMode: "External" },
            { Id: "rId2", Type: "t", Target: "b.xml", TargetMode: "Internal" },
            { Id: "rId3", Type: "t", Target: "c.xml", Other: "o" },
        ]);
    });
});
