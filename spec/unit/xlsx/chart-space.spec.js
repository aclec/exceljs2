// The chart part generated for a chart created by worksheet.addChart (src/xlsx/chart-space.ts): the order of the
// elements (dml-chart.xsd sequences: Excel refuses a chart out of order; chai-xml ignores sibling order, so the raw
// strings are checked), the caches, the encoding of the texts, and the anchor of the chart in the drawing. With
// OOXML_SCHEMAS set (spec/utils/xmllint.js), one chart of each type is validated against dml-chart.xsd.
const fs = require("fs");
const path = require("path");
const { schemaValidator } = require("../../utils/xmllint");

const { renderChartSpace, chartAnchorModel } = verquire("xlsx/chart-space");

const OUT = path.join(__dirname, "../../out/charts-xml");

const numbers = (formula, values, formatCode = "General") => ({
    formula,
    cache: { type: "number", formatCode, count: values.length, points: values.map((value, index) => ({ index, value })) },
});
const strings = (formula, values) => ({
    formula,
    cache: { type: "string", count: values.length, points: values.map((value, index) => ({ index, value })) },
});

const AXIS = { majorGridlines: false, visible: true };

// a chart of two series (the second without name, categories nor color)
function chart(overrides = {}) {
    return {
        type: "column",
        marker: true,
        line: false,
        legend: "r",
        date1904: false,
        series: [
            {
                name: strings("Sheet1!$B$1", ["Sales"]),
                categories: strings("Sheet1!$A$2:$A$3", ["Jan", "Feb"]),
                values: numbers("Sheet1!$B$2:$B$3", [1, 2]),
                color: "FF0000",
            },
            { values: numbers("Sheet1!$C$2:$C$3", [3, 4]) },
        ],
        categoryAxis: AXIS,
        valueAxis: { majorGridlines: true, visible: true },
        ...overrides,
    };
}

// asserts that the strings are found in the xml in this order (each after the previous one)
function expectOrder(xml, ...parts) {
    let from = 0;
    for (const part of parts) {
        const at = xml.indexOf(part, from);
        expect(at, `${part} after position ${from}`).to.not.equal(-1);
        from = at + part.length;
    }
}

const ALL_TYPES = [
    chart({ title: "Column" }),
    chart({ type: "bar", grouping: "stacked" }),
    chart({ type: "column", grouping: "percentStacked" }),
    chart({ type: "line", title: "Line" }),
    chart({ type: "line", marker: false, grouping: "stacked" }),
    chart({ type: "area", grouping: "stacked" }),
    chart({ type: "pie", series: chart().series.slice(0, 1) }),
    chart({ type: "doughnut" }),
    chart({ type: "scatter", series: [{ categories: numbers("Sheet1!$A$2:$A$3", [1, 2]), values: numbers("Sheet1!$B$2:$B$3", [3, 4]) }] }),
    chart({ type: "scatter", line: true, marker: false }),
    chart({
        type: "column",
        legend: undefined,
        categoryAxis: { title: "Month", numFmt: "mmm", majorGridlines: true, visible: false },
        valueAxis: { title: "Amount", min: -1, max: 1.5, numFmt: "0%", majorGridlines: false, visible: true },
    }),
];

describe("renderChartSpace", () => {
    it("writes a column chart in the order of the schema", () => {
        const xml = renderChartSpace(chart({ title: "Title" }));
        expect(xml.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c=')).to.be.true();
        expectOrder(
            xml,
            '<c:date1904 val="0"/>',
            '<c:roundedCorners val="0"/>',
            "<c:chart>",
            "<c:title>",
            '<c:autoTitleDeleted val="0"/>',
            "<c:plotArea>",
            "<c:layout/>",
            "<c:barChart>",
            '<c:barDir val="col"/>',
            '<c:grouping val="clustered"/>',
            '<c:varyColors val="0"/>',
            "<c:ser>",
            "<c:ser>",
            '<c:gapWidth val="150"/>',
            '<c:axId val="500000001"/>',
            '<c:axId val="500000002"/>',
            "</c:barChart>",
            "<c:catAx>",
            "<c:valAx>",
            "</c:plotArea>",
            '<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>',
            '<c:plotVisOnly val="1"/>',
            '<c:dispBlanksAs val="gap"/>',
            "</c:chart></c:chartSpace>"
        );
        expect(xml).to.not.include("c:overlap");
    });

    it("writes the elements of a bar series in order, with its caches", () => {
        const xml = renderChartSpace(chart());
        expectOrder(
            xml,
            '<c:ser><c:idx val="0"/><c:order val="0"/>',
            '<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Sales</c:v></c:pt></c:strCache></c:strRef></c:tx>',
            '<c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr>',
            '<c:invertIfNegative val="0"/>',
            '<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Jan</c:v></c:pt><c:pt idx="1"><c:v>Feb</c:v></c:pt></c:strCache></c:strRef></c:cat>',
            '<c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>',
            '<c:ser><c:idx val="1"/><c:order val="1"/><c:invertIfNegative val="0"/><c:val>'
        );
    });

    it("writes a sparse cache: the count of cells, a point for each value only", () => {
        const values = {
            formula: "S!$B$1:$B$5",
            cache: {
                type: "number",
                formatCode: "0.00",
                count: 5,
                points: [
                    { index: 1, value: -0 },
                    { index: 4, value: 1e-7 },
                ],
            },
        };
        const xml = renderChartSpace(chart({ series: [{ values, name: "Plain name" }] }));
        expect(xml).to.include(
            '<c:numCache><c:formatCode>0.00</c:formatCode><c:ptCount val="5"/><c:pt idx="1"><c:v>0</c:v></c:pt><c:pt idx="4"><c:v>1e-7</c:v></c:pt></c:numCache>'
        );
        expect(xml).to.include("<c:tx><c:v>Plain name</c:v></c:tx>");
    });

    it("overlaps the bars of a stacked bar chart", () => {
        const xml = renderChartSpace(chart({ type: "bar", grouping: "stacked" }));
        expectOrder(
            xml,
            '<c:barDir val="bar"/>',
            '<c:grouping val="stacked"/>',
            '<c:gapWidth val="150"/>',
            '<c:overlap val="100"/>',
            "<c:axId"
        );
        // horizontal: the categories on the left, the values at the bottom
        expectOrder(xml, "<c:catAx>", '<c:axPos val="l"/>', "<c:valAx>", '<c:axPos val="b"/>');
    });

    it("writes the markers and lines of a line chart in order", () => {
        const xml = renderChartSpace(chart({ type: "line" }));
        expectOrder(
            xml,
            "<c:lineChart>",
            '<c:grouping val="standard"/>',
            '<c:varyColors val="0"/>',
            "<c:ser>",
            '<c:spPr><a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr>',
            '<c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>',
            "<c:cat>",
            "<c:val>",
            '<c:smooth val="0"/></c:ser>',
            "<c:ser>",
            '<c:smooth val="0"/></c:ser>',
            '<c:marker val="1"/>',
            "<c:axId",
            "</c:lineChart>"
        );
        const noMarker = renderChartSpace(chart({ type: "line", marker: false }));
        expect(noMarker.split('<c:marker><c:symbol val="none"/></c:marker>').length).to.equal(3);
    });

    it("writes an area chart with its points on the categories", () => {
        const xml = renderChartSpace(chart({ type: "area", grouping: "percentStacked" }));
        expectOrder(xml, "<c:areaChart>", '<c:grouping val="percentStacked"/>', "<c:ser>", "<c:axId", "</c:areaChart>");
        expect(xml).to.include('<c:crossBetween val="midCat"/>');
        expect(xml).to.not.include("c:marker");
    });

    it("writes a pie and a doughnut without axes, one color per slice", () => {
        const pie = renderChartSpace(chart({ type: "pie" }));
        expectOrder(pie, "<c:pieChart>", '<c:varyColors val="1"/>', "<c:ser>", '<c:firstSliceAng val="0"/>', "</c:pieChart></c:plotArea>");
        expect(pie).to.not.include("c:axId");
        expect(pie).to.not.include("c:invertIfNegative");
        const doughnut = renderChartSpace(chart({ type: "doughnut" }));
        expectOrder(
            doughnut,
            "<c:doughnutChart>",
            '<c:varyColors val="1"/>',
            '<c:firstSliceAng val="0"/>',
            '<c:holeSize val="50"/>',
            "</c:doughnutChart></c:plotArea>"
        );
    });

    it("writes a scatter chart: x and y values, two value axes, no lines unless asked", () => {
        const xml = renderChartSpace(chart({ type: "scatter" }));
        expectOrder(
            xml,
            "<c:scatterChart>",
            '<c:scatterStyle val="lineMarker"/>',
            '<c:varyColors val="0"/>',
            '<c:ser><c:idx val="0"/><c:order val="0"/><c:tx>',
            "<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>",
            "<c:marker>",
            "<c:xVal><c:strRef>",
            "<c:yVal><c:numRef>",
            '<c:smooth val="0"/></c:ser>',
            "</c:scatterChart>",
            '<c:valAx><c:axId val="500000001"/>',
            '<c:axPos val="b"/>',
            '<c:crossBetween val="midCat"/></c:valAx>',
            '<c:valAx><c:axId val="500000002"/>'
        );
        expect(xml).to.not.include("c:catAx");
        expect(renderChartSpace(chart({ type: "scatter", line: true }))).to.include('<a:ln><a:solidFill><a:srgbClr val="FF0000"/>');
    });

    it("writes the options of the axes in the order of EG_AxShared, CT_Scaling (max before min)", () => {
        const xml = renderChartSpace(ALL_TYPES[ALL_TYPES.length - 1]);
        expectOrder(
            xml,
            '<c:catAx><c:axId val="500000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="b"/><c:majorGridlines/>',
            "<c:title><c:tx><c:rich><a:bodyPr/>",
            "<a:t>Month</a:t>",
            '<c:numFmt formatCode="mmm" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>',
            '<c:valAx><c:axId val="500000002"/><c:scaling><c:orientation val="minMax"/><c:max val="1.5"/><c:min val="-1"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:title>',
            // the title of a vertical axis is rotated
            '<a:bodyPr rot="-5400000" vert="horz"/>',
            '<c:numFmt formatCode="0%" sourceLinked="0"/>',
            '<c:crossBetween val="between"/></c:valAx>'
        );
        // no legend, no title: the title is not taken from the series name
        expect(xml).to.not.include("c:legend");
        expect(xml).to.include('<c:chart><c:autoTitleDeleted val="1"/>');
    });

    it("encodes the texts and formulas (characters invalid in XML removed)", () => {
        const xml = renderChartSpace(
            chart({
                title: 'R&D <2024> "q" \u0001ok',
                series: [
                    {
                        name: "a<b",
                        values: numbers("'It''s & <ok>'!$B$2:$B$3", [1, 2]),
                        categories: strings("S!$A$1:$A$2", ["x&y", "]]>"]),
                    },
                ],
            })
        );
        expect(xml).to.include("<a:t>R&amp;D &lt;2024&gt; &quot;q&quot; ok</a:t>");
        expect(xml).to.include("<c:v>a&lt;b</c:v>");
        expect(xml).to.include("<c:f>&apos;It&apos;&apos;s &amp; &lt;ok&gt;&apos;!$B$2:$B$3</c:f>");
        expect(xml).to.include("<c:v>x&amp;y</c:v>");
        expect(xml).to.include("<c:v>]]&gt;</c:v>");
    });

    it("writes the date system of the workbook", () => {
        expect(renderChartSpace(chart({ date1904: true }))).to.include('<c:date1904 val="1"/>');
    });

    describe("against dml-chart.xsd (OOXML_SCHEMAS)", () => {
        const validate = schemaValidator();
        before(function () {
            if (!validate) {
                this.skip();
            }
            fs.mkdirSync(OUT, { recursive: true });
        });

        it("validates a chart of each type and option", () => {
            const files = ALL_TYPES.map((spec, index) => {
                const file = path.join(OUT, `chart-${index + 1}-${spec.type}.xml`);
                fs.writeFileSync(file, renderChartSpace(spec));
                return file;
            });
            const result = validate("dml-chart.xsd", files);
            expect(result.ok, result.output).to.be.true();
        });

        it("rejects a chart whose elements are out of order (the validation is real)", () => {
            const file = path.join(OUT, "chart-out-of-order.xml");
            fs.writeFileSync(
                file,
                renderChartSpace(chart())
                    .replace('<c:roundedCorners val="0"/>', "")
                    .replace("<c:chart>", '<c:chart><c:roundedCorners val="0"/>')
            );
            expect(validate("dml-chart.xsd", [file]).ok).to.be.false();
        });
    });
});

describe("chartAnchorModel", () => {
    const from = { nativeCol: 4, nativeColOff: 10, nativeRow: 1, nativeRowOff: 20 };
    const to = { nativeCol: 12, nativeColOff: 0, nativeRow: 20, nativeRowOff: 0 };

    it("anchors the chart by two cells, its chart part as bytes", () => {
        const model = chartAnchorModel({ from, to, name: "Chart 3" }, "<xml/>");
        const { anchor, rels } = model;
        expect(anchor.name).to.equal("xdr:twoCellAnchor");
        expect(anchor.attributes).to.deep.equal([]);
        expect(anchor.children.map((child) => child.name)).to.deep.equal(["xdr:from", "xdr:to", "xdr:graphicFrame", "xdr:clientData"]);
        expect(anchor.children[0].children.map((child) => [child.name, child.children[0]])).to.deep.equal([
            ["xdr:col", "4"],
            ["xdr:colOff", "10"],
            ["xdr:row", "1"],
            ["xdr:rowOff", "20"],
        ]);
        const [nvPr, xfrm, graphic] = anchor.children[2].children;
        expect(nvPr.children[0]).to.deep.equal({
            name: "xdr:cNvPr",
            attributes: [
                { name: "id", value: "1", role: "shapeId" },
                { name: "name", value: "Chart 3" },
            ],
            children: [],
        });
        expect(xfrm.name).to.equal("xdr:xfrm");
        const chartElement = graphic.children[0].children[0];
        expect(chartElement.name).to.equal("c:chart");
        expect(chartElement.attributes[2]).to.deep.equal({ name: "r:id", value: "rId1", role: "rel" });
        expect(rels).to.have.length(1);
        expect(rels[0].Id).to.equal("rId1");
        expect(rels[0].Type).to.equal("http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart");
        expect(rels[0].part.kind).to.equal("chart");
        expect(rels[0].part.extension).to.equal("xml");
        expect(Buffer.from(rels[0].part.data).toString("utf8")).to.equal("<xml/>");
    });

    it("writes editAs on a two-cell anchor, the size in EMU on a one-cell anchor", () => {
        expect(chartAnchorModel({ from, to, editAs: "absolute", name: "C" }, "").anchor.attributes).to.deep.equal([
            { name: "editAs", value: "absolute" },
        ]);
        const { anchor } = chartAnchorModel({ from, ext: { width: 400.5, height: 300 }, name: "C" }, "");
        expect(anchor.name).to.equal("xdr:oneCellAnchor");
        expect(anchor.children.map((child) => child.name)).to.deep.equal(["xdr:from", "xdr:ext", "xdr:graphicFrame", "xdr:clientData"]);
        expect(anchor.children[1].attributes).to.deep.equal([
            { name: "cx", value: "3814762" },
            { name: "cy", value: "2857500" },
        ]);
    });
});
