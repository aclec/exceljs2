// The chart part of a chart created by `worksheet.addChart` (src/doc/chart.ts): a DrawingML `<c:chartSpace>`
// (ISO/IEC 29500-1 21.2, dml-chart.xsd), generated from a validated ChartSpaceSpec whose data references were resolved,
// with their cached values, when the workbook is written. Pure functions (no worksheet, no Node API): the streaming
// writer can reuse them.
//
// The chart is then written as the charts read from a file are (chart-parts.ts): renderChartSpace gives the bytes of the
// chart part, chartAnchorModel its anchor in the drawing of the worksheet; ChartPartWriter names and writes the part,
// prepareChartAnchors adds its relationship to the drawing, DrawingXform numbers the shape.
//
// Element order follows the sequences of dml-chart.xsd (CT_ChartSpace, CT_Chart, CT_PlotArea, CT_BarChart, CT_BarSer,
// EG_AxShared...): Excel refuses a chart whose elements are out of order. Every text and attribute goes through XmlStream
// (encoded, characters invalid in XML removed).

import XmlStream from "../utils/xml-stream";
import type {
    ChartAbsoluteAnchorSpec,
    ChartAnchorModel,
    ChartAnchorSpec,
    ChartAxisSpec,
    ChartDataRef,
    ChartNumberCache,
    ChartPartModel,
    ChartSeriesSpec,
    ChartSpaceSpec,
    ChartStringCache,
    ChartXmlAttribute,
    ChartXmlElement,
    AnchorModel,
} from "../types";

const NS_CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const NS_MAIN = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";

/** The ids of the two axes of a chart (unique within the chart). */
const CATEGORY_AXIS_ID = 500000001;
const VALUE_AXIS_ID = 500000002;

/** https://en.wikipedia.org/wiki/Office_Open_XML_file_formats#DrawingML */
export const EMU_PER_PIXEL = 9525;
/** The largest coordinate of DrawingML, in EMU (ST_Coordinate, ST_PositiveCoordinate: 27273042316900). */
export const MAX_COORDINATE = 27273042316900;

/** Whether a chart type is drawn without axes. */
function isPie(spec: ChartSpaceSpec): boolean {
    return spec.type === "pie" || spec.type === "doughnut";
}

/** `<c:{name} val="{value}"/>` */
function val(xml: XmlStream, name: string, value: string | number): void {
    xml.leafNode(name, { val: value });
}

/** `<a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill>` */
function solidFill(xml: XmlStream, color: string): void {
    xml.openNode("a:solidFill");
    xml.leafNode("a:srgbClr", { val: color });
    xml.closeNode();
}

/** A rich text (CT_TextBody) of one paragraph: a title. `vertical`: rotated a quarter turn (the title of a vertical axis). */
function richText(xml: XmlStream, text: string, vertical: boolean): void {
    xml.openNode("c:rich");
    xml.leafNode("a:bodyPr", vertical ? { rot: -5400000, vert: "horz" } : undefined);
    xml.leafNode("a:lstStyle");
    xml.openNode("a:p");
    xml.openNode("a:r");
    xml.leafNode("a:t", undefined, text);
    xml.closeNode();
    xml.closeNode();
    xml.closeNode();
}

/** `<c:title>` (CT_Title: tx, layout, overlay). */
function title(xml: XmlStream, text: string, vertical = false): void {
    xml.openNode("c:title");
    xml.openNode("c:tx");
    richText(xml, text, vertical);
    xml.closeNode();
    val(xml, "c:overlay", 0);
    xml.closeNode();
}

/** The text of a number in a cache (`<c:v>`). */
function numberText(value: number): string {
    // -0 is written 0
    return String(value === 0 ? 0 : value);
}

/**
 * `<c:numRef>`: the formula and the cached numbers (CT_NumRef: f, numCache?; CT_NumData: formatCode, ptCount, pt); no
 * `<c:numCache>` when the values are not known (optional in the schema).
 */
function numRef(xml: XmlStream, ref: ChartDataRef<ChartNumberCache>): void {
    xml.openNode("c:numRef");
    xml.leafNode("c:f", undefined, ref.formula);
    const { cache } = ref;
    if (cache) {
        xml.openNode("c:numCache");
        xml.leafNode("c:formatCode", undefined, cache.formatCode);
        val(xml, "c:ptCount", cache.count);
        for (const point of cache.points) {
            xml.openNode("c:pt", { idx: point.index });
            xml.leafNode("c:v", undefined, numberText(point.value));
            xml.closeNode();
        }
        xml.closeNode();
    }
    xml.closeNode();
}

/**
 * `<c:strRef>`: the formula and the cached texts (CT_StrRef: f, strCache?; CT_StrData: ptCount, pt); no `<c:strCache>`
 * when the values are not known.
 */
function strRef(xml: XmlStream, ref: ChartDataRef<ChartStringCache>): void {
    xml.openNode("c:strRef");
    xml.leafNode("c:f", undefined, ref.formula);
    const { cache } = ref;
    if (cache) {
        xml.openNode("c:strCache");
        val(xml, "c:ptCount", cache.count);
        for (const point of cache.points) {
            xml.openNode("c:pt", { idx: point.index });
            xml.leafNode("c:v", undefined, point.value);
            xml.closeNode();
        }
        xml.closeNode();
    }
    xml.closeNode();
}

/**
 * A data source of categories or x values (CT_AxDataSource): numbers or texts, as cached. Without a cache: numbers for
 * the x values of a scatter chart, texts for categories.
 */
function axisData(xml: XmlStream, spec: ChartSpaceSpec, name: string, ref: ChartDataRef<ChartNumberCache | ChartStringCache>): void {
    xml.openNode(name);
    const { cache } = ref;
    if (cache ? cache.type === "number" : spec.type === "scatter") {
        numRef(xml, { formula: ref.formula, cache: cache?.type === "number" ? cache : undefined });
    } else {
        strRef(xml, { formula: ref.formula, cache: cache?.type === "string" ? cache : undefined });
    }
    xml.closeNode();
}

/** The values of a series (CT_NumDataSource). */
function valueData(xml: XmlStream, name: string, ref: ChartDataRef<ChartNumberCache>): void {
    xml.openNode(name);
    numRef(xml, ref);
    xml.closeNode();
}

/** The shape properties of a series (`<c:spPr>`): its color as a fill, or as a line; scatter: no line unless asked. */
function seriesShape(xml: XmlStream, spec: ChartSpaceSpec, series: ChartSeriesSpec): void {
    const lineSeries = spec.type === "line" || spec.type === "scatter";
    const noLine = spec.type === "scatter" && !spec.line;
    if (series.color === undefined && !noLine) {
        return;
    }
    xml.openNode("c:spPr");
    if (!lineSeries && series.color !== undefined) {
        solidFill(xml, series.color);
    }
    if (lineSeries) {
        xml.openNode("a:ln");
        if (noLine) {
            xml.leafNode("a:noFill");
        } else if (series.color !== undefined) {
            solidFill(xml, series.color);
        }
        xml.closeNode();
    }
    xml.closeNode();
}

/** The marker of a line or scatter series (CT_Marker: symbol, size, spPr): none, the series color, or automatic. */
function seriesMarker(xml: XmlStream, spec: ChartSpaceSpec, series: ChartSeriesSpec): void {
    if (!spec.marker) {
        xml.openNode("c:marker");
        val(xml, "c:symbol", "none");
        xml.closeNode();
    } else if (series.color !== undefined) {
        xml.openNode("c:marker");
        val(xml, "c:symbol", "circle");
        val(xml, "c:size", 5);
        xml.openNode("c:spPr");
        solidFill(xml, series.color);
        xml.openNode("a:ln");
        solidFill(xml, series.color);
        xml.closeNode();
        xml.closeNode();
        xml.closeNode();
    }
}

/** A series, `<c:ser>`: EG_SerShared (idx, order, tx, spPr), then the elements of its type in schema order. */
function series(xml: XmlStream, spec: ChartSpaceSpec, item: ChartSeriesSpec, index: number): void {
    xml.openNode("c:ser");
    val(xml, "c:idx", index);
    val(xml, "c:order", index);
    if (item.name !== undefined) {
        xml.openNode("c:tx");
        if (typeof item.name === "string") {
            xml.leafNode("c:v", undefined, item.name);
        } else {
            strRef(xml, item.name);
        }
        xml.closeNode();
    }
    seriesShape(xml, spec, item);
    switch (spec.type) {
        case "bar":
        case "column":
            // CT_BarSer: invertIfNegative, ..., cat, val
            val(xml, "c:invertIfNegative", 0);
            break;
        case "line":
        case "scatter":
            // CT_LineSer / CT_ScatterSer: marker, ..., cat / xVal, val / yVal, smooth
            seriesMarker(xml, spec, item);
            break;
        case "pie":
        case "doughnut":
        case "area":
            break;
    }
    if (spec.type === "scatter") {
        if (item.categories) {
            axisData(xml, spec, "c:xVal", item.categories);
        }
        valueData(xml, "c:yVal", item.values);
    } else {
        if (item.categories) {
            axisData(xml, spec, "c:cat", item.categories);
        }
        valueData(xml, "c:val", item.values);
    }
    if (spec.type === "line" || spec.type === "scatter") {
        val(xml, "c:smooth", 0);
    }
    xml.closeNode();
}

/** The axis ids of a chart with axes (`<c:axId>` twice). */
function axisIds(xml: XmlStream): void {
    val(xml, "c:axId", CATEGORY_AXIS_ID);
    val(xml, "c:axId", VALUE_AXIS_ID);
}

/** The chart of the plot area, e.g. `<c:barChart>`, with its series. */
function plot(xml: XmlStream, spec: ChartSpaceSpec): void {
    const eachSeries = (): void => spec.series.forEach((item, index) => series(xml, spec, item, index));
    switch (spec.type) {
        case "bar":
        case "column": {
            // CT_BarChart: barDir, grouping, varyColors, ser, dLbls, gapWidth, overlap, serLines, axId
            const grouping = spec.grouping ?? "clustered";
            xml.openNode("c:barChart");
            val(xml, "c:barDir", spec.type === "bar" ? "bar" : "col");
            val(xml, "c:grouping", grouping);
            val(xml, "c:varyColors", 0);
            eachSeries();
            val(xml, "c:gapWidth", 150);
            if (grouping === "stacked" || grouping === "percentStacked") {
                // the bars of a stack are drawn over each other
                val(xml, "c:overlap", 100);
            }
            axisIds(xml);
            xml.closeNode();
            break;
        }
        case "line":
            // CT_LineChart: grouping, varyColors, ser, dLbls, dropLines, hiLowLines, upDownBars, marker, smooth, axId
            xml.openNode("c:lineChart");
            val(xml, "c:grouping", spec.grouping ?? "standard");
            val(xml, "c:varyColors", 0);
            eachSeries();
            val(xml, "c:marker", 1);
            axisIds(xml);
            xml.closeNode();
            break;
        case "area":
            // CT_AreaChart: grouping, varyColors, ser, dLbls, dropLines, axId
            xml.openNode("c:areaChart");
            val(xml, "c:grouping", spec.grouping ?? "standard");
            val(xml, "c:varyColors", 0);
            eachSeries();
            axisIds(xml);
            xml.closeNode();
            break;
        case "scatter":
            // CT_ScatterChart: scatterStyle, varyColors, ser, dLbls, axId (the lines and markers are set per series)
            xml.openNode("c:scatterChart");
            val(xml, "c:scatterStyle", "lineMarker");
            val(xml, "c:varyColors", 0);
            eachSeries();
            axisIds(xml);
            xml.closeNode();
            break;
        case "pie":
            // CT_PieChart: varyColors, ser, dLbls, firstSliceAng
            xml.openNode("c:pieChart");
            val(xml, "c:varyColors", 1);
            eachSeries();
            val(xml, "c:firstSliceAng", 0);
            xml.closeNode();
            break;
        case "doughnut":
            // CT_DoughnutChart: varyColors, ser, dLbls, firstSliceAng, holeSize
            xml.openNode("c:doughnutChart");
            val(xml, "c:varyColors", 1);
            eachSeries();
            val(xml, "c:firstSliceAng", 0);
            val(xml, "c:holeSize", 50);
            xml.closeNode();
            break;
    }
}

/** What differs between the axes of a chart. */
interface AxisLayout {
    readonly element: "c:catAx" | "c:valAx";
    readonly id: number;
    readonly crossAx: number;
    readonly position: "b" | "l";
}

/** An axis (EG_AxShared in schema order, then the elements of a category or a value axis). */
function axis(xml: XmlStream, spec: ChartSpaceSpec, layout: AxisLayout, options: ChartAxisSpec): void {
    xml.openNode(layout.element);
    val(xml, "c:axId", layout.id);
    // CT_Scaling: logBase, orientation, max, min
    xml.openNode("c:scaling");
    val(xml, "c:orientation", "minMax");
    if (options.max !== undefined) {
        val(xml, "c:max", options.max);
    }
    if (options.min !== undefined) {
        val(xml, "c:min", options.min);
    }
    xml.closeNode();
    val(xml, "c:delete", options.visible ? 0 : 1);
    val(xml, "c:axPos", layout.position);
    if (options.majorGridlines) {
        xml.leafNode("c:majorGridlines");
    }
    if (options.title !== undefined) {
        title(xml, options.title, layout.position === "l");
    }
    if (options.numFmt !== undefined) {
        xml.leafNode("c:numFmt", { formatCode: options.numFmt, sourceLinked: 0 });
    } else {
        xml.leafNode("c:numFmt", { formatCode: "General", sourceLinked: 1 });
    }
    val(xml, "c:majorTickMark", "out");
    val(xml, "c:minorTickMark", "none");
    val(xml, "c:tickLblPos", "nextTo");
    val(xml, "c:crossAx", layout.crossAx);
    val(xml, "c:crosses", "autoZero");
    if (layout.element === "c:catAx") {
        // CT_CatAx: auto, lblAlgn, lblOffset, tickLblSkip, tickMarkSkip, noMultiLvlLbl
        val(xml, "c:auto", 1);
        val(xml, "c:lblAlgn", "ctr");
        val(xml, "c:lblOffset", 100);
        val(xml, "c:noMultiLvlLbl", 0);
    } else {
        // CT_ValAx: crossBetween (the points of an area or scatter chart are on the categories, the bars between them)
        val(xml, "c:crossBetween", spec.type === "area" || spec.type === "scatter" ? "midCat" : "between");
    }
    xml.closeNode();
}

/** The two axes of a chart: the category axis (scatter: x values) first, then the value axis. */
function axes(xml: XmlStream, spec: ChartSpaceSpec): void {
    // a bar chart is horizontal: its categories on the left, its values at the bottom
    const horizontal = spec.type === "bar";
    axis(
        xml,
        spec,
        {
            element: spec.type === "scatter" ? "c:valAx" : "c:catAx",
            id: CATEGORY_AXIS_ID,
            crossAx: VALUE_AXIS_ID,
            position: horizontal ? "l" : "b",
        },
        spec.categoryAxis
    );
    axis(xml, spec, { element: "c:valAx", id: VALUE_AXIS_ID, crossAx: CATEGORY_AXIS_ID, position: horizontal ? "b" : "l" }, spec.valueAxis);
}

/**
 * The XML of the chart part (`xl/charts/chartN.xml`) of a chart: CT_ChartSpace (date1904, roundedCorners, chart), CT_Chart
 * (title, autoTitleDeleted, plotArea, legend, plotVisOnly, dispBlanksAs), CT_PlotArea (layout, the chart, the axes).
 *
 * @param spec - validated (src/doc/chart.ts): no check is made here
 */
export function renderChartSpace(spec: ChartSpaceSpec): string {
    const xml = new XmlStream();
    xml.openXml(XmlStream.StdDocAttributes);
    xml.openNode("c:chartSpace", { "xmlns:c": NS_CHART, "xmlns:a": NS_MAIN, "xmlns:r": NS_RELATIONSHIPS });
    val(xml, "c:date1904", spec.date1904 ? 1 : 0);
    // Excel rounds the corners of a chart without this element
    val(xml, "c:roundedCorners", 0);
    xml.openNode("c:chart");
    if (spec.title !== undefined) {
        title(xml, spec.title);
        val(xml, "c:autoTitleDeleted", 0);
    } else {
        // otherwise Excel titles a chart of one series with the name of the series
        val(xml, "c:autoTitleDeleted", 1);
    }
    xml.openNode("c:plotArea");
    xml.leafNode("c:layout");
    plot(xml, spec);
    if (!isPie(spec)) {
        axes(xml, spec);
    }
    xml.closeNode();
    if (spec.legend !== undefined) {
        // CT_Legend: legendPos, legendEntry, layout, overlay
        xml.openNode("c:legend");
        val(xml, "c:legendPos", spec.legend);
        val(xml, "c:overlay", 0);
        xml.closeNode();
    }
    val(xml, "c:plotVisOnly", 1);
    val(xml, "c:dispBlanksAs", "gap");
    xml.closeNode();
    xml.closeNode();
    return xml.xml;
}

// the anchor of a chart as a tree of elements, as chart-anchor.ts ChartAnchorCapture reads it from a file
function element(
    name: string,
    attributes: readonly ChartXmlAttribute[] = [],
    children: readonly (ChartXmlElement | string)[] = []
): ChartXmlElement {
    return { name, attributes, children };
}

function attribute(name: string, value: string | number): ChartXmlAttribute {
    return { name, value: String(value) };
}

/** A corner of the anchor, `<xdr:from>` / `<xdr:to>` (CT_Marker: col, colOff, row, rowOff). */
function marker(name: string, position: AnchorModel): ChartXmlElement {
    return element(
        name,
        [],
        [
            element("xdr:col", [], [String(position.nativeCol)]),
            element("xdr:colOff", [], [String(position.nativeColOff)]),
            element("xdr:row", [], [String(position.nativeRow)]),
            element("xdr:rowOff", [], [String(position.nativeRowOff)]),
        ]
    );
}

/** The id of the relationship of the drawing to the chart part, in the anchor built (renumbered on write). */
const ANCHOR_REL_ID = "rId1";

/**
 * The anchor of a chart in the drawing of its worksheet, and its relationship to the chart part: the model of a chart read
 * from a file (chart-parts.ts), written by the same code. CT_TwoCellAnchor (from, to, graphicFrame, clientData),
 * CT_OneCellAnchor (from, ext, graphicFrame, clientData) or, for the chart of a chartsheet, CT_AbsoluteAnchor (pos, ext,
 * graphicFrame, clientData); CT_GraphicalObjectFrame (nvGraphicFramePr, xfrm, graphic). The shape id (`<xdr:cNvPr id>`)
 * and the relationship id are renumbered when the drawing is written.
 *
 * @param anchor - validated (src/doc/chart.ts): `to` or `ext`, or `absolute`
 * @param xml - the chart part, {@link renderChartSpace}
 */
export function chartAnchorModel(anchor: ChartAnchorSpec | ChartAbsoluteAnchorSpec, xml: string): ChartAnchorModel {
    const part: ChartPartModel = { kind: "chart", extension: "xml", data: new TextEncoder().encode(xml), rels: [] };
    const frame = element(
        "xdr:graphicFrame",
        [attribute("macro", "")],
        [
            element(
                "xdr:nvGraphicFramePr",
                [],
                [
                    element("xdr:cNvPr", [{ name: "id", value: "1", role: "shapeId" }, attribute("name", anchor.name)]),
                    element("xdr:cNvGraphicFramePr"),
                ]
            ),
            element(
                "xdr:xfrm",
                [],
                [element("a:off", [attribute("x", 0), attribute("y", 0)]), element("a:ext", [attribute("cx", 0), attribute("cy", 0)])]
            ),
            element(
                "a:graphic",
                [],
                [
                    element(
                        "a:graphicData",
                        [attribute("uri", NS_CHART)],
                        [
                            element("c:chart", [
                                attribute("xmlns:c", NS_CHART),
                                attribute("xmlns:r", NS_RELATIONSHIPS),
                                { name: "r:id", value: ANCHOR_REL_ID, role: "rel" },
                            ]),
                        ]
                    ),
                ]
            ),
        ]
    );
    const clientData = element("xdr:clientData");
    let root: ChartXmlElement;
    if ("absolute" in anchor) {
        const { x, y, cx, cy } = anchor.absolute;
        root = element(
            "xdr:absoluteAnchor",
            [],
            [
                element("xdr:pos", [attribute("x", x), attribute("y", y)]),
                element("xdr:ext", [attribute("cx", cx), attribute("cy", cy)]),
                frame,
                clientData,
            ]
        );
    } else if (anchor.to) {
        root = element("xdr:twoCellAnchor", anchor.editAs ? [attribute("editAs", anchor.editAs)] : [], [
            marker("xdr:from", anchor.from),
            marker("xdr:to", anchor.to),
            frame,
            clientData,
        ]);
    } else {
        const ext = anchor.ext ?? { width: 0, height: 0 };
        root = element(
            "xdr:oneCellAnchor",
            [],
            [
                marker("xdr:from", anchor.from),
                element("xdr:ext", [
                    attribute("cx", Math.floor(ext.width * EMU_PER_PIXEL)),
                    attribute("cy", Math.floor(ext.height * EMU_PER_PIXEL)),
                ]),
                frame,
                clientData,
            ]
        );
    }
    return { anchor: root, rels: [{ Id: ANCHOR_REL_ID, Type: REL_CHART, part }] };
}
