// Hand-written pivot table fixtures (spec/integration/pivot-tables-read.spec.js): a workbook written by ExcelJS (Data: the
// source, Pivot: the pivot table), to which a minimal but valid pivot table, its cache definition and records, and
// optionally a pivot chart, are added. Built in code (xlsx-tamper), no binary fixture.
const { strFromU8, strToU8 } = require("fflate");
const { baseWorkbookBuffer, editParts } = require("./xlsx-tamper");

const NS = {
    main: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    pkg: "http://schemas.openxmlformats.org/package/2006/relationships",
    c: "http://schemas.openxmlformats.org/drawingml/2006/chart",
    a: "http://schemas.openxmlformats.org/drawingml/2006/main",
    xdr: "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    x14: "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
};
const REL = {
    pivotTable: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable",
    pivotCacheDefinition: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition",
    pivotCacheRecords: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords",
    externalLinkPath: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath",
    drawing: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
    chart: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
    slicer: "http://schemas.microsoft.com/office/2007/relationships/slicer",
    slicerCache: "http://schemas.microsoft.com/office/2007/relationships/slicerCache",
};
const CT = {
    pivotTable: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml",
    pivotCacheDefinition: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml",
    pivotCacheRecords: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml",
    drawing: "application/vnd.openxmlformats-officedocument.drawing+xml",
    chart: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
};

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

// the source (Data!A1:C4) and the pivot table of the Pivot worksheet (A3:B6): the sum of Sales by Region
const DATA = [
    ["Region", "Product", "Sales"],
    ["East", "A", 10],
    ["West", "B", 20],
    ["East", "B", 30],
];
const PIVOT_CELLS = [
    ["A3", "Row Labels"],
    ["B3", "Sum of Sales"],
    ["A4", "East"],
    ["B4", 40],
    ["A5", "West"],
    ["B5", 20],
    ["A6", "Grand Total"],
    ["B6", 60],
];

function cacheDefinitionXml({ source = '<cacheSource type="worksheet"><worksheetSource ref="A1:C4" sheet="Data"/></cacheSource>' } = {}) {
    return (
        `${XML_DECL}<pivotCacheDefinition xmlns="${NS.main}" xmlns:r="${NS.r}" r:id="rId1" refreshOnLoad="1" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="3">` +
        source +
        '<cacheFields count="3">' +
        '<cacheField name="Region" numFmtId="0"><sharedItems count="2"><s v="East"/><s v="West"/></sharedItems></cacheField>' +
        '<cacheField name="Product" numFmtId="0"><sharedItems count="2"><s v="A"/><s v="B"/></sharedItems></cacheField>' +
        '<cacheField name="Sales" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="10" maxValue="30"/></cacheField>' +
        "</cacheFields></pivotCacheDefinition>"
    );
}
const CACHE_DEFINITION_XML = cacheDefinitionXml();
const CACHE_RECORDS_XML = `${XML_DECL}<pivotCacheRecords xmlns="${NS.main}" xmlns:r="${NS.r}" count="3"><r><x v="0"/><x v="0"/><n v="10"/></r><r><x v="1"/><x v="1"/><n v="20"/></r><r><x v="0"/><x v="1"/><n v="30"/></r></pivotCacheRecords>`;

function pivotTableXml({ name = "PivotTable1", cacheId = 5, extLst = "" } = {}) {
    return (
        `${XML_DECL}<pivotTableDefinition xmlns="${NS.main}" name="${name}" cacheId="${cacheId}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" outline="1" outlineData="1" multipleFieldFilters="0" chartFormat="1">` +
        '<location ref="A3:B6" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>' +
        '<pivotFields count="3"><pivotField axis="axisRow" showAll="0"><items count="3"><item x="0"/><item x="1"/><item t="default"/></items></pivotField><pivotField showAll="0"/><pivotField dataField="1" showAll="0"/></pivotFields>' +
        '<rowFields count="1"><field x="0"/></rowFields>' +
        '<rowItems count="3"><i><x/></i><i><x v="1"/></i><i t="grand"><x/></i></rowItems>' +
        '<colItems count="1"><i/></colItems>' +
        '<dataFields count="1"><dataField name="Sum of Sales" fld="2" baseField="0" baseItem="0"/></dataFields>' +
        '<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>' +
        extLst +
        "</pivotTableDefinition>"
    );
}

// a pivot chart of the pivot table PivotTable1 of the Pivot worksheet (column chart of B4:B5 by A4:A5)
// the extensions of the chart space of a pivot chart (c14:pivotOptions, c16:pivotOptions16), and another one
const PIVOT_CHART_EXTENSIONS = {
    pivotOptions:
        '<c:ext uri="{781A3756-C4B2-4CAC-9D66-4F8BD8637D16}" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:pivotOptions><c14:dropZoneFilter val="1"/><c14:dropZoneCategories val="1"/><c14:dropZoneData val="1"/></c14:pivotOptions></c:ext>',
    pivotOptions16:
        '<c:ext uri="{E28EC0CA-F0BB-4C9C-879D-F8772B89E7AC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"><c16:pivotOptions16><c16:showExpandCollapseFieldButtons val="1"/></c16:pivotOptions16></c:ext>',
    other: '<c:ext uri="{56B9EC1D-385E-4148-901F-78D8002777C0}" xmlns:c16r3="http://schemas.microsoft.com/office/drawing/2017/03/chart"><c16r3:dataDisplayOptions16><c16r3:dispNaAsBlank val="1"/></c16r3:dataDisplayOptions16></c:ext>',
};

/**
 * A pivot chart of the pivot table `source` (column chart of Pivot!B4:B5 by A4:A5). `pivot: false`: the same chart without
 * its pivot elements (`<c:pivotSource>`, `<c:pivotFmts>`), as written when its pivot table is not; `extensions`: the
 * `<c:ext>` of its `<c:extLst>`.
 */
function pivotChartXml({ source = "[Book1.xlsx]Pivot!PivotTable1", pivot = true, extensions = [] } = {}) {
    return (
        `${XML_DECL}<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}"><c:roundedCorners val="0"/>` +
        (pivot ? `<c:pivotSource><c:name>${source}</c:name><c:fmtId val="0"/></c:pivotSource>` : "") +
        '<c:chart><c:autoTitleDeleted val="1"/>' +
        (pivot ? '<c:pivotFmts><c:pivotFmt><c:idx val="0"/></c:pivotFmt></c:pivotFmts>' : "") +
        '<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>' +
        '<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Pivot!$B$3</c:f></c:strRef></c:tx>' +
        "<c:cat><c:strRef><c:f>Pivot!$A$4:$A$5</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>Pivot!$B$4:$B$5</c:f></c:numRef></c:val></c:ser>" +
        '<c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>' +
        '<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea><c:plotVisOnly val="1"/></c:chart>' +
        (extensions.length ? `<c:extLst>${extensions.join("")}</c:extLst>` : "") +
        "</c:chartSpace>"
    );
}
const PIVOT_CHART_XML = pivotChartXml();

const CHART_ANCHOR =
    '<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
    "<xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>14</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>" +
    '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>' +
    '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
    `<c:chart xmlns:c="${NS.c}" xmlns:r="${NS.r}" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;

function rels(list) {
    const items = list
        .map(
            ([id, type, target, external]) =>
                `<Relationship Id="${id}" Type="${type}" Target="${target}"${external ? ' TargetMode="External"' : ""}/>`
        )
        .join("");
    return `${XML_DECL}<Relationships xmlns="${NS.pkg}">${items}</Relationships>`;
}

/**
 * A workbook with the worksheets Data (the source) and Pivot, whose pivot table PivotTable1 (cacheId 5) uses the cache
 * pivotCacheDefinition1 and its records. `chart`: a pivot chart in the drawing of Pivot (`chartXml`). `edit(parts)` may change the parts
 * (strings) before they are zipped.
 */
async function pivotWorkbook({ chart = false, chartXml = PIVOT_CHART_XML, edit } = {}) {
    const base = await baseWorkbookBuffer((wb, ws) => {
        ws.name = "Data";
        DATA.forEach((values, index) => {
            ws.getRow(index + 1).values = values;
        });
        const pivot = wb.addWorksheet("Pivot");
        for (const [address, value] of PIVOT_CELLS) {
            pivot.getCell(address).value = value;
        }
    });
    return editParts(base, (parts) => {
        const text = Object.create(null);
        for (const [name, value] of Object.entries(parts)) {
            text[name] = strFromU8(value);
        }
        text["xl/workbook.xml"] = text["xl/workbook.xml"].replace(
            "</workbook>",
            '<pivotCaches><pivotCache cacheId="5" r:id="rId99"/></pivotCaches></workbook>'
        );
        text["xl/_rels/workbook.xml.rels"] = text["xl/_rels/workbook.xml.rels"].replace(
            "</Relationships>",
            `<Relationship Id="rId99" Type="${REL.pivotCacheDefinition}" Target="pivotCache/pivotCacheDefinition1.xml"/></Relationships>`
        );
        const sheetRels = [["rId1", REL.pivotTable, "../pivotTables/pivotTable1.xml"]];
        let overrides =
            `<Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="${CT.pivotTable}"/>` +
            `<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="${CT.pivotCacheDefinition}"/>` +
            `<Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="${CT.pivotCacheRecords}"/>`;
        if (chart) {
            sheetRels.push(["rId2", REL.drawing, "../drawings/drawing1.xml"]);
            text["xl/worksheets/sheet2.xml"] = text["xl/worksheets/sheet2.xml"].replace(
                "</worksheet>",
                '<drawing r:id="rId2"/></worksheet>'
            );
            text["xl/drawings/drawing1.xml"] = `${XML_DECL}<xdr:wsDr xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}">${CHART_ANCHOR}</xdr:wsDr>`;
            text["xl/drawings/_rels/drawing1.xml.rels"] = rels([["rId1", REL.chart, "../charts/chart1.xml"]]);
            text["xl/charts/chart1.xml"] = chartXml;
            overrides +=
                `<Override PartName="/xl/drawings/drawing1.xml" ContentType="${CT.drawing}"/>` +
                `<Override PartName="/xl/charts/chart1.xml" ContentType="${CT.chart}"/>`;
        }
        text["xl/worksheets/_rels/sheet2.xml.rels"] = rels(sheetRels);
        text["xl/pivotTables/pivotTable1.xml"] = pivotTableXml();
        text["xl/pivotTables/_rels/pivotTable1.xml.rels"] = rels([
            ["rId1", REL.pivotCacheDefinition, "../pivotCache/pivotCacheDefinition1.xml"],
        ]);
        text["xl/pivotCache/pivotCacheDefinition1.xml"] = CACHE_DEFINITION_XML;
        text["xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels"] = rels([["rId1", REL.pivotCacheRecords, "pivotCacheRecords1.xml"]]);
        text["xl/pivotCache/pivotCacheRecords1.xml"] = CACHE_RECORDS_XML;
        text["[Content_Types].xml"] = text["[Content_Types].xml"].replace("</Types>", `${overrides}</Types>`);
        if (edit) {
            edit(text);
        }
        for (const name of Object.keys(parts)) {
            delete parts[name];
        }
        for (const [name, value] of Object.entries(text)) {
            if (value !== undefined) {
                parts[name] = typeof value === "string" ? strToU8(value) : value;
            }
        }
    });
}

module.exports = {
    NS,
    REL,
    CT,
    XML_DECL,
    DATA,
    CACHE_DEFINITION_XML,
    CACHE_RECORDS_XML,
    PIVOT_CHART_XML,
    PIVOT_CHART_EXTENSIONS,
    pivotChartXml,
    cacheDefinitionXml,
    pivotTableXml,
    rels,
    pivotWorkbook,
};
