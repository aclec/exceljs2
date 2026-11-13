// Pivot tables of the worksheets are carried opaque from a file read to the file written: the pivot table part, its pivot
// cache definition and records, with their relationships, and the <pivotCache> of xl/workbook.xml. The fixtures are built
// in code (spec/utils/pivot-fixtures.js).
const path = require("path");
const { strFromU8 } = require("fflate");
const { baseWorkbookBuffer, unzipParts } = require("../utils/xlsx-tamper");
const { validateDrawingParts } = require("../utils/xmllint");
const {
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
} = require("../utils/pivot-fixtures");

const ExcelJS = verquire("exceljs");
const { loadPivotParts, PivotPartWriter } = verquire("xlsx/pivot-parts");
async function roundTrip(buffer, change) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    if (change) {
        change(wb);
    }
    return { wb, buffer: await wb.xlsx.writeBuffer() };
}

// the parts of a zip, as text
function textParts(buffer) {
    const parts = unzipParts(buffer);
    return { parts, text: (name) => (parts[name] ? strFromU8(parts[name]) : undefined) };
}

// the relationships of a .rels part: [{Id, Type, Target, TargetMode}]
function parseRels(xml) {
    return [...xml.matchAll(/<Relationship ([^>]*)\/>/g)].map(([, attributes]) =>
        Object.fromEntries([...attributes.matchAll(/(\w+)="([^"]*)"/g)].map(([, name, value]) => [name, value]))
    );
}

function overrides(contentTypes) {
    return Object.fromEntries(
        [...contentTypes.matchAll(/<Override PartName="([^"]+)" ContentType="([^"]+)"\/>/g)].map(([, n, t]) => [n, t])
    );
}

function pivotPartNames(buffer) {
    return Object.keys(unzipParts(buffer)).filter((name) => /pivot/i.test(name));
}

// the <pivotCache>s of xl/workbook.xml: {cacheId: target of its workbook relationship}
function workbookCaches(text) {
    const bookRels = parseRels(text("xl/_rels/workbook.xml.rels"));
    const caches = {};
    for (const [, cacheId, rId] of text("xl/workbook.xml").matchAll(/<pivotCache cacheId="(\d+)" r:id="([^"]+)"\/>/g)) {
        const rel = bookRels.find((r) => r.Id === rId);
        expect(rel.Type).to.equal(REL.pivotCacheDefinition);
        caches[cacheId] = `xl/${rel.Target}`;
    }
    return caches;
}

// a part name resolved from a relationship target (relative to the part holding the relationship)
function resolve(source, target) {
    const segments = source.split("/").slice(0, -1);
    for (const segment of target.split("/")) {
        if (segment === "..") segments.pop();
        else segments.push(segment);
    }
    return segments.join("/");
}

/**
 * Checks a pivot table written: the worksheet relationship, the pivot table part (its bytes as read, `tableXml`), its cache
 * (the relationship and the cacheId of xl/workbook.xml agree), the records, the content types.
 * @returns the part names of the pivot table and its cache
 */
function expectPivotTable(buffer, { sheet = "xl/worksheets/sheet2.xml", tableXml = pivotTableXml(), cacheId = "5" } = {}) {
    const { text } = textParts(buffer);
    const sheetRelsName = sheet.replace("worksheets/", "worksheets/_rels/") + ".rels";
    const tableNames = parseRels(text(sheetRelsName))
        .filter((rel) => rel.Type === REL.pivotTable)
        .map((rel) => resolve(sheet, rel.Target));
    const tableName = tableNames.find((name) => text(name) === tableXml);
    expect(tableName, "pivot table part").to.be.a("string");

    const tableRels = parseRels(text(tableName.replace("pivotTables/", "pivotTables/_rels/") + ".rels"));
    expect(tableRels).to.have.length(1);
    expect(tableRels[0].Id).to.equal("rId1");
    expect(tableRels[0].Type).to.equal(REL.pivotCacheDefinition);
    const cacheName = resolve(tableName, tableRels[0].Target);
    expect(text(cacheName)).to.equal(CACHE_DEFINITION_XML);
    expect(workbookCaches(text)[cacheId]).to.equal(cacheName);

    const cacheRels = parseRels(text(cacheName.replace("pivotCache/", "pivotCache/_rels/") + ".rels"));
    expect(cacheRels).to.have.length(1);
    // the id that <pivotCacheDefinition r:id="rId1"> references
    expect(cacheRels[0].Id).to.equal("rId1");
    expect(cacheRels[0].Type).to.equal(REL.pivotCacheRecords);
    const recordsName = resolve(cacheName, cacheRels[0].Target);
    expect(text(recordsName)).to.equal(CACHE_RECORDS_XML);

    const types = overrides(text("[Content_Types].xml"));
    expect(types[`/${tableName}`]).to.equal(CT.pivotTable);
    expect(types[`/${cacheName}`]).to.equal(CT.pivotCacheDefinition);
    expect(types[`/${recordsName}`]).to.equal(CT.pivotCacheRecords);
    return { tableName, cacheName, recordsName, parts: textParts(buffer).parts, text };
}

// checks that a written file has no pivot table at all, and no reference to one
function expectNoPivotTable(buffer) {
    const { text } = textParts(buffer);
    expect(pivotPartNames(buffer)).to.deep.equal([]);
    expect(text("xl/workbook.xml")).to.not.include("pivotCache");
    expect(text("xl/_rels/workbook.xml.rels")).to.not.include("pivot");
    expect(text("[Content_Types].xml")).to.not.include("pivot");
    expect(text("xl/worksheets/_rels/sheet2.xml.rels") ?? "").to.not.include("pivot");
}

describe("Pivot tables carried from a file read to the file written", () => {
    it("keeps a pivot table, its cache and its records, with consistent relationships and cacheId", async () => {
        const first = await roundTrip(await pivotWorkbook());
        const { tableName, cacheName, recordsName, text } = expectPivotTable(first.buffer);
        expect([tableName, cacheName, recordsName]).to.deep.equal([
            "xl/pivotTables/pivotTable1.xml",
            "xl/pivotCache/pivotCacheDefinition1.xml",
            "xl/pivotCache/pivotCacheRecords1.xml",
        ]);
        // <pivotCaches> after <calcPr> (CT_Workbook), the last child
        expect(text("xl/workbook.xml")).to.match(
            /<calcPr [^>]*\/><pivotCaches><pivotCache cacheId="5" r:id="rId\d+"\/><\/pivotCaches><\/workbook>$/
        );
        // the cells of the pivot table are kept as values
        expect(first.wb.getWorksheet("Pivot").getCell("B6").value).to.equal(60);
        // no pivot table of the editing API
        expect(first.wb.pivotTables).to.deep.equal([]);

        // load → write again: the same parts
        const second = await roundTrip(first.buffer);
        expectPivotTable(second.buffer);
        const a = textParts(first.buffer);
        const b = textParts(second.buffer);
        for (const name of [
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            "xl/worksheets/_rels/sheet2.xml.rels",
            "[Content_Types].xml",
        ]) {
            expect(b.text(name), name).to.equal(a.text(name));
        }
    });

    it("keeps a pivot chart with its pivot table: the chart names the pivot table written", async () => {
        const { buffer } = await roundTrip(await pivotWorkbook({ chart: true }));
        const { text } = expectPivotTable(buffer);
        expect(text("xl/charts/chart1.xml")).to.equal(PIVOT_CHART_XML);
        // the chart of the drawing of the Pivot worksheet, the pivot table of the same worksheet, named as the chart says
        const sheetRels = parseRels(text("xl/worksheets/_rels/sheet2.xml.rels"));
        expect(sheetRels.map((rel) => rel.Type).sort()).to.deep.equal([REL.drawing, REL.pivotTable].sort());
        expect(text("xl/workbook.xml")).to.include('name="Pivot"');
        expect(text("xl/pivotTables/pivotTable1.xml")).to.include('name="PivotTable1"');
        expect(text("xl/charts/chart1.xml")).to.include("<c:name>[Book1.xlsx]Pivot!PivotTable1</c:name>");
        expect(parseRels(text("xl/drawings/_rels/drawing1.xml.rels"))[0].Target).to.equal("../charts/chart1.xml");
    });

    it("writes the pivot tables that share a cache with one cache, and keeps external relationships", async () => {
        const buffer = await pivotWorkbook({
            edit(parts) {
                parts["xl/worksheets/_rels/sheet2.xml.rels"] = rels([
                    ["rId1", REL.pivotTable, "../pivotTables/pivotTable1.xml"],
                    ["rId2", REL.pivotTable, "/xl/pivotTables/pivotTable9.xml"],
                ]);
                parts["xl/pivotTables/pivotTable9.xml"] = pivotTableXml({ name: "PivotTable2" });
                parts["xl/pivotTables/_rels/pivotTable9.xml.rels"] = rels([
                    ["rId1", REL.pivotCacheDefinition, "../pivotCache/pivotCacheDefinition1.xml"],
                ]);
                parts["xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels"] = rels([
                    ["rId1", REL.pivotCacheRecords, "pivotCacheRecords1.xml"],
                    ["rId2", REL.externalLinkPath, "file:///C:/a&amp;b.xlsx", true],
                ]);
            },
        });
        const { buffer: written } = await roundTrip(buffer);
        const { text } = textParts(written);
        expect(pivotPartNames(written).sort()).to.deep.equal(
            [
                "xl/pivotTables/pivotTable1.xml",
                "xl/pivotTables/_rels/pivotTable1.xml.rels",
                "xl/pivotTables/pivotTable2.xml",
                "xl/pivotTables/_rels/pivotTable2.xml.rels",
                "xl/pivotCache/pivotCacheDefinition1.xml",
                "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
                "xl/pivotCache/pivotCacheRecords1.xml",
            ].sort()
        );
        expect(text("xl/pivotTables/pivotTable2.xml")).to.equal(pivotTableXml({ name: "PivotTable2" }));
        expect(workbookCaches(text)).to.deep.equal({ 5: "xl/pivotCache/pivotCacheDefinition1.xml" });
        expect(parseRels(text("xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels"))).to.deep.equal([
            { Id: "rId1", Type: REL.pivotCacheRecords, Target: "pivotCacheRecords1.xml" },
            { Id: "rId2", Type: REL.externalLinkPath, Target: "file:///C:/a&amp;b.xlsx", TargetMode: "External" },
        ]);
    });

    it("writes the pivot tables of a worksheet copied twice with one cache; drops those of a worksheet removed", async () => {
        const { buffer } = await roundTrip(await pivotWorkbook(), (wb) => {
            const copy = wb.addWorksheet("Copy");
            copy.model = { ...wb.getWorksheet("Pivot").model, id: copy.id, name: "Copy" };
        });
        const { text } = textParts(buffer);
        expectPivotTable(buffer);
        const copyName = parseRels(text("xl/worksheets/_rels/sheet3.xml.rels")).find((rel) => rel.Type === REL.pivotTable);
        expect(copyName.Target).to.equal("../pivotTables/pivotTable2.xml");
        expect(text("xl/pivotTables/pivotTable2.xml")).to.equal(pivotTableXml());
        expect(pivotPartNames(buffer).filter((name) => name.includes("pivotCacheDefinition"))).to.have.length(2); // part and rels

        const { buffer: removed } = await roundTrip(await pivotWorkbook(), (wb) => wb.removeWorksheet(wb.getWorksheet("Pivot").id));
        expect(pivotPartNames(removed)).to.deep.equal([]);
        expect(textParts(removed).text("xl/workbook.xml")).to.not.include("pivotCache");
    });

    it("keeps the pivot table of a renamed worksheet; the source of its cache is not rewritten", async () => {
        const { buffer } = await roundTrip(await pivotWorkbook(), (wb) => {
            wb.getWorksheet("Data").name = "Source";
        });
        const { text } = expectPivotTable(buffer);
        expect(text("xl/pivotCache/pivotCacheDefinition1.xml")).to.include('sheet="Data"');
    });

    describe("with a pivot table added by addPivotTable", () => {
        const INPUT = { rows: ["Region"], columns: ["Product"], values: ["Sales"] };
        function addPivotTable(wb, sheetName = "Added") {
            const target = wb.addWorksheet(sheetName);
            target.addPivotTable({ sourceSheet: wb.getWorksheet("Data"), ...INPUT });
        }

        it("refuses to add a pivot table to a worksheet that carries pivot tables read from the file", async () => {
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.load(await pivotWorkbook());
            const pivot = wb.getWorksheet("Pivot");
            // its fixed location (A3:E15) and name would overlap or duplicate those of the pivot tables read
            expect(() => pivot.addPivotTable({ sourceSheet: wb.getWorksheet("Data"), ...INPUT })).to.throw(
                /Pivot already has pivot tables read from the file/
            );
            expect(wb.pivotTables).to.deep.equal([]);
            expect(pivot.pivotTables).to.deep.equal([]);
            // another worksheet still can
            addPivotTable(wb);
            expect(wb.pivotTables).to.have.length(1);
        });

        it("refuses to write a worksheet that has both a pivot table added and pivot tables read", async () => {
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.load(await pivotWorkbook());
            addPivotTable(wb);
            // a pivot table moved by hand to the worksheet that carries one
            wb.getWorksheet("Pivot").pivotTables = wb.getWorksheet("Added").pivotTables;
            wb.getWorksheet("Added").pivotTables = [];
            let error;
            try {
                await wb.xlsx.writeBuffer();
            } catch (e) {
                error = e;
            }
            expect(error && error.message).to.match(/Pivot has both a pivot table added .* and pivot tables read from the file/);
        });

        it("writes both: the pivot table added as pivotTable1 with its cacheId, the one read numbered after it", async () => {
            const { buffer } = await roundTrip(await pivotWorkbook(), (wb) => addPivotTable(wb));
            const { tableName, cacheName, text } = expectPivotTable(buffer);
            expect(tableName).to.equal("xl/pivotTables/pivotTable2.xml");
            expect(cacheName).to.equal("xl/pivotCache/pivotCacheDefinition2.xml");
            expect(workbookCaches(text)).to.deep.equal({
                5: "xl/pivotCache/pivotCacheDefinition2.xml",
                10: "xl/pivotCache/pivotCacheDefinition1.xml",
            });
            expect(text("xl/pivotTables/pivotTable1.xml")).to.include('cacheId="10"');
            const types = overrides(text("[Content_Types].xml"));
            for (const name of [
                "/xl/pivotTables/pivotTable1.xml",
                "/xl/pivotTables/pivotTable2.xml",
                "/xl/pivotCache/pivotCacheDefinition1.xml",
                "/xl/pivotCache/pivotCacheDefinition2.xml",
                "/xl/pivotCache/pivotCacheRecords1.xml",
                "/xl/pivotCache/pivotCacheRecords2.xml",
            ]) {
                expect(types[name], name).to.be.a("string");
            }
        });

        it("gives the pivot table added another cacheId when a pivot table read has its cacheId", async () => {
            // a file written by ExcelJS with addPivotTable (cacheId 10), read, and a pivot table added again
            const source = new ExcelJS.Workbook();
            source.addWorksheet("Data").addRows(DATA);
            addPivotTable(source, "First");
            const written = await source.xlsx.writeBuffer();

            const { wb, buffer } = await roundTrip(written, (book) => addPivotTable(book, "Second"));
            expect(wb.pivotTables).to.have.length(1);
            // the pivot table added is not changed: only its part
            expect(wb.pivotTables[0].cacheId).to.equal("10");
            const { text } = textParts(buffer);
            expect(Object.keys(workbookCaches(text)).sort()).to.deep.equal(["10", "11"]);
            expect(text("xl/pivotTables/pivotTable1.xml")).to.include('cacheId="11"');
            expect(text("xl/pivotTables/pivotTable2.xml")).to.include('cacheId="10"');
            expect(workbookCaches(text)[10]).to.equal("xl/pivotCache/pivotCacheDefinition2.xml");
            expect(workbookCaches(text)[11]).to.equal("xl/pivotCache/pivotCacheDefinition1.xml");

            // read again: both pivot tables are carried
            const { buffer: again } = await roundTrip(buffer);
            const parts = textParts(again);
            expect(Object.keys(workbookCaches(parts.text)).sort()).to.deep.equal(["10", "11"]);
            expect(pivotPartNames(again).filter((name) => /pivotTable\d\.xml$/.test(name))).to.have.length(2);
        });
    });

    describe("pivot charts", () => {
        const SAMPLES = path.join(__dirname, "../out/pivot-samples/xml");

        // the chart written, validated against dml-chart.xsd when OOXML_SCHEMAS and xmllint are available
        function writtenChart(buffer, label) {
            const parts = unzipParts(buffer);
            const result = validateDrawingParts(parts, path.join(SAMPLES, label));
            if (result) {
                expect(result.ok, result.output).to.equal(true);
            }
            return strFromU8(parts["xl/charts/chart1.xml"]);
        }

        const dropped = {
            "its cache comes from an external connection": (parts) => {
                parts["xl/pivotCache/pivotCacheDefinition1.xml"] = cacheDefinitionXml({
                    source: '<cacheSource type="external" connectionId="1"/>',
                });
            },
            "its cacheId does not match": (parts) => {
                parts["xl/pivotTables/pivotTable1.xml"] = pivotTableXml({ cacheId: 6 });
            },
            "its part is missing": (parts) => {
                delete parts["xl/pivotTables/pivotTable1.xml"];
            },
        };
        for (const [reason, edit] of Object.entries(dropped)) {
            it(`writes a static chart when its pivot table is dropped: ${reason}`, async () => {
                const extensions = Object.values(PIVOT_CHART_EXTENSIONS);
                const { wb, buffer } = await roundTrip(await pivotWorkbook({ chart: true, chartXml: pivotChartXml({ extensions }), edit }));
                expectNoPivotTable(buffer);
                expect(wb.getWorksheet("Pivot")._charts).to.have.length(1);
                // as Excel turns the pivot chart of a pivot table deleted into a static chart: no pivot source, no pivot
                // formats, no pivot options; the other extensions, the series and their cached values are kept
                const chart = writtenChart(buffer, reason.replace(/\W+/g, "-"));
                expect(chart).to.equal(pivotChartXml({ pivot: false, extensions: [PIVOT_CHART_EXTENSIONS.other] }));
            });
        }

        it("writes a static chart when the pivot table it names is not on the worksheet it names", async () => {
            for (const source of ["[Book1.xlsx]Pivot!PivotTable9", "[Book1.xlsx]Data!PivotTable1", "PivotTable1", "[Book1.xlsx]"]) {
                const { buffer } = await roundTrip(await pivotWorkbook({ chart: true, chartXml: pivotChartXml({ source }) }));
                expectPivotTable(buffer);
                expect(writtenChart(buffer, "other-name"), source).to.equal(pivotChartXml({ pivot: false }));
            }
        });

        it("keeps a pivot chart whose source names its pivot table in any of the forms of Excel", async () => {
            for (const source of [
                "[Book1.xlsx]Pivot!PivotTable1",
                "[Other File.xlsx]pivot!PivotTable1",
                "Pivot!PivotTable1",
                "'Pivot'!PivotTable1",
                "'[Book1.xlsx]Pivot'!PivotTable1",
            ]) {
                const chartXml = pivotChartXml({ source, extensions: Object.values(PIVOT_CHART_EXTENSIONS) });
                const { buffer } = await roundTrip(await pivotWorkbook({ chart: true, chartXml }));
                expectPivotTable(buffer);
                expect(writtenChart(buffer, "kept"), source).to.equal(chartXml);
            }
        });

        it("writes a static chart when the worksheet of its pivot table is renamed (its source is not rewritten)", async () => {
            const { buffer } = await roundTrip(await pivotWorkbook({ chart: true }), (wb) => {
                wb.getWorksheet("Pivot").name = "Renamed";
            });
            expectPivotTable(buffer);
            expect(writtenChart(buffer, "renamed")).to.equal(pivotChartXml({ pivot: false }));
        });

        it("matches a quoted worksheet name with a quote", async () => {
            const writer = new PivotPartWriter();
            const cache = { kind: "pivotCacheDefinition", cacheId: 5, data: new Uint8Array([1]), rels: [] };
            writer.addTable({ kind: "pivotTable", name: "PT", data: new Uint8Array([2]), cache, rels: [] }, "Bob's Data");
            expect(writer.hasPivotTable("[a.xlsx]Bob's Data!PT")).to.equal(true);
            expect(writer.hasPivotTable("'[a.xlsx]Bob''s Data'!PT")).to.equal(true);
            expect(writer.hasPivotTable("'Bob''s Data'!PT")).to.equal(true);
            expect(writer.hasPivotTable("'Bob's Data'!PT")).to.equal(false);
            expect(writer.hasPivotTable("[a.xlsx]Bob's Data!pt")).to.equal(false);
        });

        it("writes a chart that is not well-formed, or has no pivot source, as read", () => {
            const { chartPartBytes } = verquire("xlsx/pivot-chart");
            const bytes = (text) => new TextEncoder().encode(text);
            for (const text of [
                pivotChartXml().replace("</c:chartSpace>", ""),
                pivotChartXml({ pivot: false }),
                pivotChartXml().replace(/c:chartSpace/g, "c:other"),
                "<c:pivotSource",
            ]) {
                const data = bytes(text);
                expect(chartPartBytes(data, undefined)).to.equal(data);
            }
            // whatever the prefix of the namespace of the chart
            const unprefixed = pivotChartXml({ extensions: [PIVOT_CHART_EXTENSIONS.pivotOptions] })
                .replace(/<c:/g, "<")
                .replace(/<\/c:/g, "</")
                .replace(`xmlns:c="${NS.c}"`, `xmlns="${NS.c}"`);
            const written = new TextDecoder().decode(chartPartBytes(bytes(unprefixed), undefined));
            expect(written).to.not.match(/pivotSource|pivotFmts|pivotOptions/);
            expect(written).to.include("<plotArea>");
        });

        it("strips the pivot elements of a large chart in linear time", async function () {
            this.timeout(20000);
            const count = 50000;
            const formats = `<c:pivotFmts>${'<c:pivotFmt><c:idx val="0"/></c:pivotFmt>'.repeat(count)}</c:pivotFmts>`;
            const chartXml = pivotChartXml({ source: "[Book1.xlsx]Pivot!Missing" }).replace(/<c:pivotFmts>.*<\/c:pivotFmts>/, formats);
            const start = Date.now();
            const { buffer } = await roundTrip(await pivotWorkbook({ chart: true, chartXml }));
            expect(Date.now() - start).to.be.below(5000);
            expect(writtenChart(buffer, "large")).to.equal(pivotChartXml({ pivot: false }));
        });
    });

    it("drops slicers and slicer caches, leaving no reference to them", async () => {
        const slicerExt = `<extLst><ext uri="{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}" xmlns:x14="${NS.x14}"><x14:pivotTableDefinition hideValuesRow="1"/></ext></extLst>`;
        const buffer = await pivotWorkbook({
            edit(parts) {
                parts["xl/pivotTables/pivotTable1.xml"] = pivotTableXml({ extLst: slicerExt });
                parts["xl/worksheets/sheet2.xml"] = parts["xl/worksheets/sheet2.xml"].replace(
                    "</worksheet>",
                    `<extLst><ext uri="{A8765BA9-456A-4dab-B4F3-ACF838C121DE}" xmlns:x14="${NS.x14}"><x14:slicerList><x14:slicer r:id="rId2"/></x14:slicerList></ext></extLst></worksheet>`
                );
                parts["xl/worksheets/_rels/sheet2.xml.rels"] = rels([
                    ["rId1", REL.pivotTable, "../pivotTables/pivotTable1.xml"],
                    ["rId2", REL.slicer, "../slicers/slicer1.xml"],
                ]);
                parts["xl/slicers/slicer1.xml"] =
                    `${XML_DECL}<slicers xmlns="${NS.x14}"><slicer name="Region" cache="Slicer_Region" caption="Region" rowHeight="241300"/></slicers>`;
                parts["xl/slicerCaches/slicerCache1.xml"] =
                    `${XML_DECL}<slicerCacheDefinition xmlns="${NS.x14}" name="Slicer_Region" sourceName="Region"><pivotTables><pivotTable tabId="2" name="PivotTable1"/></pivotTables><data><tabular pivotCacheId="1"><items count="2"><i x="0" s="1"/><i x="1" s="1"/></items></tabular></data></slicerCacheDefinition>`;
                parts["xl/workbook.xml"] = parts["xl/workbook.xml"].replace(
                    "</workbook>",
                    `<extLst><ext uri="{BBE1A952-AA13-448e-AADC-164F8A28A991}" xmlns:x14="${NS.x14}"><x14:slicerCaches><x14:slicerCache r:id="rId98"/></x14:slicerCaches></ext></extLst></workbook>`
                );
                parts["xl/_rels/workbook.xml.rels"] = parts["xl/_rels/workbook.xml.rels"].replace(
                    "</Relationships>",
                    `<Relationship Id="rId98" Type="${REL.slicerCache}" Target="slicerCaches/slicerCache1.xml"/></Relationships>`
                );
            },
        });
        const { buffer: written } = await roundTrip(buffer);
        const { parts, text } = expectPivotTable(written, { tableXml: pivotTableXml({ extLst: slicerExt }) });
        expect(Object.keys(parts).filter((name) => /slicer/i.test(name))).to.deep.equal([]);
        for (const name of Object.keys(parts)) {
            // (xl/styles.xml: the default slicer style of the styles written by ExcelJS)
            if (!name.startsWith("xl/pivotTables/pivotTable") && name !== "xl/styles.xml") {
                expect(text(name), name).to.not.match(/slicer/i);
            }
        }
    });

    describe("hostile or broken files: the pivot table is dropped, the file still loads and writes", () => {
        const cases = {
            "a missing cache definition part": (parts) => {
                delete parts["xl/pivotCache/pivotCacheDefinition1.xml"];
            },
            "a missing records part": (parts) => {
                delete parts["xl/pivotCache/pivotCacheRecords1.xml"];
            },
            "a missing pivot table part": (parts) => {
                delete parts["xl/pivotTables/pivotTable1.xml"];
            },
            "no <pivotCache> in xl/workbook.xml": (parts) => {
                parts["xl/workbook.xml"] = parts["xl/workbook.xml"].replace(/<pivotCaches>.*<\/pivotCaches>/, "");
            },
            "a cacheId that does not match the <pivotCache> of its cache": (parts) => {
                parts["xl/pivotTables/pivotTable1.xml"] = pivotTableXml({ cacheId: 6 });
            },
            "a cacheId listed twice in xl/workbook.xml": (parts) => {
                parts["xl/workbook.xml"] = parts["xl/workbook.xml"].replace(
                    "</pivotCaches>",
                    '<pivotCache cacheId="5" r:id="rId97"/></pivotCaches>'
                );
                parts["xl/_rels/workbook.xml.rels"] = parts["xl/_rels/workbook.xml.rels"].replace(
                    "</Relationships>",
                    `<Relationship Id="rId97" Type="${REL.pivotCacheDefinition}" Target="pivotCache/pivotCacheDefinition2.xml"/></Relationships>`
                );
                parts["xl/pivotCache/pivotCacheDefinition2.xml"] = CACHE_DEFINITION_XML;
            },
            "a cache listed twice with two cacheIds": (parts) => {
                parts["xl/workbook.xml"] = parts["xl/workbook.xml"].replace(
                    "</pivotCaches>",
                    '<pivotCache cacheId="6" r:id="rId99"/></pivotCaches>'
                );
            },
            "a cacheId that is not an unsigned int": (parts) => {
                parts["xl/workbook.xml"] = parts["xl/workbook.xml"].replace('cacheId="5"', 'cacheId="-5"');
                parts["xl/pivotTables/pivotTable1.xml"] = pivotTableXml({ cacheId: -5 });
            },
            "a pivot table target escaping the package": (parts) => {
                parts["xl/worksheets/_rels/sheet2.xml.rels"] = rels([["rId1", REL.pivotTable, "../../../../pivotTables/pivotTable1.xml"]]);
            },
            "a cache target with a scheme": (parts) => {
                parts["xl/pivotTables/_rels/pivotTable1.xml.rels"] = rels([
                    ["rId1", REL.pivotCacheDefinition, "file:///xl/pivotCache/pivotCacheDefinition1.xml"],
                ]);
            },
            "a records target that is a relationships part": (parts) => {
                parts["xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels"] = rels([
                    ["rId1", REL.pivotCacheRecords, "_rels/pivotCacheDefinition1.xml.rels"],
                ]);
            },
            "a cycle: the records of the cache are the cache": (parts) => {
                parts["xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels"] = rels([
                    ["rId1", REL.pivotCacheRecords, "pivotCacheDefinition1.xml"],
                ]);
            },
            "a cycle: the records of the cache are the pivot table": (parts) => {
                parts["xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels"] = rels([
                    ["rId1", REL.pivotCacheRecords, "../pivotTables/pivotTable1.xml"],
                ]);
            },
            "a cycle: the cache of the pivot table is the pivot table": (parts) => {
                parts["xl/pivotTables/_rels/pivotTable1.xml.rels"] = rels([["rId1", REL.pivotCacheDefinition, "pivotTable1.xml"]]);
            },
            "records with relationships of their own": (parts) => {
                parts["xl/pivotCache/_rels/pivotCacheRecords1.xml.rels"] = rels([
                    ["rId1", REL.pivotCacheDefinition, "pivotCacheDefinition1.xml"],
                ]);
            },
            "a pivot table with two caches": (parts) => {
                parts["xl/pivotTables/_rels/pivotTable1.xml.rels"] = rels([
                    ["rId1", REL.pivotCacheDefinition, "../pivotCache/pivotCacheDefinition1.xml"],
                    ["rId2", REL.pivotCacheDefinition, "../pivotCache/pivotCacheDefinition1.xml"],
                ]);
            },
            "an unknown relationship of the pivot table": (parts) => {
                parts["xl/pivotTables/_rels/pivotTable1.xml.rels"] = rels([
                    ["rId1", REL.pivotCacheDefinition, "../pivotCache/pivotCacheDefinition1.xml"],
                    ["rId2", REL.chart, "../../docProps/app.xml"],
                ]);
            },
            "a duplicate relationship id": (parts) => {
                parts["xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels"] = rels([
                    ["rId1", REL.pivotCacheRecords, "pivotCacheRecords1.xml"],
                    ["rId1", REL.externalLinkPath, "x.xlsx", true],
                ]);
            },
            "a malformed relationships part": (parts) => {
                parts["xl/pivotTables/_rels/pivotTable1.xml.rels"] = "<Relationships><Relationship";
            },
            "a malformed pivot table": (parts) => {
                parts["xl/pivotTables/pivotTable1.xml"] = "<pivotTableDefinition cacheId='5'><location></pivot";
            },
            "a pivot table that is not a <pivotTableDefinition>": (parts) => {
                parts["xl/pivotTables/pivotTable1.xml"] = `${XML_DECL}<worksheet cacheId="5"/>`;
            },
            "a cache that is not a <pivotCacheDefinition>": (parts) => {
                parts["xl/pivotCache/pivotCacheDefinition1.xml"] =
                    `${XML_DECL}<pivotTableDefinition><cacheSource type="worksheet"/></pivotTableDefinition>`;
            },
            "a cache whose source is an external connection": (parts) => {
                parts["xl/pivotCache/pivotCacheDefinition1.xml"] = cacheDefinitionXml({
                    source: '<cacheSource type="external" connectionId="1"/>',
                });
            },
            "a cache with a connectionId": (parts) => {
                parts["xl/pivotCache/pivotCacheDefinition1.xml"] = cacheDefinitionXml({
                    source: '<cacheSource type="worksheet" connectionId="1"><worksheetSource ref="A1:C4" sheet="Data"/></cacheSource>',
                });
            },
            "a workbook relationship of the cache that is external": (parts) => {
                parts["xl/_rels/workbook.xml.rels"] = parts["xl/_rels/workbook.xml.rels"].replace(
                    'Target="pivotCache/pivotCacheDefinition1.xml"',
                    'Target="pivotCache/pivotCacheDefinition1.xml" TargetMode="External"'
                );
            },
        };
        for (const [name, edit] of Object.entries(cases)) {
            it(name, async () => {
                const { wb, buffer } = await roundTrip(await pivotWorkbook({ edit }));
                expect(wb.getWorksheet("Pivot").getCell("B6").value).to.equal(60);
                expectNoPivotTable(buffer);
                // the file written reads again
                await roundTrip(buffer);
            });
        }

        it("a pivot table referenced by two worksheets belongs to the first one", async () => {
            const buffer = await pivotWorkbook({
                edit(parts) {
                    parts["xl/worksheets/_rels/sheet1.xml.rels"] = rels([["rId1", REL.pivotTable, "../pivotTables/pivotTable1.xml"]]);
                },
            });
            const { buffer: written } = await roundTrip(buffer);
            const { text } = expectPivotTable(written, { sheet: "xl/worksheets/sheet1.xml" });
            expect(text("xl/worksheets/_rels/sheet2.xml.rels") ?? "").to.not.include("pivot");
            expect(pivotPartNames(written).filter((name) => /pivotTable\d+\.xml$/.test(name))).to.have.length(1);
        });

        it("bounds the number of parts read", async () => {
            const count = 30;
            const buffer = await pivotWorkbook({
                edit(parts) {
                    const sheetRels = [];
                    for (let i = 1; i <= count; i++) {
                        sheetRels.push([`rId${i}`, REL.pivotTable, `../pivotTables/pivotTable${i}.xml`]);
                        parts[`xl/pivotTables/pivotTable${i}.xml`] = pivotTableXml({ name: `PivotTable${i}` });
                        parts[`xl/pivotTables/_rels/pivotTable${i}.xml.rels`] = parts["xl/pivotTables/_rels/pivotTable1.xml.rels"];
                    }
                    parts["xl/worksheets/_rels/sheet2.xml.rels"] = rels(sheetRels);
                },
            });
            // all of them by default
            const { buffer: written } = await roundTrip(buffer);
            expect(pivotPartNames(written).filter((name) => /pivotTable\d+\.xml$/.test(name))).to.have.length(count);

            // with a bound of 10 parts: the cache and its records (2 parts) and 8 pivot tables
            const { readZip } = verquire("utils/zip-reader");
            const entries = new Map(readZip(new Uint8Array(buffer)).map((entry) => [entry.name, entry]));
            const text = (name) => strFromU8(unzipParts(buffer)[name]);
            const parseRelsPart = (name) =>
                parseRels(text(name)).map((rel) => ({ ...rel, Target: rel.Target, Type: rel.Type, Id: rel.Id }));
            const tables = await loadPivotParts(
                { 2: parseRelsPart("xl/worksheets/_rels/sheet2.xml.rels") },
                [{ cacheId: "5", rId: "rId99" }],
                parseRelsPart("xl/_rels/workbook.xml.rels"),
                entries,
                10
            );
            expect(tables.size).to.equal(8);
            for (const table of tables.values()) {
                expect(table.cache).to.equal([...tables.values()][0].cache);
            }
        });

        it("reads many pivot table relationships in linear time", async function () {
            this.timeout(20000);
            const count = 50000;
            const sheetRels = [];
            for (let i = 1; i <= count; i++) {
                sheetRels.push([`rId${i}`, REL.pivotTable, `../pivotTables/missing${i}.xml`]);
            }
            const buffer = await pivotWorkbook({
                edit(parts) {
                    parts["xl/worksheets/_rels/sheet2.xml.rels"] = rels(sheetRels);
                    parts["xl/workbook.xml"] = parts["xl/workbook.xml"].replace(
                        "</pivotCaches>",
                        `${'<pivotCache cacheId="7" r:id="rId99"/>'.repeat(count)}</pivotCaches>`
                    );
                },
            });
            const start = Date.now();
            const { buffer: written } = await roundTrip(buffer);
            expect(Date.now() - start).to.be.below(5000);
            expectNoPivotTable(written);
        });

        it("checks a cacheId with an anchored pattern (10k digits: no backtracking)", async () => {
            const digits = "9".repeat(10000);
            const buffer = await pivotWorkbook({
                edit(parts) {
                    parts["xl/workbook.xml"] = parts["xl/workbook.xml"].replace('cacheId="5"', `cacheId="${digits}"`);
                    parts["xl/pivotTables/pivotTable1.xml"] = pivotTableXml({ cacheId: digits });
                },
            });
            const start = Date.now();
            const { buffer: written } = await roundTrip(buffer);
            expect(Date.now() - start).to.be.below(2000);
            expectNoPivotTable(written);
        });
    });

    it("does not write a pivot table whose cache has the cacheId of another cache (two files in one workbook)", () => {
        const cache = (data) => ({ kind: "pivotCacheDefinition", cacheId: 5, data, rels: [] });
        const table = (c) => ({
            kind: "pivotTable",
            name: "PivotTable1",
            data: new Uint8Array([1]),
            cache: c,
            rels: [{ Id: "rId1", Type: REL.pivotCacheDefinition, part: c }],
        });
        const writer = new PivotPartWriter(2);
        const first = cache(new Uint8Array([2]));
        expect(writer.addTable(table(first), "Sheet1")).to.equal("xl/pivotTables/pivotTable2.xml");
        expect(writer.addTable(table(first), "Sheet2")).to.equal("xl/pivotTables/pivotTable3.xml");
        expect(writer.addTable(table(cache(new Uint8Array([3]))), "Sheet3")).to.equal(undefined);
        expect(writer.caches).to.deep.equal([{ name: "xl/pivotCache/pivotCacheDefinition2.xml", cacheId: 5 }]);
        expect(writer.freeCacheId(5)).to.equal(6);
        expect(writer.freeCacheId(10)).to.equal(10);
    });

    it("reads a file with pivot tables with the streaming reader", async () => {
        const buffer = await pivotWorkbook({ chart: true });
        const { Readable } = require("stream");
        const reader = new ExcelJS.stream.xlsx.WorkbookReader(Readable.from([buffer]), {});
        const values = [];
        for await (const worksheetReader of reader) {
            for await (const row of worksheetReader) {
                values.push(row.values.slice(1));
            }
        }
        expect(values).to.deep.include(["Grand Total", 60]);
    });

    it("writes a workbook without pivot tables as before", async () => {
        const base = await baseWorkbookBuffer();
        const { wb, buffer } = await roundTrip(base);
        expect(pivotPartNames(buffer)).to.deep.equal([]);
        expect("pivotTableParts" in wb.getWorksheet("Sheet1").model).to.equal(false);
    });
});
