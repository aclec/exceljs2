// The chartsheets of a workbook (sheets that are a single chart: `xl/chartsheets/sheetN.xml` → its drawing → the chart),
// carried from the file read to the file written, in their place among the worksheets: read (loadChartsheets), kept by
// the workbook model (reconcileChartsheets), and written again (prepareChartsheets, prepareChartsheetDrawings, then
// writeChartsheetParts and writeDrawingParts). A chartsheet created by `workbook.addChartsheet` (or by the streaming
// writer) is written by the same code: its part is createdChartsheetPart, its chart generated (src/doc/chart.ts).
//
// Security (the parts come from an untrusted file):
// - the targets of the relationships (workbook → chartsheet → drawing) are resolved to entries of the zip by
//   chart-parts.ts resolvePartName (no scheme, no escape above the root...); only a `<sheet>` whose relationship has the
//   chartsheet type is read as a chartsheet, only a relationship of the drawing type to an `xl/drawings/<name>.xml` part
//   read as a drawing is followed; external relationships are not followed;
// - each chartsheet part is read once (a part referenced by several `<sheet>` elements belongs to the first one), as is
//   each drawing (claimed by one sheet: worksheet or chartsheet), and the number of chartsheets is bounded by
//   MAX_CHARTSHEETS; their charts go through chart-parts.ts (bounded number of parts, a chart claimed by one anchor);
// - the chartsheet XML is carried as parsed (chartsheet-xform.ts): only the relationship id of its drawing is kept,
//   replaced on write; the elements referencing other parts are left out; the part names written are chosen by the
//   writer, never taken from the file;
// - a chartsheet that cannot be carried whole (a missing or refused part, no chart, malformed XML, a name already used
//   by another sheet) is dropped, as every chartsheet was before; the file still loads;
// - maps keyed by part names or relationship ids are Maps.

import RelationshipsXform from "./xform/core/relationships-xform";
import ChartsheetXform from "./xform/chartsheet/chartsheet-xform";
import { ownEntry } from "./xform/base-xform";
import { prepareChartAnchors, relsPartName, resolvePartName, type ZipEntries } from "./chart-parts";
import type { ZipWriter } from "../utils/zip-stream";
import type {
    ChartPartRegistry,
    ChartXmlElement,
    ChartsheetModel,
    ChartsheetPartModel,
    ChartsheetWriteModel,
    DrawingModel,
    LoadedChartsheetModel,
    RelationshipModel,
    SheetDefModel,
    WorkbookViewModel,
} from "../types";

const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/**
 * The part of a chartsheet created (`workbook.addChartsheet`), as a chartsheet read is carried: CT_Chartsheet (sheetPr,
 * sheetViews, pageMargins, drawing), as Excel writes a new chartsheet: zoomed to fit the window, the margins of a new
 * sheet. Its drawing (`<drawing r:id>`, renumbered on write) holds the chart, `xdr:absoluteAnchor`.
 */
export function createdChartsheetPart(): ChartsheetPartModel {
    const element = (
        name: string,
        attributes: ChartXmlElement["attributes"] = [],
        children: ChartXmlElement["children"] = []
    ): ChartXmlElement => ({ name, attributes, children });
    const drawingRelId = "rId1";
    const sheet = element(
        "chartsheet",
        [
            { name: "xmlns", value: MAIN_NS },
            { name: "xmlns:r", value: RELATIONSHIPS_NS },
        ],
        [
            element("sheetPr"),
            element(
                "sheetViews",
                [],
                [
                    element("sheetView", [
                        { name: "zoomToFit", value: "1" },
                        { name: "workbookViewId", value: "0" },
                    ]),
                ]
            ),
            element("pageMargins", [
                { name: "left", value: "0.7" },
                { name: "right", value: "0.7" },
                { name: "top", value: "0.75" },
                { name: "bottom", value: "0.75" },
                { name: "header", value: "0.3" },
                { name: "footer", value: "0.3" },
            ]),
            element("drawing", [{ name: "r:id", value: drawingRelId, role: "rel" }]),
        ]
    );
    return { sheet, drawingRelId };
}

// the local name of an element (after its prefix)
function localName(name: string): string {
    return name.slice(name.indexOf(":") + 1);
}

/**
 * The `<chartsheet>` element with the `tabSelected` attribute of each of its `<sheetView>` elements set from `selected`
 * (by the `workbookViewId` of the view, 0 when absent or not a number): written `tabSelected="1"` when the chartsheet is
 * the active tab of that workbook view, left out otherwise, so that Excel does not open several tabs grouped.
 *
 * @returns `sheet` itself when no view changes; otherwise a copy of the elements on the path to the views changed (the
 * other elements are shared)
 */
export function withTabSelected(sheet: ChartXmlElement, selected: (workbookViewId: number) => boolean): ChartXmlElement {
    let changed = false;
    const children = sheet.children.map((child) => {
        if (typeof child === "string" || localName(child.name) !== "sheetViews") {
            return child;
        }
        let viewsChanged = false;
        const views = child.children.map((view) => {
            if (typeof view === "string" || localName(view.name) !== "sheetView") {
                return view;
            }
            const viewId = Number.parseInt(view.attributes.find((attribute) => attribute.name === "workbookViewId")?.value ?? "0", 10);
            const want = selected(Number.isNaN(viewId) ? 0 : viewId);
            const value = view.attributes.find((attribute) => attribute.name === "tabSelected")?.value;
            if (want === (value === "1" || value === "true")) {
                return view;
            }
            viewsChanged = true;
            const attributes = view.attributes.filter((attribute) => attribute.name !== "tabSelected");
            return { ...view, attributes: want ? [{ name: "tabSelected", value: "1" }, ...attributes] : attributes };
        });
        if (!viewsChanged) {
            return child;
        }
        changed = true;
        return { ...child, children: views };
    });
    return changed ? { ...sheet, children } : sheet;
}

/** The bound on the number of chartsheets carried from a file. */
export const MAX_CHARTSHEETS = 10000;

/** The relationship type of a chartsheet (from the workbook part). */
export const CHARTSHEET_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
const DRAWING_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
/** The content type of a chartsheet part. */
export const CHARTSHEET_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";

// a drawing part as load() reads it (xlsx.ts DRAWING_PART, anchored): its name is the key of the drawings read
const DRAWING_PART = /^xl\/drawings\/([a-zA-Z0-9]+)[.]xml$/;

// the relationships by id (the first one of an id: a duplicate is ignored)
function relsById(rels: readonly RelationshipModel[] | undefined): Map<string, RelationshipModel> {
    const byId = new Map<string, RelationshipModel>();
    for (const rel of rels ?? []) {
        if (!byId.has(rel.Id)) {
            byId.set(rel.Id, rel);
        }
    }
    return byId;
}

// the part name of the chartsheet of a <sheet> (an internal relationship of the chartsheet type), if any
function chartsheetPartName(sheet: SheetDefModel, rels: ReadonlyMap<string, RelationshipModel>): string | undefined {
    const rel = sheet.rId === undefined ? undefined : rels.get(sheet.rId);
    if (!rel || rel.Type !== CHARTSHEET_REL_TYPE || rel.TargetMode === "External" || typeof rel.Target !== "string") {
        return undefined;
    }
    return resolvePartName("xl/workbook.xml", rel.Target);
}

// a chartsheet part, or undefined for malformed XML or a part that chartsheet-xform refuses
async function parseChartsheet(data: Uint8Array): Promise<ChartsheetPartModel | undefined> {
    try {
        return (await new ChartsheetXform().parseStream([data])) ?? undefined;
    } catch {
        // malformed XML: the chartsheet is dropped (it was never read before), the file still loads
        return undefined;
    }
}

// the relationships of a part, or undefined for a malformed part
async function parseRels(data: Uint8Array): Promise<RelationshipModel[] | undefined> {
    try {
        return (await new RelationshipsXform().parseStream([data])) ?? [];
    } catch {
        // malformed XML: the chartsheet is dropped, the file still loads
        return undefined;
    }
}

/**
 * Reads the chartsheet parts that the `<sheet>` elements of the workbook reference (each part once), with the name of
 * their drawing.
 *
 * @param sheets - the `<sheet>` elements of `xl/workbook.xml`
 * @param workbookRels - the relationships of the workbook part
 * @param entries - the entries of the zip
 * @param maxChartsheets - the bound on the number of chartsheet parts read
 * @returns the chartsheets read (part name → chartsheet): those whose part and relationship to a drawing part are valid
 * @throws (rejects) when a decompression limit of the zip reader is exceeded
 */
export async function loadChartsheets(
    sheets: readonly SheetDefModel[] | undefined,
    workbookRels: readonly RelationshipModel[] | undefined,
    entries: ZipEntries,
    maxChartsheets = MAX_CHARTSHEETS
): Promise<Map<string, LoadedChartsheetModel>> {
    const rels = relsById(workbookRels);
    const tried = new Set<string>();
    const chartsheets = new Map<string, LoadedChartsheetModel>();
    for (const sheet of sheets ?? []) {
        const partName = chartsheetPartName(sheet, rels);
        if (partName === undefined || tried.has(partName) || tried.size >= maxChartsheets) {
            continue;
        }
        tried.add(partName);
        const entry = entries.get(partName);
        if (!entry) {
            continue;
        }
        const part = await parseChartsheet(await entry.data());
        const relsEntry = part && entries.get(relsPartName(partName));
        const partRels = relsEntry && (await parseRels(await relsEntry.data()));
        if (!part || !partRels) {
            continue;
        }
        const rel = relsById(partRels).get(part.drawingRelId);
        const target =
            rel && rel.Type === DRAWING_REL_TYPE && rel.TargetMode !== "External" ? resolvePartName(partName, rel.Target) : undefined;
        const drawing = target === undefined ? undefined : DRAWING_PART.exec(target)?.[1];
        if (drawing !== undefined) {
            chartsheets.set(partName, { ...part, drawing });
        }
    }
    return chartsheets;
}

/** What {@link reconcileChartsheets} reads from the read model of xlsx.ts. */
export interface ChartsheetReconcileModel {
    sheets?: readonly SheetDefModel[];
    workbookRels?: readonly RelationshipModel[];
    /** the worksheet of each `<sheet>` (by position; WorkbookXform.reconcile) */
    sheetWorksheets?: readonly (object | undefined)[];
    /** the names of the worksheets kept */
    worksheetNames?: readonly string[];
    chartsheetParts?: ReadonlyMap<string, LoadedChartsheetModel>;
    drawings: Readonly<Record<string, DrawingModel>>;
    views: WorkbookViewModel[];
}

/**
 * The chartsheets of a file read, in the order of their `<sheet>` elements: those whose part was read
 * ({@link loadChartsheets}), whose drawing holds a chart carried and is not claimed by another sheet, and whose name is
 * not the name of another sheet (case-insensitive, as Excel). Each gets its position among the sheets kept (worksheets and
 * chartsheets); the `activeTab` / `firstSheet` of the workbook views, which count every `<sheet>`, are moved to the same
 * sheets among those kept (unchanged when every sheet is kept).
 *
 * @param claim - claims a drawing for one sheet: `false` when it was claimed already
 */
export function reconcileChartsheets(model: ChartsheetReconcileModel, claim: (drawing: DrawingModel) => boolean): ChartsheetModel[] {
    const sheets = model.sheets ?? [];
    const rels = relsById(model.workbookRels);
    const names = new Set((model.worksheetNames ?? []).map((name) => name.toLowerCase()));
    const worksheetsSeen = new Set<object>();
    const usedParts = new Set<string>();
    const chartsheets: ChartsheetModel[] = [];
    // the position of each <sheet> among the sheets kept (a sheet not kept: the position of the next one kept)
    const positions: number[] = [];
    let kept = 0;
    sheets.forEach((sheet, i) => {
        positions.push(kept);
        const worksheet = model.sheetWorksheets?.[i];
        if (worksheet) {
            if (!worksheetsSeen.has(worksheet)) {
                worksheetsSeen.add(worksheet);
                kept++;
            }
            return;
        }
        const partName = chartsheetPartName(sheet, rels);
        const part = partName === undefined ? undefined : model.chartsheetParts?.get(partName);
        if (partName === undefined || !part || usedParts.has(partName) || chartsheets.length >= MAX_CHARTSHEETS) {
            return;
        }
        const drawing = ownEntry(model.drawings, part.drawing);
        const name = sheet.name.toLowerCase();
        if (!drawing || !drawing.charts || !drawing.charts.length || names.has(name) || !claim(drawing)) {
            return;
        }
        usedParts.add(partName);
        names.add(name);
        const chartsheet: ChartsheetModel = {
            id: sheet.id,
            name: sheet.name,
            index: kept++,
            sheet: part.sheet,
            drawingRelId: part.drawingRelId,
            charts: drawing.charts,
        };
        if (sheet.state) {
            chartsheet.state = sheet.state;
        }
        chartsheets.push(chartsheet);
    });

    if (kept < sheets.length) {
        const move = (index: number | undefined): number | undefined => {
            const position = index === undefined ? undefined : positions[index];
            return position === undefined ? index : Math.min(position, Math.max(kept - 1, 0));
        };
        for (const view of model.views) {
            if (view.activeTab !== undefined) {
                view.activeTab = move(view.activeTab);
            }
            if (view.firstSheet !== undefined) {
                view.firstSheet = move(view.firstSheet);
            }
        }
    }
    return chartsheets;
}

// an index of a model (any value from a hand-built model): not a finite number appends
function sheetIndex(sheet: { readonly index: number }): number {
    return Number.isFinite(sheet.index) ? sheet.index : Infinity;
}

/**
 * The sheets in the order they are written: the worksheets (in their order), each chartsheet inserted at its index (by
 * increasing index; an index beyond the end appends). Linear (after sorting the chartsheets).
 *
 * @returns `worksheets` itself when there are no chartsheets
 */
export function mergeSheets<W, C extends { readonly index: number }>(
    worksheets: readonly W[],
    chartsheets: readonly C[] | undefined
): readonly (W | C)[] {
    if (!chartsheets || !chartsheets.length) {
        return worksheets;
    }
    // a stable sort: chartsheets of the same index keep their order
    const sorted = [...chartsheets].sort((a, b) => sheetIndex(a) - sheetIndex(b) || 0);
    const sheets: (W | C)[] = [];
    let w = 0;
    let c = 0;
    while (w < worksheets.length || c < sorted.length) {
        const chartsheet = sorted[c];
        if (chartsheet && (w >= worksheets.length || sheetIndex(chartsheet) <= sheets.length)) {
            sheets.push(chartsheet);
            c++;
        } else {
            // w < worksheets.length here
            const worksheet = worksheets[w++];
            if (worksheet !== undefined) {
                sheets.push(worksheet);
            }
        }
    }
    return sheets;
}

/**
 * The chartsheets to write (xlsx.ts prepare): numbered from 1 in the order of their index, each with a `sheetId` unique
 * among the sheets written (its own, unless a worksheet or an earlier chartsheet has it, or it is not an integer from 1
 * to 65535: then the lowest free one). Their drawings are made by {@link prepareChartsheetDrawings}.
 *
 * @param worksheets - the worksheets written (their ids are kept)
 */
export function prepareChartsheets(
    worksheets: readonly { readonly id: number }[],
    chartsheets: readonly ChartsheetModel[] | undefined
): ChartsheetWriteModel[] {
    if (!chartsheets || !chartsheets.length) {
        return [];
    }
    const used = new Set(worksheets.map((worksheet) => worksheet.id));
    let nextId = 1;
    const sorted = [...chartsheets].sort((a, b) => sheetIndex(a) - sheetIndex(b) || 0);
    return sorted.map((chartsheet, i) => {
        let { id } = chartsheet;
        if (!Number.isInteger(id) || id < 1 || id > 65535 || used.has(id)) {
            while (used.has(nextId)) {
                nextId++;
            }
            id = nextId;
        }
        used.add(id);
        const written: ChartsheetWriteModel = {
            id,
            name: chartsheet.name,
            index: chartsheet.index,
            partNo: i + 1,
            sheet: chartsheet.sheet,
            drawingRelId: chartsheet.drawingRelId,
            drawingRId: "rId1",
            charts: chartsheet.charts,
            rels: [],
        };
        if (chartsheet.state) {
            written.state = chartsheet.state;
        }
        return written;
    });
}

/**
 * Gives each chartsheet written its drawing (`drawing<N>`, numbered after those of the worksheets) holding its chart
 * anchors, and the relationship of its part to that drawing.
 *
 * @param options - the drawings written so far, and their count (both updated)
 */
export function prepareChartsheetDrawings(
    chartsheets: readonly ChartsheetWriteModel[],
    options: { drawingsCount: number; drawings: DrawingModel[] },
    registry: ChartPartRegistry
): void {
    for (const chartsheet of chartsheets) {
        const name = `drawing${++options.drawingsCount}`;
        const drawing: DrawingModel = { name, anchors: [], rels: [] };
        drawing.chartAnchors = prepareChartAnchors(chartsheet.charts, name, (drawing.rels ??= []), registry);
        options.drawings.push(drawing);
        chartsheet.drawing = drawing;
        chartsheet.rels = [{ Id: chartsheet.drawingRId, Type: DRAWING_REL_TYPE, Target: `../drawings/${name}.xml` }];
    }
}

/** Adds the chartsheet parts written ({@link prepareChartsheets}, {@link prepareChartsheetDrawings}) and their relationships. */
export function writeChartsheetParts(zip: ZipWriter, chartsheets: readonly ChartsheetWriteModel[]): void {
    const chartsheetXform = new ChartsheetXform();
    const relsXform = new RelationshipsXform();
    for (const chartsheet of chartsheets) {
        zip.append(chartsheetXform.toXml(chartsheet), { name: `xl/chartsheets/sheet${chartsheet.partNo}.xml` });
        zip.append(relsXform.toXml(chartsheet.rels), { name: `xl/chartsheets/_rels/sheet${chartsheet.partNo}.xml.rels` });
    }
}
