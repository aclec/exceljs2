// The parts of the charts of the worksheets, carried opaque from the file read to the file written (no chart model, no
// authoring API): read (loadChartParts), attached to the chart anchors of the drawings (resolveChartAnchors), named and
// written again (ChartPartWriter).
//
// Security (the parts come from an untrusted file):
// - a relationship target is resolved against the part that holds it, to the name of an entry of the zip only: `../` is
//   normalized and may not climb above the root, a target with a scheme, a backslash, a query, a fragment, a percent
//   escape or an empty segment is refused, as are relationships parts and `[Content_Types].xml`;
// - only the relationship types of a chart are followed (chart, chartEx, chart style and colors, user shapes, theme
//   override, embedded package, OLE object, image); each part kind has a fixed content type and a part name chosen by the
//   writer: nothing from the file reaches a part name or a content type. The ids and the external targets of the
//   relationships are written back through XmlStream (encoded);
// - external relationships (TargetMode="External", e.g. the link to the source workbook of the chart data or a hyperlink)
//   are kept as they are, never followed: dropping them would leave the chart referencing a missing relationship id;
// - a chart is carried whole or not at all: when one of its parts is missing, refused, or beyond the bound on the number
//   of parts, its anchor is dropped (what happened to every chart before);
// - the graph is walked with a work list and a visited map (a cycle of relationships ends), the number of parts is bounded
//   by MAX_CHART_PARTS, their size by the limits of the zip reader (each part is inflated once, through them);
// - maps keyed by part names or relationship ids are Maps.

import type { ZipFileEntry } from "../utils/zip-reader";
import type { ZipWriter } from "../utils/zip-stream";
import RelationshipsXform from "./xform/core/relationships-xform";
import DrawingXform from "./xform/drawing/drawing-xform";
import { anchorRelIds } from "./xform/drawing/chart-anchor";
import { chartPartBytes, type PivotSourceLookup } from "./pivot-chart";
import type {
    ChartAnchorModel,
    ChartPartKind,
    ChartPartModel,
    ChartPartRegistry,
    ChartRelModel,
    ChartXmlElement,
    ContentTypeOverride,
    DrawingChartModel,
    DrawingModel,
    RelationshipModel,
} from "../types";

/** The bound on the number of chart parts carried from a file (a chart has 1 to 4 parts, usually). */
export const MAX_CHART_PARTS = 10000;

interface KindInfo {
    /** the relationship type */
    readonly type: string;
    /** the directory and the base of the part name written: `xl/<dir>/<base><n>.<extension>` */
    readonly dir: string;
    readonly base: string;
    /** the content type of an XML part, or by extension (lower case) */
    readonly contentType: string | Readonly<Record<string, string>>;
}

const KINDS: Readonly<Record<ChartPartKind, KindInfo>> = {
    chart: {
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
        dir: "charts",
        base: "chart",
        contentType: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
    },
    chartEx: {
        type: "http://schemas.microsoft.com/office/2014/relationships/chartEx",
        dir: "charts",
        base: "chartEx",
        contentType: "application/vnd.ms-office.chartex+xml",
    },
    chartStyle: {
        type: "http://schemas.microsoft.com/office/2011/relationships/chartStyle",
        dir: "charts",
        base: "style",
        contentType: "application/vnd.ms-office.chartstyle+xml",
    },
    chartColorStyle: {
        type: "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle",
        dir: "charts",
        base: "colors",
        contentType: "application/vnd.ms-office.chartcolorstyle+xml",
    },
    chartUserShapes: {
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes",
        dir: "charts",
        base: "userShapes",
        contentType: "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml",
    },
    themeOverride: {
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride",
        dir: "charts",
        base: "themeOverride",
        contentType: "application/vnd.openxmlformats-officedocument.themeOverride+xml",
    },
    package: {
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
        dir: "embeddings",
        base: "Microsoft_Excel_Worksheet",
        contentType: {
            xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            xlsm: "application/vnd.ms-excel.sheet.macroEnabled.12",
            docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        },
    },
    oleObject: {
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
        dir: "embeddings",
        base: "oleObject",
        contentType: { bin: "application/vnd.openxmlformats-officedocument.oleObject", xls: "application/vnd.ms-excel" },
    },
    image: {
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        // not xl/media: the reader takes those for workbook images
        dir: "charts",
        base: "image",
        contentType: {
            png: "image/png",
            jpeg: "image/jpeg",
            jpg: "image/jpeg",
            gif: "image/gif",
            bmp: "image/bmp",
            tif: "image/tiff",
            tiff: "image/tiff",
            emf: "image/x-emf",
            wmf: "image/x-wmf",
            svg: "image/svg+xml",
        },
    },
};

const KIND_NAMES: readonly ChartPartKind[] = [
    "chart",
    "chartEx",
    "chartStyle",
    "chartColorStyle",
    "chartUserShapes",
    "themeOverride",
    "package",
    "oleObject",
    "image",
];
const KIND_BY_TYPE: ReadonlyMap<string, ChartPartKind> = new Map(KIND_NAMES.map((kind) => [KINDS[kind].type, kind]));

/** The kinds a drawing anchor holds: a chart anchor carries at least one. */
const CHART_KINDS: ReadonlySet<ChartPartKind> = new Set(["chart", "chartEx"]);

// a relationship id as written back (xsd:ID, ASCII): anchored, a single character class (linear)
const REL_ID = /^[A-Za-z_][A-Za-z0-9_.-]*$/;
// a character refused in an internal target
const BAD_TARGET_CHAR = /[:\\?#%]/;

/** The content type of a part of a kind with an extension (lower case), or undefined when the extension is refused. */
function contentTypeOf(kind: ChartPartKind, extension: string): string | undefined {
    const { contentType } = KINDS[kind];
    if (typeof contentType === "string") {
        return contentType;
    }
    return Object.hasOwn(contentType, extension) ? contentType[extension] : undefined;
}

/**
 * Resolves the target of an internal relationship of a part to a part name (OPC: relative to the directory of the
 * source part, or from the root with a leading `/`).
 *
 * @returns the part name (no leading `/`), or undefined for a target refused (see the header of this module)
 */
export function resolvePartName(sourcePart: string, target: string): string | undefined {
    if (!target || BAD_TARGET_CHAR.test(target)) {
        return undefined;
    }
    const segments = target.startsWith("/") ? [] : sourcePart.split("/").slice(0, -1);
    const relative = target.startsWith("/") ? target.slice(1) : target;
    for (const segment of relative.split("/")) {
        if (segment === "..") {
            if (!segments.length) {
                return undefined;
            }
            segments.pop();
        } else if (segment === "") {
            return undefined;
        } else if (segment !== ".") {
            segments.push(segment);
        }
    }
    const name = segments.join("/");
    if (!name || name === "[Content_Types].xml" || segments.includes("_rels")) {
        return undefined;
    }
    return name;
}

/** The relative target of a relationship from a part to another (both part names, no leading `/`). */
export function relativeTarget(sourcePart: string, targetPart: string): string {
    const from = sourcePart.split("/").slice(0, -1);
    const to = targetPart.split("/");
    let common = 0;
    while (common < from.length && common < to.length - 1 && from[common] === to[common]) {
        common++;
    }
    return "../".repeat(from.length - common) + to.slice(common).join("/");
}

/** The relationships part of a part: `dir/_rels/name.rels`. */
export function relsPartName(partName: string): string {
    const slash = partName.lastIndexOf("/");
    return `${partName.slice(0, slash + 1)}_rels/${partName.slice(slash + 1)}.rels`;
}

// the extension of a part name, in lower case ("" without one)
function extensionOf(partName: string): string {
    const base = partName.slice(partName.lastIndexOf("/") + 1);
    const dot = base.lastIndexOf(".");
    return dot === -1 ? "" : base.slice(dot + 1).toLowerCase();
}

/** A part while the graph is read: refused (`bad`), or its bytes and its relationships (to parts by name, or external). */
interface PartNode {
    readonly kind: ChartPartKind;
    bad: boolean;
    extension: string;
    data: Uint8Array;
    readonly rels: { readonly Id: string; readonly Type: string; readonly target?: string; readonly externalTarget?: string }[];
}

/** The zip entries of a file read, by name (without a leading `/`). */
export type ZipEntries = ReadonlyMap<string, Pick<ZipFileEntry, "data">>;

/** The relationships of a drawing by id (the first one of an id: a duplicate is ignored). */
function relsById(rels: readonly RelationshipModel[] | undefined): Map<string, RelationshipModel> {
    const byId = new Map<string, RelationshipModel>();
    for (const rel of rels ?? []) {
        if (!byId.has(rel.Id)) {
            byId.set(rel.Id, rel);
        }
    }
    return byId;
}

// the part name of a drawing read (drawings are keyed by the name of their part, `drawing1`)
function drawingPartName(name: string): string {
    return `xl/drawings/${name}.xml`;
}

/**
 * Reads the parts of the charts of the drawings (those that their chart anchors reference, and the parts these reference),
 * inflating each entry once.
 *
 * @param drawings - the drawings read (name → model), with their `chartFrames`
 * @param drawingRels - the relationships of the drawings (name → relationships)
 * @param entries - the entries of the zip
 * @param maxParts - the bound on the number of parts read
 * @returns the parts carried (part name → part), complete: each one, and every part it references, was read and accepted
 * @throws (rejects) when a decompression limit of the zip reader is exceeded
 */
export async function loadChartParts(
    drawings: Readonly<Record<string, DrawingModel>>,
    drawingRels: Readonly<Record<string, RelationshipModel[]>>,
    entries: ZipEntries,
    maxParts = MAX_CHART_PARTS
): Promise<Map<string, ChartPartModel>> {
    const nodes = new Map<string, PartNode>();
    // the parts to read (name and kind), and the parts that reference each part (to propagate a refusal)
    const queue: { readonly name: string; readonly kind: ChartPartKind }[] = [];
    const referrers = new Map<string, string[]>();

    for (const [name, drawing] of Object.entries(drawings)) {
        const rels = relsById(Object.hasOwn(drawingRels, name) ? drawingRels[name] : undefined);
        for (const frame of drawing.chartFrames ?? []) {
            for (const id of anchorRelIds(frame)) {
                const rel = rels.get(id);
                const kind = rel && rel.TargetMode !== "External" ? KIND_BY_TYPE.get(rel.Type) : undefined;
                const target = rel && kind ? resolvePartName(drawingPartName(name), rel.Target) : undefined;
                if (kind && target) {
                    queue.push({ name: target, kind });
                }
            }
        }
    }

    for (let i = 0; i < queue.length; i++) {
        const item = queue[i];
        if (!item) break;
        const { name, kind } = item;
        const known = nodes.get(name);
        if (known) {
            // a part reached as two different kinds is refused
            if (known.kind !== kind) {
                known.bad = true;
            }
            continue;
        }
        const node: PartNode = { kind, bad: true, extension: "", data: new Uint8Array(0), rels: [] };
        nodes.set(name, node);
        const entry = entries.get(name);
        if (!entry || nodes.size > maxParts) {
            continue;
        }
        const extension = extensionOf(name);
        if (typeof KINDS[kind].contentType === "string") {
            node.extension = "xml";
        } else if (contentTypeOf(kind, extension)) {
            node.extension = extension;
        } else {
            continue;
        }
        node.data = await entry.data();

        const relsEntry = entries.get(relsPartName(name));
        if (relsEntry) {
            const relsData = await relsEntry.data();
            let rels: RelationshipModel[] | null | undefined;
            try {
                // a new xform per part: a failed parse leaves its state behind
                rels = await new RelationshipsXform().parseStream([relsData]);
            } catch {
                // a malformed relationships part: the chart is not carried (it was never read before), the file still loads
                continue;
            }
            const ids = new Set<string>();
            let refused = false;
            for (const rel of rels ?? []) {
                if (!REL_ID.test(rel.Id) || ids.has(rel.Id)) {
                    refused = true;
                    break;
                }
                ids.add(rel.Id);
                if (rel.TargetMode === "External") {
                    node.rels.push({ Id: rel.Id, Type: rel.Type, externalTarget: rel.Target });
                    continue;
                }
                const relKind = KIND_BY_TYPE.get(rel.Type);
                const target = relKind ? resolvePartName(name, rel.Target) : undefined;
                if (!relKind || !target || !entries.has(target)) {
                    refused = true;
                    break;
                }
                node.rels.push({ Id: rel.Id, Type: rel.Type, target });
                queue.push({ name: target, kind: relKind });
                const sources = referrers.get(target);
                if (sources) {
                    sources.push(name);
                } else {
                    referrers.set(target, [name]);
                }
            }
            if (refused) {
                continue;
            }
        }
        node.bad = false;
    }

    // a part that references a refused part (directly or not) is refused
    const refusedParts = [...nodes].filter(([, node]) => node.bad).map(([name]) => name);
    for (let name = refusedParts.pop(); name !== undefined; name = refusedParts.pop()) {
        for (const source of referrers.get(name) ?? []) {
            const node = nodes.get(source);
            if (node && !node.bad) {
                node.bad = true;
                refusedParts.push(source);
            }
        }
    }

    // the models (created first: the relationships may form a cycle)
    const parts = new Map<string, ChartPartModel & { rels: ChartRelModel[] }>();
    for (const [name, node] of nodes) {
        if (!node.bad) {
            parts.set(name, { kind: node.kind, extension: node.extension, data: node.data, rels: [] });
        }
    }
    for (const [name, part] of parts) {
        for (const rel of nodes.get(name)?.rels ?? []) {
            const target = rel.target === undefined ? undefined : parts.get(rel.target);
            part.rels.push(
                target ? { Id: rel.Id, Type: rel.Type, part: target } : { Id: rel.Id, Type: rel.Type, externalTarget: rel.externalTarget }
            );
        }
    }
    return parts;
}

/**
 * The charts of a drawing: each chart anchor whose relationships all resolve (to a part carried, or external) and that
 * holds a chart not taken by another anchor.
 *
 * @param name - the name of the drawing (`drawing1`)
 * @param frames - its chart anchors, as parsed
 * @param rels - its relationships
 * @param parts - the parts carried ({@link loadChartParts})
 * @param claimed - the chart parts already taken (mutated): a chart belongs to one anchor
 */
export function resolveChartAnchors(
    name: string,
    frames: readonly ChartXmlElement[],
    rels: readonly RelationshipModel[] | undefined,
    parts: ReadonlyMap<string, ChartPartModel>,
    claimed: Set<ChartPartModel>
): ChartAnchorModel[] {
    const byId = relsById(rels);
    const charts: ChartAnchorModel[] = [];
    for (const anchor of frames) {
        const anchorRels: ChartRelModel[] = [];
        const chartParts: ChartPartModel[] = [];
        let complete = true;
        for (const id of anchorRelIds(anchor)) {
            const rel = byId.get(id);
            if (!rel) {
                complete = false;
                break;
            }
            if (rel.TargetMode === "External") {
                anchorRels.push({ Id: id, Type: rel.Type, externalTarget: rel.Target });
                continue;
            }
            const target = resolvePartName(drawingPartName(name), rel.Target);
            const part = target === undefined ? undefined : parts.get(target);
            if (!part || KINDS[part.kind].type !== rel.Type) {
                complete = false;
                break;
            }
            if (CHART_KINDS.has(part.kind)) {
                chartParts.push(part);
            }
            anchorRels.push({ Id: id, Type: rel.Type, part });
        }
        if (complete && chartParts.length && !chartParts.some((part) => claimed.has(part))) {
            for (const part of chartParts) {
                claimed.add(part);
            }
            charts.push({ anchor, rels: anchorRels });
        }
    }
    return charts;
}

/**
 * A copy of the graph of parts reachable from a part (the bytes are shared, never modified), for a chart written twice.
 *
 * @param copies - the copies made so far (original → copy): parts reachable twice, or in a cycle, are copied once
 */
function copyPartGraph(part: ChartPartModel, copies: Map<ChartPartModel, ChartPartModel & { rels: ChartRelModel[] }>): ChartPartModel {
    // the originals whose relationships are still to copy
    const pending: ChartPartModel[] = [];
    const copyOf = (original: ChartPartModel): ChartPartModel & { rels: ChartRelModel[] } => {
        let copy = copies.get(original);
        if (!copy) {
            copy = { kind: original.kind, extension: original.extension, data: original.data, rels: [] };
            copies.set(original, copy);
            pending.push(original);
        }
        return copy;
    };
    const result = copyOf(part);
    for (let original = pending.pop(); original; original = pending.pop()) {
        const copy = copyOf(original);
        for (const rel of original.rels) {
            copy.rels.push(rel.part ? { Id: rel.Id, Type: rel.Type, part: copyOf(rel.part) } : rel);
        }
    }
    return result;
}

/**
 * Adds the relationships of the charts of a worksheet to its drawing being written, and gives each chart anchor the new id
 * of each relationship id of the file.
 *
 * A chart belongs to one anchor of the workbook written: a chart already written for another anchor (a worksheet model
 * copied to another worksheet, which shares the charts of the original) is written again as a copy of its parts, under
 * new names, so that each worksheet has its own chart, as the copy of a worksheet in Excel does.
 *
 * @param drawingName - the name of the drawing written (`drawing1`)
 * @param drawingRels - its relationships so far (the pictures'): appended to
 * @returns the chart anchors to write
 */
export function prepareChartAnchors(
    charts: readonly ChartAnchorModel[],
    drawingName: string,
    drawingRels: RelationshipModel[],
    registry: ChartPartRegistry
): DrawingChartModel[] {
    const source = drawingPartName(drawingName);
    return charts.map(({ anchor, rels }) => {
        // every chart part of the anchor is claimed (no short cut): a copy is needed when one was claimed before
        let shared = false;
        for (const rel of rels) {
            if (rel.part && CHART_KINDS.has(rel.part.kind) && !registry.claim(rel.part)) {
                shared = true;
            }
        }
        let anchorRels = rels;
        if (shared) {
            const copies = new Map<ChartPartModel, ChartPartModel & { rels: ChartRelModel[] }>();
            anchorRels = rels.map((rel) => (rel.part ? { Id: rel.Id, Type: rel.Type, part: copyPartGraph(rel.part, copies) } : rel));
            for (const rel of anchorRels) {
                if (rel.part && CHART_KINDS.has(rel.part.kind)) {
                    registry.claim(rel.part);
                }
            }
        }
        const ids = new Map<string, string>();
        for (const rel of anchorRels) {
            const Id = `rId${drawingRels.length + 1}`;
            ids.set(rel.Id, Id);
            drawingRels.push(writtenRel(source, Id, rel, registry));
        }
        return { anchor, ids };
    });
}

// a relationship as written: to the part written (its kind's type), or external as read
function writtenRel(source: string, Id: string, rel: ChartRelModel, registry: ChartPartRegistry): RelationshipModel {
    if (rel.part) {
        return { Id, Type: KINDS[rel.part.kind].type, Target: relativeTarget(source, registry.partName(rel.part)) };
    }
    return { Id, Type: rel.Type, Target: rel.externalTarget ?? "", TargetMode: "External" };
}

/**
 * Names the chart parts of a workbook being written (`xl/charts/chart1.xml`, `xl/charts/style1.xml`, ...: numbered by
 * kind, in the order they are named), then writes them, with their relationships. A part shared by several charts is
 * written once.
 */
export class ChartPartWriter implements ChartPartRegistry {
    private readonly names = new Map<ChartPartModel, string>();
    private readonly counts = new Map<ChartPartKind, number>();
    private readonly parts: ChartPartModel[] = [];
    private readonly claimed = new Set<ChartPartModel>();

    /** @returns `false` when the chart part was claimed already (by another anchor) */
    claim(part: ChartPartModel): boolean {
        if (this.claimed.has(part)) {
            return false;
        }
        this.claimed.add(part);
        return true;
    }

    /** @returns the part name of a part, assigned on first call */
    partName(part: ChartPartModel): string {
        const known = this.names.get(part);
        if (known !== undefined) {
            return known;
        }
        const count = (this.counts.get(part.kind) ?? 0) + 1;
        this.counts.set(part.kind, count);
        const { dir, base } = KINDS[part.kind];
        const name = `xl/${dir}/${base}${count}.${part.extension}`;
        this.names.set(part, name);
        this.parts.push(part);
        return name;
    }

    /**
     * Names every part that the parts named so far reference (once every chart anchor is prepared).
     * @returns the content type of each part
     */
    complete(): ContentTypeOverride[] {
        // the list grows while it is walked: each part is named (and appended) once
        for (let i = 0; i < this.parts.length; i++) {
            for (const rel of this.parts[i]?.rels ?? []) {
                if (rel.part) {
                    this.partName(rel.part);
                }
            }
        }
        return this.parts.map((part) => ({
            name: this.partName(part),
            contentType: contentTypeOf(part.kind, part.extension) ?? "application/octet-stream",
        }));
    }

    /**
     * Adds the parts named (call {@link complete} first) and their relationships parts to the zip.
     *
     * @param pivotTables - the pivot tables written: a pivot chart whose pivot table is not is written as a static chart
     *   (pivot-chart.ts; without it, every pivot chart is)
     */
    write(zip: ZipWriter, pivotTables?: PivotSourceLookup): void {
        const relsXform = new RelationshipsXform();
        for (const part of this.parts) {
            const name = this.partName(part);
            zip.append(part.kind === "chart" ? chartPartBytes(part.data, pivotTables) : part.data, { name });
            if (part.rels.length) {
                const rels = part.rels.map((rel) => writtenRel(name, rel.Id, rel, this));
                zip.append(relsXform.toXml(rels), { name: relsPartName(name) });
            }
        }
    }
}

/**
 * Adds the drawing parts of the sheets written (worksheets and chartsheets: pictures and charts) and their relationships
 * to the zip.
 * @throws Error for a drawing without relationships (not prepared)
 */
export function writeDrawingParts(zip: ZipWriter, drawings: readonly DrawingModel[]): void {
    const drawingXform = new DrawingXform();
    const relsXform = new RelationshipsXform();
    for (const drawing of drawings) {
        drawingXform.prepare(drawing);
        zip.append(drawingXform.toXml(drawing), { name: `xl/drawings/${drawing.name}.xml` });
        // the prepare of the worksheet, or of the chartsheet, gives the drawing its relationships
        if (!drawing.rels) {
            throw new Error(`addDrawings: the drawing ${drawing.name} has no relationships (prepareModel was not run)`);
        }
        zip.append(relsXform.toXml(drawing.rels), { name: `xl/drawings/_rels/${drawing.name}.xml.rels` });
    }
}
