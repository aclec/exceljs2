// The pivot tables of the worksheets of a file read, carried opaque to the file written (no pivot model, no editing API):
// read (loadPivotParts), attached to their worksheets (worksheetPivotTables), named and written again (PivotPartWriter).
//
// A pivot table is three parts: the pivot table (`xl/pivotTables/pivotTableN.xml`, a relationship of its worksheet), the
// pivot cache definition it uses (a relationship of the pivot table, and a `<pivotCache cacheId r:id>` of xl/workbook.xml)
// and, optionally, the cache records (a relationship of the cache definition). The bytes of the pivot table reference the
// cache by its `cacheId`, those of the cache reference the records by their relationship id: the cacheIds and the
// relationship ids of the file are kept as they are, the part names are new.
//
// Security (the parts come from an untrusted file):
// - a relationship target is resolved against the part that holds it, to the name of an entry of the zip only
//   (chart-parts.ts resolvePartName: `../` normalized, no climbing above the root, no scheme, backslash, query, fragment,
//   percent escape or empty segment, no relationships part or `[Content_Types].xml`);
// - only the relationships of a pivot table are followed: worksheet → pivot table, pivot table → its cache definition (one),
//   cache definition → its records (at most one); a part reached as two kinds is refused, so the graph has no cycle. The
//   part names and content types are chosen by the writer; relationship ids and external targets are written back through
//   XmlStream (encoded);
// - external relationships (e.g. the source workbook of a cache) are kept, never followed;
// - a pivot table is carried whole or not at all: a missing or refused part, a cacheId that does not match the
//   `<pivotCache>` of its cache in xl/workbook.xml (or is listed twice there), a cache whose source is an external
//   connection (xl/connections.xml is not carried: the cache would reference a missing connection), or the bound on the
//   number of parts exceeded: the pivot table is dropped, as every pivot table was before;
// - a pivot table and a cache definition are parsed with the SAX parser of the library (linear), to check that they are
//   well-formed (malformed XML refuses the part) and to read their root element and the `<cacheSource>` of a cache; the
//   cache records (a copy of the source data, the largest part) are not parsed;
// - maps keyed by part names, relationship ids or cacheIds are Maps.

import parseSax from "../utils/parse-sax";
import type { SaxTag } from "../utils/parse-sax";
import type { ZipWriter } from "../utils/zip-stream";
import RelationshipsXform from "./xform/core/relationships-xform";
import RelType from "./rel-type";
import { relativeTarget, resolvePartName, type ZipEntries } from "./chart-parts";
import { readPartRels, relsPartName } from "./opaque-parts";
import type { PivotSourceLookup } from "./pivot-chart";
import type {
    ContentTypeOverride,
    PivotCachePartModel,
    PivotPartKind,
    PivotPartModel,
    PivotPartRegistry,
    PivotRelModel,
    PivotTablePartModel,
    RelationshipModel,
    WorkbookPivotCacheModel,
} from "../types";

/** The bound on the number of parts of the pivot tables carried from a file (a pivot table has 1 to 3 parts). */
export const MAX_PIVOT_PARTS = 10000;

const CONTENT_TYPES: Readonly<Record<PivotPartKind, string>> = {
    pivotTable: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml",
    pivotCacheDefinition: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml",
    pivotCacheRecords: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml",
};

// a cacheId (xsd:unsignedInt) as written: digits only (anchored, linear)
const CACHE_ID = /^\d{1,10}$/;
const MAX_CACHE_ID = 0xffffffff;

/** A cacheId of the file as a number, or undefined when it is not an xsd:unsignedInt. */
function cacheIdOf(value: string | undefined): number | undefined {
    if (value === undefined || !CACHE_ID.test(value)) {
        return undefined;
    }
    const id = Number(value);
    return id <= MAX_CACHE_ID ? id : undefined;
}

// the local name of a qualified name
function localName(name: string): string {
    return name.slice(name.indexOf(":") + 1);
}

// the bytes of a part, in chunks for the SAX parser
function* chunks(data: Uint8Array): Generator<Uint8Array, void, undefined> {
    const size = 16 * 1024;
    for (let i = 0; i < data.length; i += size) {
        yield data.subarray(i, i + size);
    }
}

/**
 * The first `count` start tags of an XML part (fewer when the document has fewer elements), the whole part checked to be
 * well-formed.
 *
 * @returns the tags, or undefined for malformed XML
 */
async function firstTags(data: Uint8Array, count: number): Promise<SaxTag[] | undefined> {
    const tags: SaxTag[] = [];
    try {
        for await (const events of parseSax(chunks(data))) {
            for (const event of events) {
                if (event.eventType === "opentag" && tags.length < count) {
                    tags.push(event.value);
                }
            }
        }
    } catch {
        return undefined;
    }
    return tags;
}

/**
 * The cacheId and the name of a pivot table part (`<pivotTableDefinition cacheId name>`, both required), or undefined when
 * the part is refused.
 */
async function pivotTableRoot(data: Uint8Array): Promise<{ cacheId: number; name: string } | undefined> {
    const [root] = (await firstTags(data, 1)) ?? [];
    if (!root || localName(root.name) !== "pivotTableDefinition") {
        return undefined;
    }
    const cacheId = cacheIdOf(root.attributes.cacheId);
    const { name } = root.attributes;
    return cacheId === undefined || !name ? undefined : { cacheId, name };
}

/**
 * Whether a cache definition part can be carried: its root is `<pivotCacheDefinition>`, its first child `<cacheSource>`
 * (CT_PivotCacheDefinition), whose source is not an external connection (`type="external"` or a `connectionId`).
 */
async function isCarriedCache(data: Uint8Array): Promise<boolean> {
    const [root, source] = (await firstTags(data, 2)) ?? [];
    return Boolean(
        root &&
        source &&
        localName(root.name) === "pivotCacheDefinition" &&
        localName(source.name) === "cacheSource" &&
        source.attributes.type !== "external" &&
        source.attributes.connectionId === undefined
    );
}

/** The state of a load: the kind of each part name met, the records parts taken by a cache, the number of parts read. */
interface LoadState {
    readonly entries: ZipEntries;
    readonly kinds: Map<string, PivotPartKind>;
    readonly records: Set<string>;
    readonly maxParts: number;
    parts: number;
}

// registers the kind of a part name: false when it was met as another kind
function claimKind(state: LoadState, name: string, kind: PivotPartKind): boolean {
    const known = state.kinds.get(name);
    if (known === undefined) {
        state.kinds.set(name, kind);
        return true;
    }
    return known === kind;
}

// the bytes of a part (inflated once), or undefined when it is missing or beyond the bound on the number of parts
async function readPart(state: LoadState, name: string): Promise<Uint8Array | undefined> {
    const entry = state.entries.get(name);
    if (!entry || ++state.parts > state.maxParts) {
        return undefined;
    }
    return entry.data();
}

/**
 * The relationships of a part as carried: the internal ones of `type` (at most `max`) resolved to the part names of the zip,
 * the external ones kept as they are; undefined when the part is refused (another internal relationship, a target refused
 * or missing, more than `max`).
 */
async function partRels(
    state: LoadState,
    name: string,
    type: string | undefined,
    max: number
): Promise<{ rels: RelationshipModel[]; targets: string[] } | undefined> {
    const rels = await readPartRels(state.entries, name);
    if (!rels) {
        return undefined;
    }
    const targets: string[] = [];
    for (const rel of rels) {
        if (rel.TargetMode === "External") {
            continue;
        }
        const target = rel.Type === type ? resolvePartName(name, rel.Target) : undefined;
        if (target === undefined || !state.entries.has(target) || targets.length >= max) {
            return undefined;
        }
        targets.push(target);
    }
    return { rels, targets };
}

// the relationships of a part as carried: to the part of an internal one, or external
function carriedRels(rels: readonly RelationshipModel[], part: PivotPartModel | undefined): PivotRelModel[] {
    return rels.map((rel) =>
        rel.TargetMode === "External" ? { Id: rel.Id, Type: rel.Type, externalTarget: rel.Target } : { Id: rel.Id, Type: rel.Type, part }
    );
}

/**
 * Reads a pivot cache: its definition and its records.
 *
 * @returns the cache, or null when it is refused
 */
async function loadCache(state: LoadState, name: string, cacheId: number): Promise<PivotCachePartModel | null> {
    const data = await readPart(state, name);
    if (!data || !(await isCarriedCache(data))) {
        return null;
    }
    const read = await partRels(state, name, RelType.PivotCacheRecords, 1);
    if (!read) {
        return null;
    }
    let records: PivotPartModel | undefined;
    const [recordsName] = read.targets;
    if (recordsName !== undefined) {
        // the records of one cache, without relationships of their own
        if (!claimKind(state, recordsName, "pivotCacheRecords") || state.records.has(recordsName)) {
            return null;
        }
        state.records.add(recordsName);
        const recordsRels = await readPartRels(state.entries, recordsName);
        const recordsData = await readPart(state, recordsName);
        if (!recordsData || !recordsRels || recordsRels.length) {
            return null;
        }
        records = { kind: "pivotCacheRecords", data: recordsData, rels: [] };
    }
    return { kind: "pivotCacheDefinition", cacheId, data, rels: carriedRels(read.rels, records) };
}

/**
 * Reads the pivot tables of the worksheets (those that the relationships of the worksheets reference), with their caches,
 * inflating each part once.
 *
 * @param worksheetRels - the relationships of the worksheet parts (by the number of the part, `sheet3.xml` → "3")
 * @param pivotCaches - the `<pivotCache>`s of xl/workbook.xml
 * @param workbookRels - the relationships of xl/workbook.xml
 * @param entries - the entries of the zip
 * @param maxParts - the bound on the number of parts read
 * @returns the pivot tables carried (part name → pivot table), complete: each one, its cache and its records were read and
 * accepted
 * @throws (rejects) when a decompression limit of the zip reader is exceeded
 */
export async function loadPivotParts(
    worksheetRels: Readonly<Record<string, readonly RelationshipModel[]>>,
    pivotCaches: readonly WorkbookPivotCacheModel[] | undefined,
    workbookRels: readonly RelationshipModel[] | undefined,
    entries: ZipEntries,
    maxParts = MAX_PIVOT_PARTS
): Promise<Map<string, PivotTablePartModel>> {
    const state: LoadState = { entries, kinds: new Map(), records: new Set(), maxParts, parts: 0 };

    // the caches of the workbook: part name → cacheId, for the caches listed once, with a cacheId listed once (a duplicate
    // makes the cache of a pivot table ambiguous)
    const bookRels = new Map<string, RelationshipModel>();
    for (const rel of workbookRels ?? []) {
        if (!bookRels.has(rel.Id)) {
            bookRels.set(rel.Id, rel);
        }
    }
    const listed: { readonly cacheId: number; readonly name: string }[] = [];
    const nameCounts = new Map<string, number>();
    const idCounts = new Map<number, number>();
    for (const pivotCache of pivotCaches ?? []) {
        const cacheId = cacheIdOf(pivotCache.cacheId);
        const rel = pivotCache.rId === undefined ? undefined : bookRels.get(pivotCache.rId);
        const name =
            rel && rel.Type === RelType.PivotCacheDefinition && rel.TargetMode !== "External"
                ? resolvePartName("xl/workbook.xml", rel.Target)
                : undefined;
        if (cacheId !== undefined && name !== undefined) {
            listed.push({ cacheId, name });
            nameCounts.set(name, (nameCounts.get(name) ?? 0) + 1);
            idCounts.set(cacheId, (idCounts.get(cacheId) ?? 0) + 1);
            claimKind(state, name, "pivotCacheDefinition");
        }
    }
    const cacheIdByPart = new Map<string, number>();
    for (const { cacheId, name } of listed) {
        if (nameCounts.get(name) === 1 && idCounts.get(cacheId) === 1) {
            cacheIdByPart.set(name, cacheId);
        }
    }

    // the pivot tables referenced by the worksheets
    const tableNames = new Set<string>();
    for (const [sheetNo, rels] of Object.entries(worksheetRels)) {
        for (const rel of rels) {
            if (rel.Type === RelType.PivotTable && rel.TargetMode !== "External") {
                const name = resolvePartName(worksheetPartName(sheetNo), rel.Target);
                if (name !== undefined) {
                    tableNames.add(name);
                }
            }
        }
    }

    const caches = new Map<string, PivotCachePartModel | null>();
    const tables = new Map<string, PivotTablePartModel>();
    for (const name of tableNames) {
        if (state.parts >= maxParts) {
            break;
        }
        if (!claimKind(state, name, "pivotTable")) {
            continue;
        }
        const data = await readPart(state, name);
        const root = data && (await pivotTableRoot(data));
        const cacheId = root?.cacheId;
        const read = cacheId === undefined ? undefined : await partRels(state, name, RelType.PivotCacheDefinition, 1);
        const [cacheName] = read ? read.targets : [];
        // the cache of the relationship is the one of the cacheId in xl/workbook.xml
        if (!data || !root || !read || cacheName === undefined || cacheId === undefined || cacheIdByPart.get(cacheName) !== cacheId) {
            continue;
        }
        let cache = caches.get(cacheName);
        if (cache === undefined) {
            cache = await loadCache(state, cacheName, cacheId);
            caches.set(cacheName, cache);
        }
        if (cache) {
            tables.set(name, { kind: "pivotTable", name: root.name, data, cache, rels: carriedRels(read.rels, cache) });
        }
    }
    return tables;
}

// the part name of a worksheet read (worksheets are keyed by the number of their part, `sheet3.xml` → "3")
function worksheetPartName(sheetNo: string): string {
    return `xl/worksheets/sheet${sheetNo}.xml`;
}

/**
 * The pivot tables of a worksheet read: those of its relationships that were carried and that no other worksheet took.
 *
 * @param sheetNo - the number of the worksheet part
 * @param rels - its relationships
 * @param tables - the pivot tables carried ({@link loadPivotParts})
 * @param claimed - the pivot tables already taken (mutated): a pivot table belongs to one worksheet
 */
export function worksheetPivotTables(
    sheetNo: string,
    rels: readonly RelationshipModel[] | undefined,
    tables: ReadonlyMap<string, PivotTablePartModel>,
    claimed: Set<PivotTablePartModel>
): PivotTablePartModel[] {
    const result: PivotTablePartModel[] = [];
    for (const rel of rels ?? []) {
        if (rel.Type !== RelType.PivotTable || rel.TargetMode === "External") {
            continue;
        }
        const name = resolvePartName(worksheetPartName(sheetNo), rel.Target);
        const table = name === undefined ? undefined : tables.get(name);
        if (table && !claimed.has(table)) {
            claimed.add(table);
            result.push(table);
        }
    }
    return result;
}

/**
 * The sheet name of a reference: as written, or, quoted, without its quotes and with its doubled quotes undone.
 *
 * @param text - the sheet part of the reference (after the file name: `Sheet`, `'Sheet''s'`, or `Sheet''s'` when the quote
 *   opened before the file name)
 * @param quoted - the reference is quoted
 * @returns the sheet name, or undefined for a quoted name with a lone quote
 */
function unquoteSheetName(text: string, quoted: boolean): string | undefined {
    if (!quoted) {
        return text;
    }
    if (!text.endsWith("'")) {
        return undefined;
    }
    const inner = text.slice(text.startsWith("'") ? 1 : 0, -1);
    const pieces = inner.split("''");
    return pieces.some((piece) => piece.includes("'")) ? undefined : pieces.join("'");
}

/** A cache written: its part name and cacheId. */
export interface PivotCacheEntry {
    readonly name: string;
    readonly cacheId: number;
}

/**
 * Names the parts of the pivot tables carried by a workbook being written (`xl/pivotTables/pivotTableN.xml`,
 * `xl/pivotCache/pivotCacheDefinitionN.xml`, `xl/pivotCache/pivotCacheRecordsN.xml`, numbered from `first`: the parts of a
 * pivot table added with `addPivotTable` are numbered 1), then writes them, with their relationships. A cache used by
 * several pivot tables is written once.
 */
export class PivotPartWriter implements PivotPartRegistry, PivotSourceLookup {
    private readonly first: number;
    /** the names of the pivot tables written by worksheet name (lower case: Excel compares sheet names so) */
    private readonly pivotNames = new Map<string, Set<string>>();
    private readonly names = new Map<PivotPartModel, string>();
    private readonly tables: { readonly name: string; readonly table: PivotTablePartModel }[] = [];
    private readonly cacheEntries: PivotCacheEntry[] = [];
    private readonly cacheIds = new Map<number, PivotCachePartModel>();

    /** @param first - the number of the first part of each kind */
    constructor(first = 1) {
        this.first = first;
    }

    /**
     * Adds a pivot table of a worksheet being written: a new part each call; its cache (and records) named on first use.
     *
     * @returns its part name, or undefined when its cache has the cacheId of another cache (not written)
     */
    addTable(table: PivotTablePartModel, sheetName: string): string | undefined {
        const { cache } = table;
        const owner = this.cacheIds.get(cache.cacheId);
        if (owner && owner !== cache) {
            return undefined;
        }
        if (!owner) {
            const number = this.first + this.cacheEntries.length;
            const name = `xl/pivotCache/pivotCacheDefinition${number}.xml`;
            this.cacheIds.set(cache.cacheId, cache);
            this.names.set(cache, name);
            this.cacheEntries.push({ name, cacheId: cache.cacheId });
            for (const rel of cache.rels) {
                if (rel.part) {
                    this.names.set(rel.part, `xl/pivotCache/pivotCacheRecords${number}.xml`);
                }
            }
        }
        const name = `xl/pivotTables/pivotTable${this.first + this.tables.length}.xml`;
        this.tables.push({ name, table });
        const sheetKey = sheetName.toLowerCase();
        const names = this.pivotNames.get(sheetKey);
        if (names) {
            names.add(table.name);
        } else {
            this.pivotNames.set(sheetKey, new Set([table.name]));
        }
        return name;
    }

    /**
     * Whether the pivot table that the `<c:pivotSource><c:name>` of a pivot chart names is written: `[Book1.xlsx]Sheet!Name`,
     * `'[Book1.xlsx]Sheet''s'!Name`, `Sheet!Name` or `'Sheet''s'!Name` (the file name ignored, the sheet name compared
     * without case, the pivot table name exactly; a name with another form is not found: the chart is then static).
     */
    hasPivotTable(source: string): boolean {
        // a sheet name has no "]" (the file name ends at the last one); a pivot table name is taken after the last "!"
        const rest = source.slice(source.lastIndexOf("]") + 1);
        const bang = rest.lastIndexOf("!");
        if (bang === -1) {
            return false;
        }
        const sheetName = unquoteSheetName(rest.slice(0, bang), rest.startsWith("'") || source.startsWith("'"));
        return sheetName !== undefined && (this.pivotNames.get(sheetName.toLowerCase())?.has(rest.slice(bang + 1)) ?? false);
    }

    /** The caches written (in the order of their part names), for xl/workbook.xml and its relationships. */
    get caches(): readonly PivotCacheEntry[] {
        return this.cacheEntries;
    }

    /** @returns `preferred`, or the next cacheId from it that no cache written has (for the pivot table added) */
    freeCacheId(preferred: number): number {
        let id = preferred;
        while (this.cacheIds.has(id)) {
            id++;
        }
        return id;
    }

    /** @returns the content type of each part written */
    contentTypes(): ContentTypeOverride[] {
        const overrides: ContentTypeOverride[] = this.tables.map(({ name }) => ({ name, contentType: CONTENT_TYPES.pivotTable }));
        for (const [part, name] of this.names) {
            overrides.push({ name, contentType: CONTENT_TYPES[part.kind] });
        }
        return overrides;
    }

    /** Adds the parts and their relationships parts to the zip. */
    write(zip: ZipWriter): void {
        const relsXform = new RelationshipsXform();
        const append = (name: string, part: PivotPartModel): void => {
            zip.append(part.data, { name });
            if (part.rels.length) {
                const rels = part.rels.map((rel): RelationshipModel => {
                    if (!rel.part) {
                        return { Id: rel.Id, Type: rel.Type, Target: rel.externalTarget ?? "", TargetMode: "External" };
                    }
                    const target = this.names.get(rel.part);
                    if (target === undefined) {
                        throw new Error(`PivotPartWriter: a part referenced by ${name} was not named`);
                    }
                    return { Id: rel.Id, Type: rel.Type, Target: relativeTarget(name, target) };
                });
                zip.append(relsXform.toXml(rels), { name: relsPartName(name) });
            }
        };
        for (const { name, table } of this.tables) {
            append(name, table);
        }
        for (const [part, name] of this.names) {
            append(name, part);
        }
    }
}
