// Internal models of the pivot tables of a worksheet read from a file, carried opaque to the file written
// (src/xlsx/pivot-parts.ts): the pivot table part, its pivot cache definition and the cache records, as their bytes, with
// their relationships. There is no editing API: a pivot table is read, kept with its worksheet, and written back.

/** The kinds of the parts of a pivot table carried from a file read. */
export type PivotPartKind = "pivotTable" | "pivotCacheDefinition" | "pivotCacheRecords";

/** A part of a pivot table, carried as its bytes (never modified), with its relationships. */
export interface PivotPartModel {
    readonly kind: PivotPartKind;
    readonly data: Uint8Array;
    /** Its relationships, with the ids of the file (the bytes reference them, e.g. `<pivotCacheDefinition r:id>`). */
    readonly rels: readonly PivotRelModel[];
}

/** A relationship of a pivot part: to another part carried, or to an external target (kept as is, never followed). */
export interface PivotRelModel {
    /** The id in the file. */
    readonly Id: string;
    readonly Type: string;
    /** The part targeted (an internal relationship). */
    readonly part?: PivotPartModel;
    /** The target of an external relationship (`TargetMode="External"`), as in the file. */
    readonly externalTarget?: string;
}

/** A pivot cache: its definition part (the cache records, if any, are the part of its records relationship). */
export interface PivotCachePartModel extends PivotPartModel {
    readonly kind: "pivotCacheDefinition";
    /** The `cacheId` of its `<pivotCache>` in `xl/workbook.xml`, which the bytes of its pivot tables reference. */
    readonly cacheId: number;
}

/** A pivot table of a worksheet: its part, and the pivot cache it uses (also the part of one of its relationships). */
export interface PivotTablePartModel extends PivotPartModel {
    readonly kind: "pivotTable";
    /** Its `name` (`<pivotTableDefinition name>`), which the pivot charts built on it reference. */
    readonly name: string;
    readonly cache: PivotCachePartModel;
}

/** Names the parts of the pivot tables carried, while a workbook is written (pivot-parts.ts PivotPartWriter). */
export interface PivotPartRegistry {
    /**
     * Adds a pivot table of a worksheet being written (a new part each call: a pivot table of a worksheet copied to another
     * worksheet is written twice, with the same cache).
     *
     * @param sheetName - the name of the worksheet written (the pivot charts name the pivot table with it)
     * @returns its part name (`xl/pivotTables/pivotTable1.xml`), or undefined when it is not written: its cache has the
     * cacheId of another cache (pivot tables of two files in one workbook)
     */
    addTable(table: PivotTablePartModel, sheetName: string): string | undefined;
}
