// The `options` objects passed to the xforms' `prepare(model, options)` (write) and `reconcile(model, options)` (read).
// They carry the state shared by the xforms of a workbook (style and shared string managers, media...) and of a worksheet
// (merges, hyperlinks, shared formulae...). The worksheet xform adds the per-sheet fields to the workbook-level object
// (mutation), and ListXform sets `index` on it before each item. An xform narrows its `options` parameter to what it reads,
// e.g. `prepare(model: RowModel, options: CellPrepareOptions)`.

import type { FormulaCellModel, MergeCellModel } from "./cell";
import type { CommentRefsEntry, MediaModel, RelationshipModel } from "./book";
import type { DrawingModel, NoteModel, SheetHyperlinkModel, TableModel } from "./sheet";
import type { StyleLookup, StyleRegistry } from "./style";
import type { SharedStringLookup, SharedStringRegistry } from "./strings";

/**
 * Workbook-level prepare options (xlsx.ts prepareModel, passed to every worksheet xform).
 * The stream writer builds its own subset per row ({@link CellPrepareOptions}).
 */
export interface WorkbookPrepareOptions {
    styles: StyleRegistry;
    sharedStrings?: SharedStringRegistry;
    /** the workbook uses the 1904 date system */
    date1904?: boolean;
    /** the number of drawings so far (incremented by the worksheet xform) */
    drawingsCount: number;
    /** the workbook media, indexed by `imageId` */
    media: MediaModel[];
    /** collects the drawing of each worksheet (the same array as the write model's `drawings`) */
    drawings: DrawingModel[];
    /** collects the comments parts (the same array as the write model's `commentRefs`) */
    commentRefs: CommentRefsEntry[];
}

/**
 * The prepare options read by the cell / row / column xforms. The worksheet xform sets the per-sheet fields; the stream
 * writer builds the object for each row (`hyperlinks` is then a proxy with `push`, `merges` an array with a no-op `add`,
 * and `date1904` is not set).
 */
export interface CellPrepareOptions {
    styles: StyleRegistry;
    sharedStrings?: SharedStringRegistry;
    date1904?: boolean;
    /** collects the hyperlinks of the cells */
    hyperlinks: { push(hyperlink: SheetHyperlinkModel): unknown };
    /** collects the merged cells (a Merges instance) */
    merges: { add(cell: MergeCellModel): unknown };
    /** collects the notes of the cells, with their `ref` */
    comments: (NoteModel & { ref: string })[];
    /** shared formula masters by cell address (prepare) */
    formulae: Record<string, FormulaCellModel>;
    /** the next shared formula index (incremented) */
    siFormulae: number;
}

/** The prepare options inside a worksheet: the workbook-level ones plus the per-sheet ones. */
export type WorksheetPrepareOptions = WorkbookPrepareOptions & CellPrepareOptions;

/** Workbook-level reconcile options (xlsx.ts, passed to every worksheet xform). PARTIAL: the part hashes are typed loosely. */
export interface WorkbookReconcileOptions {
    styles?: StyleLookup;
    sharedStrings?: SharedStringLookup;
    media: MediaModel[];
    mediaIndex: Record<string, number>;
    date1904?: boolean;
    drawings: Record<string, DrawingModel>;
    comments: Record<string, { comments: NoteModel[] }>;
    tables: Record<string, TableModel>;
    vmlDrawings: Record<string, { comments: unknown[] }>;
    /** the parts already claimed by a worksheet (a part belongs to one worksheet) */
    claimedParts: Set<object>;
}

/**
 * The reconcile options read by the cell / row / column xforms. The maps are keyed by data of the file: they are created
 * with `Object.create(null)`.
 */
export interface CellReconcileOptions {
    styles?: StyleLookup;
    sharedStrings?: SharedStringLookup;
    date1904?: boolean;
    /** cell address → its hyperlink */
    hyperlinkMap: Record<string, { target: string; tooltip?: string }>;
    /** cell address → its note */
    commentsMap?: Record<string, NoteModel>;
    /** shared formula index (`si`, as written) → master cell address. NB: a different type than in prepare. */
    formulae: Record<string, string>;
}

/** The reconcile options inside a worksheet: the workbook-level ones plus the per-sheet ones. */
export type WorksheetReconcileOptions = WorkbookReconcileOptions & CellReconcileOptions;

/** The reconcile options of a drawing (xlsx.ts). */
export interface DrawingReconcileOptions {
    media: MediaModel[];
    mediaIndex: Record<string, number>;
    /** relationship id → relationship, of the drawing part */
    rels?: Record<string, RelationshipModel>;
}
