// Internal models of a worksheet and its parts: the JSON exchanged between src/doc/worksheet.ts (and the classes it
// holds) and the sheet xforms (src/xlsx/xform/sheet, drawing, comment, table), xlsx.ts and the stream writer / reader.
//
// PARTIAL: the models of conditional formatting, tables, drawings and VML are typed at the level the worksheet xform and
// xlsx.ts exchange them; their inner details are refined by the agents typing those xforms (see the notes on each type).

import type {
    AutoFilter,
    Comment,
    ConditionalFormattingRule,
    DataValidation,
    HeaderFooter,
    ImageEditAs,
    Margins,
    PageSetup,
    TableColumnProperties,
    TableStyleProperties,
    WorksheetProperties,
    WorksheetProtection,
    WorksheetState,
    WorksheetView,
} from "../../index";
import type { CellModel, ColModel, FormulaResult, ParsedCellModel, RowModel } from "./cell";
import type { PivotTable } from "../doc/pivot-table";
import type { DxfModel } from "./style";
import type { MediaModel, RelationshipModel } from "./book";

// ---------------------------------------------------------------------------------------------------------------------
// Notes (cell comments)

/**
 * The model of a cell note. Producer: `Note.model` (src/doc/note.ts: deep-merged with the defaults margins / protection /
 * editAs), comment-xform parse (`{ type: "note", note: { texts }, ...the <comment> attributes }`, e.g. `ref`, `authorId`
 * as strings) merged with the VML shape. Consumer: `Note.fromModel`, comments-xform / vml-notes-xform render.
 */
export interface NoteModel {
    type: "note";
    /** index.d.ts `Comment`; a parsed note may also carry the VML `anchor` (a string). */
    note: Comment & { anchor?: string };
    /** prepare (cell-xform) / parse: the address of the cell. */
    ref?: string;
    /**
     * Parse: an attribute of `<comment>`. PARTIAL: comment-xform spreads every attribute of the element into the model
     * (strings; even over `type` / `note`), only the known ones are typed.
     */
    authorId?: string;
}

/** A note as collected for the comments part by worksheet-xform prepare: `refAddress` is the decoded `ref`. */
export interface CommentRefModel extends NoteModel {
    ref: string;
    refAddress: { row: number; col: number; address: string; $col$row?: string };
}

// ---------------------------------------------------------------------------------------------------------------------
// Hyperlinks, merges, data validations

/**
 * A `<hyperlink>` of a worksheet (hyperlink-xform). Write: `{ address, target, tooltip }` collected by cell-xform prepare,
 * `rId` added by worksheet-xform prepare for an external target. Parse: `{ address, rId?, tooltip?, target? }` (`target`
 * for an internal `location`). worksheet-xform reconcile turns them into the `hyperlink` of the cells (hyperlinkMap).
 */
export interface SheetHyperlinkModel {
    address: string;
    target?: string;
    tooltip?: string;
    rId?: string;
}

/**
 * The data validations of a worksheet (`DataValidations.model`), keyed by cell address (`A1`) or by range (`A1:B10`,
 * whole rows / columns normalised to rectangles); a removed entry is kept as `undefined`. Parsing shares one object
 * between all the cells of a `<dataValidation>`, and keeps a range key beyond a budget of expanded cells.
 * Object created with `Object.create(null)` when parsed (keys come from the file).
 */
export type DataValidationsModel = Record<string, DataValidationModel | undefined>;

/**
 * One data validation (data-validations-xform). Differs from index.d.ts `DataValidation`: `type` may be `"any"` (the
 * default when parsing); `formulae` and `operator` are deleted when there is no formula.
 */
export type DataValidationModel = Omit<DataValidation, "type" | "formulae"> & {
    type: DataValidation["type"] | "any";
    formulae?: DataValidation["formulae"];
};

// ---------------------------------------------------------------------------------------------------------------------
// Conditional formatting (PARTIAL)

/**
 * A conditional formatting rule as the xforms see it: index.d.ts `ConditionalFormattingRule` plus the internal fields.
 * PARTIAL: the inner models of dataBar / iconSet / colorScale (cf/*, cf-ext/*) are refined by their xforms.
 */
export type ConditionalFormattingRuleModel = ConditionalFormattingRule & {
    /** prepare: the dxf index of `style` (StyleRegistry.addDxfStyle); reconcile turns it back into `style` and deletes it. */
    dxfId?: number;
    /** x14 rules: the id linking a main rule to its extension (prepare adds it; kept on loaded rules). */
    x14Id?: string;
    /** render: the range of the formatting, written onto each rule (mutates the caller's rule). */
    ref?: string;
    /** x14 icon sets: custom icons. */
    icons?: { iconSet: string; iconId: number }[];
    /** Parsed formula rules: the `<formula>` texts. */
    formulae?: string[];
};

/** The conditional formattings of one range (conditional-formatting-xform, x14 version merged in by worksheet-xform). */
export interface ConditionalFormattingModel {
    ref: string;
    rules: ConditionalFormattingRuleModel[];
}

// ---------------------------------------------------------------------------------------------------------------------
// Views, page setup, properties, protection

/**
 * A manual page break after a row (`row.addPageBreak`, rendered as a `<brk>` of `<rowBreaks>`): `id` is the row number,
 * `min` / `max` the 0-based first and last columns (`min` absent for the first column), `man` 1 (manual).
 */
export interface RowBreakModel {
    id: number;
    max: number;
    man: number;
    min?: number;
}

/**
 * A sheet view (sheet-view-xform): index.d.ts `WorksheetView` plus `workbookViewId` (parsed) and `tabSelected` (rendered;
 * parsed but not kept). prepare forces an unknown `state` to `"normal"`.
 */
export type SheetViewModel = Partial<WorksheetView> & {
    workbookViewId?: number;
    tabSelected?: boolean;
};

/**
 * The page setup of a worksheet: index.d.ts `PageSetup` (with `margins`, the print options, the print area and titles).
 * Parse: `{ fitToPage, margins }` merged with the page-setup and print-options models; also `useFirstPageNumber`,
 * `usePrinterDefaults`, `copies`. The document defaults have `rowBreaks` / `colBreaks` set to `null`.
 */
export type PageSetupModel = Partial<PageSetup> & {
    margins?: Margins;
    /** The document defaults only: always null (the page breaks are `rowBreaks` of the worksheet). */
    rowBreaks?: null;
    /** The document defaults only: always null. */
    colBreaks?: null;
    useFirstPageNumber?: boolean;
    usePrinterDefaults?: boolean;
    copies?: number;
};

/** Header and footer (header-footer-xform): index.d.ts `HeaderFooter` (the document defaults are `null` strings). */
export type HeaderFooterModel = { [K in keyof HeaderFooter]?: HeaderFooter[K] | (HeaderFooter[K] extends string ? null : never) };

/**
 * Sheet properties: `sheetFormatPr` (defaultRowHeight, dyDescent, outlineLevelRow, outlineLevelCol, defaultColWidth?)
 * merged with `tabColor` / `outlineProperties` of `sheetPr` (worksheet-xform). index.d.ts `WorksheetProperties`; its
 * `showGridLines` is not read or written by any xform.
 */
export type WorksheetPropertiesModel = Partial<WorksheetProperties>;

/**
 * Sheet protection (sheet-protection-xform): index.d.ts `WorksheetProtection` plus `sheet` and the password hash
 * (`algorithmName`, `hashValue`, `saltValue`, set by `worksheet.protect`).
 */
export type SheetProtectionModel = Partial<WorksheetProtection> & {
    sheet?: boolean;
    algorithmName?: string;
    hashValue?: string;
    saltValue?: string;
};

/** The auto filter of a worksheet: index.d.ts `AutoFilter` (parsed as the `ref` string). */
export type AutoFilterModel = AutoFilter;

// ---------------------------------------------------------------------------------------------------------------------
// Images, drawings

/** A cell anchor of an image (src/doc/anchor.ts `Anchor.model`): 0-based column / row and offsets in EMU. */
export interface AnchorModel {
    nativeCol: number;
    nativeColOff: number;
    nativeRow: number;
    nativeRowOff: number;
}

/** The position of an image: two-cell (`tl` + `br`) or one-cell (`tl` + `ext`, in pixels). */
export interface ImageRangeModel {
    tl: AnchorModel;
    br?: AnchorModel;
    ext?: { width: number; height: number };
    editAs?: ImageEditAs;
}

/** The hyperlink of an image. */
export interface ImageHyperlinkModel {
    hyperlink?: string;
    tooltip?: string;
    /** drawing parse: the relationship id, replaced by `hyperlink` in xlsx.ts. */
    rId?: string;
}

/**
 * A medium of a worksheet (src/doc/image.ts `Image.model`): a background image, or an image in the sheet. `imageId` is
 * the index of the image in the workbook media.
 */
export type WorksheetMediumModel =
    | { type: "background"; imageId: number | string }
    | { type: "image"; imageId: number | string; range: ImageRangeModel; hyperlinks?: ImageHyperlinkModel };

/**
 * The drawing part of a worksheet. Write (worksheet-xform prepare): `{ rId, name: "drawingN", anchors, rels }`; parse:
 * `{ anchors }`. PARTIAL: the anchor model is refined by the drawing xforms (anchorType, picture.index / creationId on
 * write; picture.rId / medium on read).
 */
export interface DrawingModel {
    rId?: string;
    name?: string;
    anchors: DrawingAnchorModel[];
    rels?: RelationshipModel[];
}

/** One anchor of a drawing (PARTIAL, see {@link DrawingModel}). */
export interface DrawingAnchorModel {
    range: ImageRangeModel;
    picture?: { rId: string; hyperlinks?: ImageHyperlinkModel; index?: number; creationId?: number | string };
    anchorType?: "xdr:twoCellAnchor" | "xdr:oneCellAnchor";
    /** reconcile: the workbook medium of the picture. */
    medium?: MediaModel;
}

// ---------------------------------------------------------------------------------------------------------------------
// Tables (PARTIAL)

/**
 * A table column as the xforms see it: index.d.ts `TableColumnProperties` plus the write-side ids (`id`, `colId`, `dxfId`)
 * and the parse-side `dxfId` (a string). PARTIAL: refined by the table xforms.
 */
export type TableColumnModel = TableColumnProperties & {
    /** Document side: the cached result of the totals row formula (written in the totals row cell). */
    totalsRowResult?: FormulaResult;
    id?: number;
    colId?: string;
    dxfId?: number | string;
    style?: DxfModel;
};

/**
 * A table (src/doc/table.ts `Table.model`, table-xform). Write: the validated document table plus `tableRef`,
 * `autoFilterRef`, `tl`, then `target` / `id` (xlsx.ts) and `rId` (worksheet-xform). Parse: `{ name, displayName,
 * tableRef, totalsRow, headerRow, columns, autoFilterRef?, style }`, without `ref` / `rows` / `tl`.
 * PARTIAL: refined by the table xforms.
 */
export interface TableModel {
    name: string;
    displayName?: string;
    ref?: string;
    tableRef?: string;
    autoFilterRef?: string;
    headerRow?: boolean;
    totalsRow?: boolean;
    style?: TableStyleProperties;
    columns: TableColumnModel[];
    rows?: unknown[][];
    tl?: { row: number; col: number; address?: string };
    /** write: `tableN.xml` */
    target?: string;
    /** write: the table id (1-based, workbook-wide) */
    id?: number;
    /** write: the relationship id from the worksheet */
    rId?: string;
}

// ---------------------------------------------------------------------------------------------------------------------
// Worksheet

/**
 * A worksheet as the document produces it (`Worksheet.model`, src/doc/worksheet.ts) and consumes it (`Worksheet.model =`),
 * and as worksheet-xform prepares and renders it.
 *
 * Doc getter only: `id`, `state`, `rowBreaks`, `dimensions`, `merges`, `pivotTables`. Doc setter only: `mergeCells` (not
 * `merges`). Write-side additions (worksheet-xform prepare): `mergeCells`, `hyperlinks`, `comments`, `rels`, `drawing`,
 * `background`, `image`; xlsx.ts adds `rId`, `shapeBlock`.
 *
 * Index.d.ts `WorksheetModel` is a subset of this (it lacks dataValidations, state, sheetProtection, tables, ...).
 */
export interface WorksheetModel {
    id: number;
    name: string;
    state?: WorksheetState;
    dataValidations: DataValidationsModel;
    properties: WorksheetPropertiesModel;
    pageSetup: PageSetupModel;
    headerFooter: HeaderFooterModel;
    rowBreaks?: RowBreakModel[];
    views: SheetViewModel[];
    autoFilter?: AutoFilterModel | null;
    media: WorksheetMediumModel[];
    sheetProtection?: SheetProtectionModel | null;
    tables: TableModel[];
    /** The document's pivot tables (src/doc/pivot-table.ts); absent after a load. */
    pivotTables?: PivotTable[];
    conditionalFormattings: ConditionalFormattingModel[];
    /** `undefined` when there is no column definition. */
    cols?: ColModel[];
    rows: RowModel<CellModel>[];
    /** Doc getter: the used range (a Range, rendered as its `range`). Parse: the `ref` string. */
    dimensions?: { readonly range: string } | string;
    /** Doc getter: the merged ranges (`A1:B2`). */
    merges?: string[];
    /** Write (prepare) and parse: the merged ranges; read by the doc setter. */
    mergeCells?: string[];

    // write side (worksheet-xform prepare, xlsx.ts)
    hyperlinks?: SheetHyperlinkModel[];
    comments?: CommentRefModel[];
    rels?: RelationshipModel[];
    drawing?: DrawingModel;
    background?: { rId: string };
    /** The workbook medium of the background image. */
    image?: MediaModel;
    /** xlsx.ts: the relationship id from the workbook. */
    rId?: string;
    /** xlsx.ts: the VML shape block of the comments. */
    shapeBlock?: number;
}

/**
 * A worksheet as worksheet-xform parses it (before reconcile), plus what xlsx.ts and WorkbookXform.reconcile add
 * (`sheetNo`, `relationships`, `name`, `id`, `state`, print areas in `pageSetup`). reconcile turns it into a
 * {@link WorksheetModel}: rows with resolved cells, `media`, `tables` as models; it deletes `relationships`,
 * `legacyDrawing`, `hyperlinks`, `comments`. PARTIAL on the inner models (see above).
 */
export interface ParsedWorksheetModel {
    id?: number;
    name?: string;
    state?: WorksheetState;
    /** xlsx.ts: the number of the part (`sheet3.xml` → "3"). */
    sheetNo?: string;
    relationships?: RelationshipModel[];
    dimensions?: string;
    cols?: ColModel[];
    rows: RowModel<ParsedCellModel>[];
    mergeCells?: string[];
    hyperlinks?: SheetHyperlinkModel[];
    dataValidations?: DataValidationsModel;
    properties: WorksheetPropertiesModel;
    views?: SheetViewModel[];
    pageSetup: PageSetupModel;
    headerFooter?: HeaderFooterModel;
    background?: { rId: string };
    drawing?: { rId: string };
    tables?: { rId: string }[];
    conditionalFormattings?: ConditionalFormattingModel[];
    autoFilter?: AutoFilterModel;
    sheetProtection?: SheetProtectionModel;
    legacyDrawing?: { rId: string };
    /** reconcile: the images, from the drawing and background. */
    media?: WorksheetMediumModel[];
    comments?: NoteModel[];
}
