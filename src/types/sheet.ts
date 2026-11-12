// Internal models of a worksheet and its parts: the JSON exchanged between src/doc/worksheet.ts (and the classes it
// holds) and the sheet xforms (src/xlsx/xform/sheet, drawing, comment, table), xlsx.ts and the stream writer / reader.
//
// PARTIAL: the models of conditional formatting, tables, drawings and VML are typed at the level the worksheet xform and
// xlsx.ts exchange them; their inner details are refined by the agents typing those xforms (see the notes on each type).

import type {
    AutoFilter,
    CellIsOperators,
    CfvoTypes,
    Comment,
    ContainsTextOperators,
    DataBarRuleType,
    DataValidation,
    HeaderFooter,
    IconSetTypes,
    ImageEditAs,
    Margins,
    PageSetup,
    TableColumnProperties,
    TableStyleProperties,
    TimePeriodTypes,
    WorksheetProperties,
    WorksheetProtection,
    WorksheetState,
    WorksheetView,
} from "../../index";
import type { CellModel, ColModel, FormulaResult, ParsedCellModel, RowModel } from "./cell";
import type { ColorModel, DxfModel } from "./style";
import type { PivotTable } from "../doc/pivot-table";
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
     * Parse: an attribute of `<comment>`. comment-xform copies every attribute of the element into the model (strings;
     * never over `type` / `note`): only the known ones are typed.
     */
    authorId?: string;
}

/** A note as collected for the comments part by worksheet-xform prepare: `refAddress` is the decoded `ref`. */
export interface CommentRefModel extends NoteModel {
    ref: string;
    refAddress: { row: number; col: number; address: string; $col$row?: string };
}

/**
 * The VML shape of a note (comment/vml-shape-xform parse): its box settings, as written in the file (unvalidated
 * strings). worksheet-xform reconcile merges it into the `note` of the {@link NoteModel} at the same index.
 */
export interface VmlNoteShapeModel {
    margins: {
        /** `o:insetmode` of `<v:shape>` (`auto` / `custom`). */
        insetmode: string | undefined;
        /** `inset` of `<v:textbox>`, in cm (left, top, right, bottom); invalid: `undefined`; no `<v:textbox>`: `null`. */
        inset?: number[] | null;
    };
    /** The text of `<x:Anchor>` ("col, dx, row, dy, col, dx, row, dy"); `null` without `<x:ClientData>`. */
    anchor: string | null | undefined;
    /** From `<x:MoveWithCells>` / `<x:SizeWithCells>`: `twoCells`, `oneCells` or `absolute`; `null` without `<x:ClientData>`. */
    editAs: string | null | undefined;
    /** The texts of `<x:Locked>` / `<x:LockText>` (`True` / `False`); `null` without `<x:ClientData>`. */
    protection: { locked?: string; lockText?: string } | null | undefined;
}

/** A VML drawing of notes, `xl/drawings/vmlDrawingN.vml` (comment/vml-notes-xform parse): the shapes, in order. */
export interface VmlNotesModel {
    comments: VmlNoteShapeModel[];
}

// ---------------------------------------------------------------------------------------------------------------------
// Hyperlinks, merges, data validations

/**
 * A `<hyperlink>` of a worksheet (hyperlink-xform). Write: `{ address, target, tooltip }` collected by cell-xform prepare,
 * `rId` added by worksheet-xform prepare for an external target. Parse: `{ address, rId?, tooltip?, target? }` (`target`
 * for an internal `location`). worksheet-xform reconcile turns them into the `hyperlink` of the cells (hyperlinkMap).
 */
export interface SheetHyperlinkModel {
    /** The cell; `undefined` in a parsed `<hyperlink>` without `ref` (ignored by reconcile). */
    address: string | undefined;
    target?: string;
    tooltip?: string;
    rId?: string;
}

/**
 * The data validations of a worksheet (`DataValidations.model`), keyed by cell address (`A1`) or by range (`A1:B10`,
 * whole rows / columns normalised to rectangles); a removed entry is kept as `undefined`. Parsing shares one object
 * between all the cells of a `<dataValidation>`, and keeps a range key beyond a budget of expanded cells. A parsed model is
 * a plain object (public API) whose keys are validated cell / range references.
 */
export type DataValidationsModel = Record<string, DataValidationModel | undefined>;

/**
 * A formula of a data validation: a number for the numeric types (`whole`, `textLength`, `decimal`), a Date for `date`
 * (written as a serial date), else the text of the formula (e.g. `"a,b,c"` with its quotes for a list, or `$A$1:$A$9`).
 */
export type DataValidationFormula = string | number | Date;

/**
 * One data validation (data-validations-xform): index.d.ts `DataValidation`. `formulae` and `operator` are deleted when
 * there is no formula. Parsed literal unions hold the attribute strings of the file, unvalidated.
 */
export type DataValidationModel = DataValidation;

/**
 * The budgets of one workbook load, shared by all its worksheets (WorksheetXform.createLoadBudget): the cells of the
 * merges materialised so far against their cap, and the cells that the data validation ranges may still be expanded to.
 */
export interface WorksheetLoadBudget {
    mergedCells: number;
    maxMergedCells: number;
    dataValidationCells: number;
}

// ---------------------------------------------------------------------------------------------------------------------
// Conditional formatting
//
// The rules are written in `<conditionalFormatting>` (cf/*: Excel 2007) and / or in the x14 extension of the worksheet
// (cf-ext/*: data bars with x14 settings, custom or x14 icon sets, formula rules referencing other sheets); worksheet-xform
// merges the two parsed versions of a rule by `x14Id`. A rule model is index.d.ts `ConditionalFormattingRule` with the
// internal fields; index.d.ts rules are assignable to it.
//
// Parsed rules are unvalidated: their literal unions (`type`, `operator`, `iconSet`, cfvo `type`...) hold the attribute
// strings of the file, and the main `<cfRule>` parse sets the keys of every rule type (`undefined` when absent).

/** A formula of a conditional formatting rule: the text of `<formula>` / `<xm:f>` (the application may give numbers). */
export type CfFormula = string | number | boolean;

/**
 * A threshold of a color scale, data bar or icon set (`<cfvo>`, `<x14:cfvo>`): index.d.ts `Cvfo`. `value` is a number
 * when the text is a number literal, else the text (a formula or a cell reference such as `$A$1`).
 */
export interface CfvoModel {
    type: CfvoTypes;
    value?: number | string;
}

/** A custom icon of an x14 icon set (`<x14:cfIcon>`). `iconId` is set by render (the index of the icon). */
export interface CfIconModel {
    iconSet: IconSetTypes;
    iconId?: number;
}

/** Fields of every rule model. */
interface CfRuleModelBase {
    /** Rules without a priority get one (after the highest one) in prepare. */
    priority?: number;
    /** The differential format; reconcile sets it from `dxfId`. */
    style?: DxfModel;
    stopIfTrue?: boolean;
    /** prepare: the dxf index of `style` (StyleRegistry.addDxfStyle); parse: the attribute; reconcile turns it into `style` and deletes it. */
    dxfId?: number;
    /** The id linking a main rule to its x14 version: set by prepare (cf-ext) for an x14 rule, kept on loaded rules. */
    x14Id?: string;
    /** render: the range of the formatting, written onto each rule (mutates the caller's rule). */
    ref?: string;
    /** The formulae: the parse collects the `<formula>` / `<xm:f>` texts of any rule; render reads them for formula rules. */
    formulae?: CfFormula[];
}

/** An `expression` rule (index.d.ts `ExpressionRuleType`). */
export interface ExpressionRuleModel extends CfRuleModelBase {
    type: "expression";
    /** Written in the x14 extension (set on the rules read from the extension only). */
    x14?: boolean;
}

/** A `cellIs` rule (index.d.ts `CellIsRuleType`). */
export interface CellIsRuleModel extends CfRuleModelBase {
    type: "cellIs";
    operator?: CellIsOperators;
    x14?: boolean;
}

/** A `top10` rule (index.d.ts `Top10RuleType`; the parsed attributes may be absent). */
export interface Top10RuleModel extends CfRuleModelBase {
    type: "top10";
    /** Default 10. */
    rank?: number;
    percent?: boolean;
    bottom?: boolean;
}

/** An `aboveAverage` rule (index.d.ts `AboveAverageRuleType`). */
export interface AboveAverageRuleModel extends CfRuleModelBase {
    type: "aboveAverage";
    /** Default true. */
    aboveAverage?: boolean;
}

/** A `colorScale` rule (index.d.ts `ColorScaleRuleType`): rendering requires `cfvo` and `color`. */
export interface ColorScaleRuleModel extends CfRuleModelBase {
    type: "colorScale";
    cfvo?: CfvoModel[];
    color?: ColorModel[];
}

/**
 * An `iconSet` rule (index.d.ts `IconSetRuleType`): rendering requires `cfvo`. A custom icon set (`custom`, or `icons`
 * when read) and the x14 icon sets (3Triangles, 3Stars, 5Boxes) are written in the extension only.
 */
export interface IconSetRuleModel extends CfRuleModelBase {
    type: "iconSet";
    /** Default 3TrafficLights. */
    iconSet?: IconSetTypes;
    reverse?: boolean;
    showValue?: boolean;
    custom?: boolean;
    cfvo?: CfvoModel[];
    /** x14 parse: the custom icons (absent from index.d.ts: rendered as `custom="1"` with `<x14:cfIcon>` elements). */
    icons?: CfIconModel[];
}

/**
 * A text rule (index.d.ts `ContainsTextRuleType`): `operator` (default containsText) is the `type` of its `<cfRule>`; a
 * rule without formulae gets one derived from the operator, `text` and the top left cell of `ref`.
 */
export interface ContainsTextRuleModel extends CfRuleModelBase {
    type: "containsText";
    operator?: ContainsTextOperators;
    text?: string;
    x14?: boolean;
}

/** A `timePeriod` rule (index.d.ts `TimePeriodRuleType`): a rule without formulae gets one derived from `timePeriod`. */
export interface TimePeriodRuleModel extends CfRuleModelBase {
    type: "timePeriod";
    timePeriod?: TimePeriodTypes;
    x14?: boolean;
}

/**
 * A `dataBar` rule (index.d.ts `DataBarRuleType`). The main `<dataBar>` holds `showValue`, `cfvo` and `color`; the x14
 * version (written unless it is a plain gradient bar, see DatabarExtXform.isExt) holds the other settings, and its
 * thresholds replace the main ones when both are read.
 */
export interface DataBarRuleModel extends CfRuleModelBase {
    type: "dataBar";
    gradient?: boolean;
    minLength?: number;
    maxLength?: number;
    showValue?: boolean;
    border?: boolean;
    negativeBarColorSameAsPositive?: boolean;
    negativeBarBorderColorSameAsPositive?: boolean;
    axisPosition?: DataBarRuleType["axisPosition"];
    direction?: DataBarRuleType["direction"];
    cfvo?: CfvoModel[];
    /** The color of the bars. */
    color?: ColorModel;
    borderColor?: ColorModel;
    negativeFillColor?: ColorModel;
    negativeBorderColor?: ColorModel;
    axisColor?: ColorModel;
}

/** The rules whose formulae and format may be written inline in the x14 extension (CfRuleXform.isFormulaRule). */
export type FormulaRuleModel = ExpressionRuleModel | CellIsRuleModel | ContainsTextRuleModel | TimePeriodRuleModel;

/**
 * A conditional formatting rule as the xforms see it, discriminated by `type`. Producer: the application (index.d.ts
 * `ConditionalFormattingRule`, assignable to this), cf-rule-xform / cf-rule-ext-xform parse (merged by worksheet-xform).
 * Consumer: the cf xforms (prepare mutates `priority`, `dxfId`, `x14Id`; render sets `ref`).
 */
export type ConditionalFormattingRuleModel =
    FormulaRuleModel | Top10RuleModel | AboveAverageRuleModel | ColorScaleRuleModel | IconSetRuleModel | DataBarRuleModel;

/** The conditional formattings of one range (conditional-formatting-xform, x14 version merged in by worksheet-xform). */
export interface ConditionalFormattingModel {
    /** The range (sqref), e.g. `A1:B10`. */
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
 * A sheet view (sheet-view-xform): index.d.ts `WorksheetView` flattened (the fields of its normal / frozen / split
 * variants, all optional) plus `workbookViewId` (parsed) and `tabSelected` (rendered; parsed but not kept). prepare
 * forces an unknown `state` to `"normal"`. Parsed literal unions (`style`, `activePane`) hold the attribute strings of the
 * file, unvalidated.
 */
export interface SheetViewModel {
    workbookViewId?: number;
    tabSelected?: boolean;
    rightToLeft?: boolean;
    activeCell?: string;
    showRuler?: boolean;
    showRowColHeaders?: boolean;
    showGridLines?: boolean;
    zoomScale?: number;
    zoomScaleNormal?: number;
    state?: WorksheetView["state"];
    style?: WorksheetView["style"];
    /** frozen: the number of frozen columns; split: the position of the splitter, in points. */
    xSplit?: number;
    /** frozen: the number of frozen rows; split: the position of the splitter, in points. */
    ySplit?: number;
    topLeftCell?: string;
    /** split only (the active pane of a frozen view is derived from the splits). */
    activePane?: "topLeft" | "topRight" | "bottomLeft" | "bottomRight";
}

/**
 * The page setup of a worksheet: index.d.ts `PageSetup` (with `margins`, the print options, the print area and titles).
 * Parse: `{ fitToPage, margins }` merged with the page-setup and print-options models; also `useFirstPageNumber`,
 * `usePrinterDefaults`, `copies`. The document defaults have `rowBreaks` / `colBreaks` set to `null`.
 */
export type PageSetupModel = Partial<Omit<PageSetup, "margins">> & {
    /** `null` when a parsed worksheet has no `<pageMargins>`. */
    margins?: Margins | null;
    /** The document defaults only: always null (the page breaks are `rowBreaks` of the worksheet). */
    rowBreaks?: null;
    /** The document defaults only: always null. */
    colBreaks?: null;
    useFirstPageNumber?: boolean;
    usePrinterDefaults?: boolean;
    copies?: number;
};

/** Header and footer (header-footer-xform): index.d.ts `HeaderFooter`; the defaults of the document and stream worksheets are `null`. */
export type HeaderFooterModel = { [K in keyof HeaderFooter]?: HeaderFooter[K] | null };

/**
 * `<outlinePr>` (outline-properties-xform): index.d.ts `WorksheetProperties.outlineProperties`; a parsed attribute that is
 * absent is an `undefined` key.
 */
export type OutlinePropertiesModel = Partial<WorksheetProperties["outlineProperties"]>;

/**
 * Sheet properties: `sheetFormatPr` (defaultRowHeight, dyDescent, outlineLevelRow, outlineLevelCol, defaultColWidth?)
 * merged with `tabColor` / `outlineProperties` of `sheetPr` (worksheet-xform). index.d.ts `WorksheetProperties`; its
 * `showGridLines` is not read or written by any xform. `tabColor` may have `tint` / `indexed` (a parsed color).
 */
export type WorksheetPropertiesModel = Partial<Omit<WorksheetProperties, "outlineProperties" | "tabColor">> & {
    outlineProperties?: OutlinePropertiesModel;
    tabColor?: ColorModel;
};

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
 * The drawing part of a worksheet (drawing/drawing-xform). Write (worksheet-xform prepare): `{ rId, name: "drawingN",
 * anchors, rels }`; parse: `{ anchors }`.
 */
export interface DrawingModel {
    rId?: string;
    name?: string;
    anchors: DrawingAnchorModel[];
    rels?: RelationshipModel[];
}

/**
 * The picture of a drawing anchor (drawing/pic-xform). Write: `{ rId, hyperlinks? }` (worksheet-xform prepare), then
 * `index` / `creationId` (pic-xform prepare). Parse: `{ hyperlinks?, rId }` (`<a:hlinkClick>`, `<a:blip r:embed>`).
 */
export interface DrawingPictureModel {
    /** The relationship id of the image (from the drawing part); a parsed `<a:blip>` may lack it. */
    rId?: string;
    hyperlinks?: ImageHyperlinkModel;
    /** prepare: the 1-based index of the picture in its drawing (`id` and name of `<xdr:cNvPr>`). */
    index?: number;
    /** prepare: the GUID-shaped `a16:creationId`, unique in the workbook. */
    creationId?: string;
}

/**
 * One anchor of a drawing: a picture over a range of cells (drawing/two-cell-anchor-xform, one-cell-anchor-xform).
 * The values parsed from a malformed file are not validated: an element absent from the anchor leaves `null` in its
 * place (`tl`, `br`, `ext`, the values of a cell position, `picture`); the cell of a position is clamped to the grid.
 */
export interface DrawingAnchorModel {
    range: ImageRangeModel;
    picture?: DrawingPictureModel;
    /** prepare: the element, from the range (`br`: two-cell). */
    anchorType?: "xdr:twoCellAnchor" | "xdr:oneCellAnchor";
    /** reconcile: the workbook medium of the picture. */
    medium?: MediaModel;
}

// ---------------------------------------------------------------------------------------------------------------------
// Tables

/** A `<customFilter val operator>` of a filter column (table/custom-filter-xform). */
export interface CustomFilterModel {
    val?: string;
    operator?: string;
}

/** A `<filter val>` of a filter column (table/filter-xform). */
export interface FilterModel {
    val?: string;
}

/**
 * A `<filterColumn colId hiddenButton>` of the auto filter of a table (table/filter-column-xform): `colId` is the
 * 0-based index of the column in the table (set by prepare; a string as written).
 */
export interface FilterColumnModel {
    colId?: string;
    filterButton?: boolean;
    customFilters?: CustomFilterModel[];
}

/**
 * A table column as the xforms see it: index.d.ts `TableColumnProperties` (its `style` is the dxf of the column) plus
 * the write-side ids (`id`: 1-based, set by table-column-xform prepare; `colId`: see {@link FilterColumnModel}; `dxfId`:
 * set by worksheet-xform prepare) and the parse-side `dxfId` (a string, resolved to `style` by table-xform reconcile).
 */
export type TableColumnModel = Omit<TableColumnProperties, "style"> &
    FilterColumnModel & {
        /** Document side: the cached result of the totals row formula (written in the totals row cell). */
        totalsRowResult?: FormulaResult;
        id?: number;
        dxfId?: number | string;
        style?: DxfModel;
    };

/**
 * A table (src/doc/table.ts `Table.model`, table-xform). Write: the validated document table plus `tableRef`,
 * `autoFilterRef`, `tl`, then `target` / `id` (xlsx.ts) and `rId` (worksheet-xform). Parse: `{ name, displayName,
 * tableRef, totalsRow, headerRow, columns, autoFilterRef?, style }`, without `ref` / `rows` / `tl`; `style` is `null`
 * without `<tableStyleInfo>`; a malformed file without `<tableColumns>` has `columns: null` (unvalidated).
 */
export interface TableModel {
    name: string;
    displayName?: string;
    ref?: string;
    tableRef?: string;
    autoFilterRef?: string;
    headerRow?: boolean;
    totalsRow?: boolean;
    style?: TableStyleProperties | null;
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
    /** The manual page breaks between rows (`<rowBreaks>`). */
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
    // the parts of the worksheet absent from the file are null (the model of a reset xform)
    dimensions?: string | null;
    cols?: ColModel[] | null;
    /** reconcile: compacted (no holes), every row and cell with its number / address. */
    rows?: RowModel<ParsedCellModel>[] | null;
    mergeCells?: string[] | null;
    hyperlinks?: SheetHyperlinkModel[] | null;
    dataValidations?: DataValidationsModel | null;
    properties: WorksheetPropertiesModel;
    views?: SheetViewModel[] | null;
    pageSetup: PageSetupModel;
    headerFooter?: HeaderFooterModel | null;
    background?: SheetPartRefModel | null;
    drawing?: SheetPartRefModel | null;
    /** The `<tablePart>`s; reconcile: the table models of the worksheet. */
    tables?: (SheetPartRefModel | TableModel)[] | null;
    conditionalFormattings: ConditionalFormattingModel[];
    autoFilter?: AutoFilterModel;
    sheetProtection?: SheetProtectionModel;
    legacyDrawing?: SheetPartRefModel;
    /** reconcile: the images, from the drawing and background. */
    media?: WorksheetMediumModel[];
    comments?: NoteModel[];
}

/**
 * A reference to a part of a worksheet by relationship id: `<drawing>`, `<legacyDrawing>`, `<picture>` (background),
 * `<tablePart>`. `rId` is `undefined` when the `r:id` attribute is absent from a parsed element.
 */
export interface SheetPartRefModel {
    rId?: string;
}
