/// <reference types="node" />

import { EventEmitter } from "events";

/**
 * Binary data returned by ExcelJS (`writeBuffer()`): a Node Buffer in Node.js and Bun, the Buffer polyfill (a `Uint8Array`
 * with the Node Buffer methods) in the browser bundle. Its `buffer` is an `ArrayBuffer`, so that it can be passed to browser
 * APIs such as `new Blob([data])` (TypeScript 5.7+ DOM typings); it can be used as, or cast to, a Node `Buffer`.
 */
export type Buffer = globalThis.Buffer & { readonly buffer: ArrayBuffer };

export declare enum RelationshipType {
    None = 0,
    OfficeDocument = 1,
    Worksheet = 2,
    CalcChain = 3,
    SharedStrings = 4,
    Styles = 5,
    Theme = 6,
    Hyperlink = 7,
}

export declare enum DocumentType {
    Xlsx = 1,
}

/**
 * Common paper sizes of `pageSetup.paperSize`. A runtime object since 5.0 (it was a `const enum`, unusable with
 * `isolatedModules`).
 */
export declare enum PaperSize {
    Legal = 5,
    Executive = 7,
    A4 = 9,
    A5 = 11,
    B5 = 13,
    Envelope_10 = 20,
    Envelope_DL = 27,
    Envelope_C5 = 28,
    Envelope_B5 = 34,
    Envelope_Monarch = 37,
    Double_Japan_Postcard_Rotated = 82,
    K16_197x273_mm = 119,
}

export interface WorksheetViewCommon {
    /**
     * Sets the worksheet view's orientation to right-to-left, `false` by default
     */
    rightToLeft: boolean;

    /**
     * The currently selected cell
     */
    activeCell: string;

    /**
     * Shows or hides the ruler in Page Layout, `true` by default
     */
    showRuler: boolean;

    /**
     * Shows or hides the row and column headers (e.g. A1, B1 at the top and 1,2,3 on the left,
     * `true` by default
     */
    showRowColHeaders: boolean;

    /**
     * Shows or hides the gridlines (shown for cells where borders have not been defined),
     * `true` by default
     */
    showGridLines: boolean;

    /**
     * 	Percentage zoom to use for the view, `100` by default
     */
    zoomScale: number;

    /**
     * 	Normal zoom for the view, `100` by default
     */
    zoomScaleNormal: number;
}

export interface WorksheetViewNormal {
    /**
     * Controls the view state
     */
    state: "normal";

    /**
     * Presentation style: page break preview or page layout (the normal presentation when absent)
     */
    style?: "pageBreakPreview" | "pageLayout";
}

export interface WorksheetViewFrozen {
    /**
     * Where a number of rows and columns to the top and left are frozen in place.
     * Only the bottom left section will scroll
     */
    state: "frozen";

    /**
     * Presentation style
     */
    style?: "pageBreakPreview";

    /**
     * How many columns to freeze. To freeze rows only, set this to 0 or undefined
     */
    xSplit?: number;

    /**
     * How many rows to freeze. To freeze columns only, set this to 0 or undefined
     */
    ySplit?: number;

    /**
     * Which cell will be top-left in the bottom-right pane. Note: cannot be a frozen cell.
     * Defaults to first unfrozen cell
     */
    topLeftCell?: string;
}

export interface WorksheetViewSplit {
    /**
     * Where the view is split into 4 sections, each semi-independently scrollable.
     */
    state: "split";

    /**
     * Presentation style
     */
    style?: "pageBreakPreview" | "pageLayout";

    /**
     * How many points from the left to place the splitter.
     * To split vertically, set this to 0 or undefined
     */
    xSplit?: number;

    /**
     * How many points from the top to place the splitter.
     * To split horizontally, set this to 0 or undefined
     */
    ySplit?: number;

    /**
     * Which cell will be top-left in the bottom-right pane
     */
    topLeftCell?: string;

    /**
     * Which pane will be active
     */
    activePane?: "topLeft" | "topRight" | "bottomLeft" | "bottomRight";
}

export type WorksheetView = WorksheetViewCommon & (WorksheetViewNormal | WorksheetViewFrozen | WorksheetViewSplit);

/**
 * A view of `worksheet.views`: the fields of the normal, frozen and split views (see {@link WorksheetView}), all
 * optional. A view read from a file is not narrowed to the variant of its `state`; a {@link WorksheetView} can be
 * assigned. A view read from a file has every field of {@link WorksheetViewCommon}, `workbookViewId`, `state`, and
 * `xSplit` / `ySplit` / `topLeftCell` for a frozen or split view.
 */
export interface WorksheetViewModel extends Partial<WorksheetViewCommon> {
    /** The view state: `"normal"` (the default: any other value is written as normal), `"frozen"` or `"split"`. */
    state?: WorksheetView["state"];
    /** Presentation style: page break preview or page layout (the normal presentation when absent). */
    style?: "pageBreakPreview" | "pageLayout";
    /**
     * Frozen: how many columns to freeze (0 or undefined: rows only). Split: the position of the vertical splitter, in
     * points (0 or undefined: no vertical split).
     */
    xSplit?: number;
    /**
     * Frozen: how many rows to freeze (0 or undefined: columns only). Split: the position of the horizontal splitter, in
     * points (0 or undefined: no horizontal split).
     */
    ySplit?: number;
    /** Frozen or split: the top-left cell of the bottom-right pane (frozen: the first unfrozen cell by default). */
    topLeftCell?: string;
    /** Split only: the active pane (the one of a frozen view follows from its splits). */
    activePane?: "topLeft" | "topRight" | "bottomLeft" | "bottomRight";
    /**
     * The index of the workbook view (window) this view belongs to: `0` when written without one. Read from a file:
     * `NaN` when the attribute is absent (malformed file).
     */
    workbookViewId?: number;
    /** Whether the tab of the worksheet is selected (written only; not kept when a file is read). */
    tabSelected?: boolean;
}

/**
 * The style option of `addRow` / `insertRow` and their plural forms: `"i"` copies the style of the row above (`"i+"`:
 * also its empty cells), `"o"` the style of the row below (`"o+"`; `insertRow` / `insertRows` only, ignored by
 * `addRow`), `"n"` (the default) none.
 */
export type RowStyleOption = "n" | "i" | "i+" | "o" | "o+";

/**
 * A window of the workbook (`workbook.views`, `<workbookView>`). On write, a missing position is 0, 0 and a missing size
 * 12000 x 24000; on read, a missing position is 0, 0, a missing size 25000 x 10000 and a missing visibility `"visible"`.
 */
export interface WorkbookView {
    /** The horizontal position of the window, in twips. */
    x: number;
    /** The vertical position of the window, in twips. */
    y: number;
    /** The width of the window, in twips. */
    width: number;
    /** The height of the window, in twips. */
    height: number;
    /** The 0-based index of the first worksheet tab shown; absent when the file does not set it. */
    firstSheet?: number;
    /** The 0-based index of the active worksheet; absent when the file does not set it. */
    activeTab?: number;
    /** `"visible"`, `"hidden"` or `"veryHidden"`. */
    visibility: string;
}

/** The pattern of a {@link FillPattern} (ECMA-376 ST_PatternType). */
export type FillPatterns =
    | "none"
    | "solid"
    | "darkVertical"
    | "darkHorizontal"
    | "darkGrid"
    | "darkTrellis"
    | "darkDown"
    | "darkUp"
    | "lightVertical"
    | "lightHorizontal"
    | "lightGrid"
    | "lightTrellis"
    | "lightDown"
    | "lightUp"
    | "darkGray"
    | "mediumGray"
    | "lightGray"
    | "gray125"
    | "gray0625";

/** A pattern fill: the pattern drawn with `fgColor` over `bgColor` (`"solid"`: `fgColor` only). */
export interface FillPattern {
    type: "pattern";
    pattern: FillPatterns;
    /** The color of the pattern (of the whole cell for `"solid"`). */
    fgColor?: Partial<Color>;
    /** The background color under the pattern. */
    bgColor?: Partial<Color>;
}

/** A color of a gradient fill, at a position (0 to 1) of the gradient. */
export interface GradientStop {
    position: number;
    /**
     * The color. A stop without color is written with an empty `<color/>` (no color attribute); a stop read from a file
     * whose `<color>` has no color attribute has none.
     */
    color?: Partial<Color>;
}

/**
 * A linear gradient fill. A gradient read from a file without `degree` (0: left to right) has neither `gradient` nor
 * `degree`.
 */
export interface FillGradientAngle {
    type: "gradient";
    /** `"angle"`; absent on a gradient read from a file without `degree`. */
    gradient?: "angle";

    /**
     * For 'angle' gradient, specifies the direction of the gradient. 0 (the default, when absent) is from the left to
     * the right. Values from 1 - 359 rotates the direction clockwise
     */
    degree?: number;

    /**
     * Specifies the gradient colour sequence. Is an array of objects containing position and
     * color starting with position 0 and ending with position 1.
     * Intermediary positions may be used to specify other colours on the path.
     */
    stops: GradientStop[];
}

/** A path gradient fill: from a rectangle (`center`) to the edges of the cell. */
export interface FillGradientPath {
    type: "gradient";
    gradient: "path";

    /**
     * For 'path' gradient. Specifies the relative coordinates (0 to 1) of the rectangle the path starts from: `left` and
     * `top`, and `right` / `bottom` when they differ from `left` / `top` (`left` / `top` are written for them when they
     * are absent).
     */
    center: { left: number; top: number; right?: number; bottom?: number };

    /**
     * Specifies the gradient colour sequence. Is an array of objects containing position and
     * color starting with position 0 and ending with position 1.
     * Intermediary positions may be used to specify other colours on the path.
     */
    stops: GradientStop[];
}

/** The fill of a cell: a pattern or a gradient. */
export type Fill = FillPattern | FillGradientAngle | FillGradientPath;

/** A font. A font read from a file has only the properties the file sets. */
export interface Font {
    name: string;
    /** The size in points (e.g. 10.5). */
    size: number;
    /** The font family (1 Roman, 2 Swiss, 3 Modern, 4 Script, 5 Decorative). */
    family: number;
    scheme: "minor" | "major" | "none";
    /** The character set (e.g. 0 ANSI, 1 default, 128 Shift JIS). */
    charset: number;
    color: Partial<Color>;
    bold: boolean;
    italic: boolean;
    /** `true` is a single underline; read from a file, `true` or the underline style of the file. */
    underline: boolean | "none" | "single" | "double" | "singleAccounting" | "doubleAccounting";
    /** `"baseline"` (the default) is read from a file that sets it. */
    vertAlign: "superscript" | "subscript" | "baseline";
    strike: boolean;
    outline: boolean;
    /** Condensed text (a legacy Macintosh setting, read and written). */
    condense: boolean;
    /** Extended text (a legacy Macintosh setting, read and written). */
    extend: boolean;
    /** Shadowed text (a legacy Macintosh setting, read and written). */
    shadow: boolean;
}

/** The style of a border edge (ECMA-376 ST_BorderStyle, without `"none"`). */
export type BorderStyle =
    | "thin"
    | "dotted"
    | "hair"
    | "medium"
    | "double"
    | "thick"
    | "dashed"
    | "dashDot"
    | "dashDotDot"
    | "slantDashDot"
    | "mediumDashed"
    | "mediumDashDotDot"
    | "mediumDashDot";

/**
 * A color: `argb`, `theme` or `indexed`, each with an optional `tint`; when several are set, only the first of them in this
 * order is written; a color with none of them (`{ auto: true }`, or `{}`) is written as the automatic color (with its
 * `tint`, if any).
 */
export interface Color {
    /**
     * Hex string for alpha-red-green-blue e.g. FF00FF00
     */
    argb: string;

    /**
     * Choose a theme by index
     */
    theme: number;

    /** Lightens (0 to 1) or darkens (-1 to 0) the color (`argb`, `theme`, `indexed` or `auto`). */
    tint: number;

    /** A color of the legacy indexed palette (64 is the system foreground color). */
    indexed: number;

    /**
     * The automatic color (`<color auto="1"/>`): the system color of the context (e.g. the window text color). Read from a
     * file as `{ auto: true }`; a color without `argb`, `theme` and `indexed` is written as automatic.
     */
    auto: boolean;
}

/** A border edge: a style and its color. An edge read from a file with a color but no style is `{ color }`. */
export interface Border {
    style: BorderStyle;
    color: Partial<Color>;
}

/** The diagonal edge of a border: its line goes up (bottom-left to top-right) and / or down. */
export interface BorderDiagonal extends Border {
    up: boolean;
    down: boolean;
}

/** The border of a cell, edge by edge. */
export interface Borders {
    top: Partial<Border>;
    left: Partial<Border>;
    bottom: Partial<Border>;
    right: Partial<Border>;
    diagonal: Partial<BorderDiagonal>;
    /** The color of the edges that have a style but no color (written only: a border read from a file has none). */
    color: Partial<Color>;
}

export interface Margins {
    top: number;
    left: number;
    bottom: number;
    right: number;
    header: number;
    footer: number;
}

export declare enum ReadingOrder {
    LeftToRight = 1,
    RightToLeft = 2,
}

/** The alignment of the content of a cell. */
export interface Alignment {
    horizontal: "left" | "center" | "right" | "fill" | "justify" | "centerContinuous" | "distributed";
    /** The vertical alignment (`"middle"` is the `center` of the file). */
    vertical: "top" | "middle" | "bottom" | "distributed" | "justify";
    wrapText: boolean;
    shrinkToFit: boolean;
    indent: number;
    readingOrder: "rtl" | "ltr";
    /**
     * The rotation in degrees (-90 to 90), or `"vertical"` (stacked letters). Read from a file, `undefined` for a value
     * outside those.
     */
    textRotation: number | "vertical";
}

/** The protection of a cell (effective when the worksheet is protected). */
export interface Protection {
    /** Whether the cell cannot be edited (`true` by default). */
    locked: boolean;
    /** Whether the formula of the cell is hidden (`false` by default). */
    hidden: boolean;
}

/** The style of a cell, a row or a column. */
export interface Style {
    /** The number format code (`"0.00%"`, `"dd/mm/yyyy"`...). */
    numFmt: string;
    font: Partial<Font>;
    alignment: Partial<Alignment>;
    protection: Partial<Protection>;
    border: Partial<Borders>;
    fill: Fill;
}

/**
 * A differential format: the style that a conditional formatting rule or a table column applies over the style of the
 * cells. Read from a file, every part is a key: `null` when the file does not set it; `numFmt` is `undefined` for a
 * format code longer than 255 characters; `fill` may be a pattern without `pattern` (Excel writes the pattern fills of
 * these formats so, the color is `bgColor`) or `{}` (a `<fill>` without pattern nor gradient). Writing the workbook sets
 * `numFmtId` on a format that has a `numFmt` (it mutates the object).
 */
export interface DifferentialStyle {
    numFmt?: string | null;
    /** The id of `numFmt`, set when the workbook is written. */
    numFmtId?: number;
    font?: Partial<Font> | null;
    alignment?: Partial<Alignment> | null;
    protection?: Partial<Protection> | null;
    border?: Partial<Borders> | null;
    fill?: Fill | (Omit<FillPattern, "pattern"> & { pattern?: FillPatterns }) | { type?: undefined } | null;
}

/** The comparison of a data validation of a number, a date or a text length (`"between"` by default). */
export type DataValidationOperator =
    "between" | "notBetween" | "equal" | "notEqual" | "greaterThan" | "lessThan" | "greaterThanOrEqual" | "lessThanOrEqual";

/** The kind of a data validation; `"any"` (any value: an input message only) is the default when reading a file. */
export type DataValidationType = "any" | "list" | "whole" | "decimal" | "date" | "textLength" | "custom";

/**
 * The data validations of a worksheet (`worksheet.dataValidations`).
 */
export interface DataValidations {
    /**
     * The validations by cell address (`'B2'`) or by range (`'B1:B1048576'`: ranges read from a file once 100,000 cells
     * have been expanded, whole columns or rows added with `add('B:B', ...)`). A key may hold `undefined` once removed.
     * Write a range key through a fresh `worksheet.dataValidations.model`, or use `add()`.
     */
    model: { [address: string]: DataValidation | undefined };
    /**
     * Sets the validation of a cell, a range (`'A1:B10'`) or whole columns or rows (`'A:A'`, `'1:3'`).
     */
    add(address: string, validation: DataValidation): DataValidation;
    /**
     * The validation of a cell, from its own key or from a range key that contains it.
     */
    find(address: string): DataValidation | undefined;
    /**
     * Removes the validation of a cell (splitting a range key around it) or a range key.
     */
    remove(address: string): void;
}

/** The validation of the values entered in a cell, with its input and error messages. */
export interface DataValidation {
    type: DataValidationType;
    /**
     * The formulas or values of the rule (`["A1:A5"]`, `[1, 10]`, `[new Date()]`, `["$B$1"]`, `["TODAY()"]`); absent for
     * `"any"`. A string is written as is (a formula), a number as is, a Date as its serial date; for `date`, a number is
     * read as milliseconds since 1970 (`new Date(number)`). Writing an invalid Date or a number that is not finite
     * throws. Read from a file: a number literal is a number for `whole`, `decimal` and `textLength` and a Date for
     * `date`; any other formula (a reference, a function...) is its text.
     */
    formulae?: (string | number | Date)[];
    /** Whether an empty cell is valid. */
    allowBlank?: boolean;
    /** The comparison (not written for `list` nor for the default `"between"`). */
    operator?: DataValidationOperator;
    /** The text of the error message. */
    error?: string;
    /** The title of the error message. */
    errorTitle?: string;
    /** The kind of the error message: `"stop"` (the default), `"warning"` or `"information"`. */
    errorStyle?: string;
    /** The text of the input message. */
    prompt?: string;
    /** The title of the input message. */
    promptTitle?: string;
    /** Whether the error message is shown for an invalid value. */
    showErrorMessage?: boolean;
    /** Whether the input message is shown when the cell is selected. */
    showInputMessage?: boolean;
}

export declare enum ErrorValue {
    NotApplicable = "#N/A",
    Ref = "#REF!",
    Name = "#NAME?",
    DivZero = "#DIV/0!",
    Null = "#NULL!",
    Value = "#VALUE!",
    Num = "#NUM!",
}

export interface CellErrorValue {
    error: "#N/A" | "#REF!" | "#NAME?" | "#DIV/0!" | "#NULL!" | "#VALUE!" | "#NUM!";
}

/**
 * A run of rich text: its text and its font. A run read from a malformed file (`<r>` without its required `<t>`) has no
 * `text`; a run without text is not written.
 */
export interface RichText {
    text: string;
    font?: Partial<Font>;
}

export interface CellRichTextValue {
    richText: RichText[];
}

/**
 * The value of a hyperlink cell: the value shown (`text`), the link target (`hyperlink`: a URL, or a place in the
 * workbook such as `"#Sheet1!A1"`) and an optional tooltip.
 */
export interface CellHyperlinkValue {
    /**
     * The value shown: any value of a cell, such as a number or a rich text (a hyperlink read from a file shows the value
     * of its cell, whatever its type); `cell.text` is its text, as for a cell of that value. It may be empty (`""`), not
     * null.
     */
    text: string | number | boolean | Date | CellErrorValue | CellRichTextValue;
    hyperlink: string;
    tooltip?: string;
}

/** The cached result of a formula. */
export type CellFormulaResult = number | string | boolean | Date | CellErrorValue;

/**
 * The value of a formula cell that has a formula of its own: a single formula, the master of shared formulas
 * (`shareType: "shared"`) or an array formula (`shareType: "array"`), with its cached result.
 */
export interface CellFormulaValue {
    /** The formula, without `=`. */
    formula: string;
    result?: CellFormulaResult;
    /** The master of shared formulas (`"shared"`) or an array formula (`"array"`); absent for a single formula. */
    shareType?: "shared" | "array";
    /** The range of a shared or array formula master (`"A1:A10"`). */
    ref?: string;
}

/**
 * The value of a clone of a shared formula: the address of its master (`sharedFormula`) and its cached result. The
 * formula translated to the cell is `cell.formula`.
 */
export interface CellSharedFormulaValue {
    /** The address of the master of the shared formula (`"A1"`). */
    sharedFormula: string;
    /** Absent from the value of a clone: `cell.formula` gives the master's formula translated to the cell. */
    readonly formula?: string;
    result?: CellFormulaResult;
}

export declare enum ValueType {
    Null = 0,
    Merge = 1,
    Number = 2,
    String = 3,
    Date = 4,
    Hyperlink = 5,
    Formula = 6,
    SharedString = 7,
    RichText = 8,
    Boolean = 9,
    Error = 10,
}

export declare enum FormulaType {
    None = 0,
    Master = 1,
    Shared = 2,
}

/**
 * The value of a cell read by the streaming reader when the shared strings are not cached (`sharedStrings: "emit"` or
 * `"ignore"`): the index of the string in the shared strings table.
 */
export interface CellSharedStringValue {
    readonly sharedString: number;
}

/**
 * The value of a cell. Reading gives `null` for an empty cell (never `undefined`); setting `undefined` or `null` empties
 * the cell.
 *
 * The kind of an object value is told by its marker properties, by their presence and type (an empty text or formula,
 * a shared string 0 count), in this order: a `text` (not null) and a string `hyperlink` (hyperlink), a string `formula`
 * or `sharedFormula` (formula; an empty formula is a formula cell without formula, `formulaType` None), an array
 * `richText`, a number `sharedString`, a string `error`. Any other object (outside this type: an untyped value, such as
 * one returned by the CSV `map` option) is held as a JSON value: its type is String, it is returned as is, and it is
 * written as its JSON text (read back as a string).
 */
export type CellValue =
    | null
    | number
    | string
    | boolean
    | Date
    | undefined
    | CellErrorValue
    | CellRichTextValue
    | CellHyperlinkValue
    | CellFormulaValue
    | CellSharedFormulaValue
    | CellSharedStringValue;

/** The margins of the box of a note. */
export interface CommentMargins {
    insetmode: "auto" | "custom";
    /**
     * The margins in cm: left, top, right, bottom. A note read from a file without (or with invalid) margins has none:
     * the VML defaults apply.
     */
    inset: number[];
}

export interface CommentProtection {
    locked: "True" | "False";
    lockText: "True" | "False";
}

export type CommentEditAs = "twoCells" | "oneCells" | "absolute";

export type ImageEditAs = "twoCell" | "oneCell" | "absolute";

/**
 * A note (comment) of a cell: its rich text, the margins and protection of its box, and how the box moves with the
 * cells. A note read from a file has all of them, and its VML `anchor`.
 */
export interface Comment {
    texts?: RichText[];
    margins?: Partial<CommentMargins>;
    protection?: Partial<CommentProtection>;
    editAs?: CommentEditAs;
    /**
     * Read from a file: the text of the VML `<x:Anchor>` of the box (`"col, dx, row, dy, col, dx, row, dy"`, 0-based
     * columns and rows). Informational: a note is always written with its box next to its cell.
     */
    readonly anchor?: string;
}

/** The note of a cell in its model (`cellModel.comment`). */
export interface CellNoteModel {
    type: "note";
    note: Comment;
}

/**
 * The model of a cell (`cell.model`): its address, its value type and the fields of that type (`value`; `text` /
 * `hyperlink` / `tooltip` of a hyperlink; `formula` / `sharedFormula` / `result` of a formula; `master` of a merged
 * cell), its style and its note.
 */
export interface CellModel {
    /** The cell reference (`"A1"`). */
    address: string;
    type: ValueType;
    style?: Partial<Style>;
    value?: number | string | boolean | Date | CellErrorValue | CellRichTextValue | CellSharedStringValue;
    /** Hyperlink: the value shown (on read, the value of the cell, whatever its type). */
    text?: number | string | boolean | Date | CellErrorValue | CellRichTextValue;
    hyperlink?: string;
    tooltip?: string;
    /** Merged cell: the address of its master. */
    master?: string;
    formula?: string;
    /** Clone of a shared formula: the address of its master. */
    sharedFormula?: string;
    /** Master of shared formulas (`"shared"`) or array formula (`"array"`). */
    shareType?: "shared" | "array";
    /** The range of a shared or array formula master. */
    ref?: string;
    result?: CellFormulaResult;
    comment?: CellNoteModel;
}

/**
 * The style properties of a cell, row or column (`cell.font`, `row.fill`...): `undefined` when the style does not set
 * them. Setting one copies the style object (styles may be shared between cells); on a row or a column it also sets the
 * property on the existing cells.
 */
export interface StyleProperties {
    numFmt: string | undefined;
    font: Partial<Font> | undefined;
    alignment: Partial<Alignment> | undefined;
    protection: Partial<Protection> | undefined;
    border: Partial<Borders> | undefined;
    fill: Fill | undefined;
}

/** A cell. Its address (`address`, `row`, `col`, `$col$row`) is read-only. */
export interface Cell extends StyleProperties, Readonly<Omit<Address, "sheetName">> {
    readonly worksheet: Worksheet;
    readonly workbook: Workbook;

    readonly effectiveType: ValueType;
    readonly isMerged: boolean;
    /** The master of the merge this cell belongs to, or the cell itself. */
    readonly master: Cell;
    readonly isHyperlink: boolean;
    /** The target of a hyperlink cell, `undefined` for other cells. */
    readonly hyperlink: string | undefined;
    /**
     * The text of the value, not formatted by `numFmt`: a string as is, the `toString()` of a number, a boolean or a
     * date, the code of an error, the runs of a rich text, the text of a hyperlink (as for a cell of that value), the
     * result of a formula; `""` for an empty cell.
     */
    readonly text: string;
    /** The text of the cell ({@link text}), HTML-escaped. */
    readonly html: string;
    readonly fullAddress: {
        sheetName: string;
        address: string;
        row: number;
        col: number;
    };
    model: CellModel;
    /**
     * Assign (or get) a name for a cell (will overwrite any other names that cell had); `undefined` when it has none.
     */
    get name(): string | undefined;
    set name(value: string);

    /**
     * Assign (or get) an array of names for a cell (cells can have more than one name): setting them replaces the names
     * of the cell.
     */
    get names(): string[];
    set names(value: readonly string[]);

    /**
     * Cells can define what values are valid or not and provide prompting to the user to help guide them; `undefined`
     * when the cell has no validation.
     */
    get dataValidation(): DataValidation | undefined;
    set dataValidation(value: DataValidation);

    /**
     * Value of the cell (see {@link CellValue}): a merged cell gives and sets the value of its master. Any other object
     * than a CellValue is held as a JSON value, returned as is.
     */
    value: CellValue;

    /**
     * comment of the cell, `undefined` when it has none
     */
    get note(): string | Comment | undefined;
    set note(value: string | Comment);

    /**
     * convenience getter to access the formula (`undefined` for a cell without formula)
     */
    readonly formula: string | undefined;

    /**
     * convenience getter to access the formula result (`undefined` for a cell without formula or without result)
     */
    readonly result: CellFormulaResult | undefined;

    /**
     * The type of the cell's value
     */
    readonly type: ValueType;

    /**
     * The type of the cell's formula: Master or Shared for a formula cell (None for a formula read from a malformed file
     * with neither formula nor master), `undefined` for other cells
     */
    readonly formulaType: FormulaType | undefined;

    /**
     * The styles of the cell
     */
    style: Partial<Style>;

    addName(name: string): void;

    /**
     * Remove a name from a cell
     */
    removeName(name: string): void;
    removeAllNames(): void;

    destroy(): void;
    /** The value as written to CSV: a number for a boolean (1 / 0). */
    toCsvString(): string | number;
    /**
     * Merges this cell into `master`: its value is dropped, and it takes the master's style unless `ignoreStyle`. Low
     * level: the merge is not recorded in the worksheet (not written); use `worksheet.mergeCells`.
     */
    merge(master: Cell, ignoreStyle?: boolean): void;
    /** Unmerges a merged cell: it becomes empty, and its style is reset from its row and column unless `ignoreStyle`. */
    unmerge(ignoreStyle?: boolean): void;
    /** Whether this cell is merged into `master`. */
    isMergedTo(master: Cell): boolean;
    toString(): string;
}

/** The model of a row (`row.model`): its number, its non-empty cells and its properties. */
export interface RowModel {
    cells: CellModel[];
    number: number;
    /** The first / last column with a cell (0 for a row without cells). */
    min?: number;
    max?: number;
    height?: number;
    style?: Partial<Style> | null;
    hidden?: boolean;
    outlineLevel?: number;
    collapsed?: boolean;
}

/** A row of a worksheet. Its cells are created on demand (`getCell`); a new cell takes the style of the row. */
export interface Row extends StyleProperties {
    readonly worksheet: Worksheet;
    readonly hasValues: boolean;
    /**
     * The min and max column numbers of the non-empty cells, or null if the row is empty
     */
    readonly dimensions: { min: number; max: number } | null;
    /**
     * The model of the row; `null` for a row without height and without cell. Setting it replaces the cells, height,
     * hidden state, outline level and style; it throws for the model of another row number.
     */
    get model(): RowModel | null;
    set model(value: RowModel);
    /**
     * Set a specific row height; absent (`undefined`) when the row has the default height
     */
    height?: number;

    /**
     * Make row hidden (`undefined`: not hidden)
     */
    get hidden(): boolean;
    set hidden(value: boolean | undefined);

    /**
     * The style of the row: the style of the cells created in it. Setting a style property (`row.font = ...`) also sets
     * it on the existing cells.
     */
    style: Partial<Style>;

    /**
     * Get a row as a sparse array (index 0 is unused: column A is at index 1)
     */
    get values(): CellValue[];
    /**
     * Set the row's values from a contiguous or sparse array, or from an object keyed by column key. The existing cells
     * are removed first; `null` / `undefined` empty the row.
     */
    set values(value: RowValues);

    /**
     * The outline level of the row: 0 when not set (setting `undefined` clears it)
     */
    get outlineLevel(): number;
    set outlineLevel(value: number | undefined);

    /**
     * The row number
     */
    readonly number: number;

    /**
     * Indicate the collapsed state based on outlineLevel
     */
    readonly collapsed: boolean;

    /**
     * Number of cells including empty ones
     */
    readonly cellCount: number;

    /**
     * Number of non-empty cells
     */
    readonly actualCellCount: number;

    /**
     * Get cell by number, column letter or column key
     */
    getCell(indexOrKey: number | string): Cell;

    /** The cell of a column (1-based number) if it exists, without creating it. */
    findCell(colNumber: number): Cell | undefined;

    /**
     * The cell at a location of this row, created if needed: `address` is its address (`"B2"`) and `col` its 1-based
     * column number (`row` is not read).
     */
    getCellEx(address: { readonly address: string; readonly col: number; readonly row?: number }): Cell;

    /**
     * Iterate over all non-null cells in a row
     */
    eachCell(callback: (cell: Cell, colNumber: number) => void): void;

    /**
     * Iterate over the non-null cells in a row; with `includeEmpty: true`, over all the cells up to the last one
     * (creating the missing ones)
     */
    eachCell(opt: { readonly includeEmpty?: boolean } | null | undefined, callback: (cell: Cell, colNumber: number) => void): void;

    /**
     * Removes `count` cells from column `start` (1-based) and inserts the `insert` values there: the cells to the right
     * are shifted (left or right)
     *
     * Note: this operation will not affect other rows
     */
    splice(start: number, count: number, ...insert: CellValue[]): void;

    /**
     * Commit a completed row (and the rows before it) to stream. In the streaming writer, returns a promise that resolves
     * once the output can take more data: awaiting it keeps memory bounded (backpressure). No effect on a document worksheet.
     */
    commit(): void | Promise<void>;
    destroy(): void;
    /**
     * Adds a manual page break after this row (to `worksheet.rowBreaks`), from column `lft` to column `rght` (1-based;
     * absent or 0: from the first column, to the last one).
     *
     * @throws Error for a column that is not one of the worksheet (an integer from 1 to 16,384)
     */
    addPageBreak(lft?: number, rght?: number): void;
}

/** A column of a worksheet: its header, key, width, style, outline level. */
export interface Column extends StyleProperties {
    /**
     * Can be a string to set one row high header or an array to set multi-row high header
     */
    header?: string | string[];

    /**
     * The name of the properties associated with this column in each row
     */
    key?: string;

    /**
     * The width of the column
     */
    width?: number;

    /**
     * The outline level of the column: 0 when not set (setting `undefined` clears it)
     */
    get outlineLevel(): number;
    set outlineLevel(value: number | undefined);

    /**
     * Hides the column
     */
    hidden: boolean;

    /**
     * Styles applied to the column
     */
    style: Partial<Style>;

    /**
     * The cell values in the column, as a sparse array indexed by row number (index 0 unused)
     */
    get values(): CellValue[];
    /**
     * Set the values of the column from a sparse array indexed by row number, or a contiguous array (index 0 set) for
     * rows 1, 2...; `null` / `undefined` change nothing.
     */
    set values(value: readonly CellValue[] | null | undefined);

    /**
     * Column letter key
     */
    readonly letter: string;
    readonly number: number;
    readonly worksheet: Worksheet;
    readonly isCustomWidth: boolean;
    /** The header text of each header row (`undefined` for a row without header). */
    readonly headers: readonly (string | undefined)[];
    readonly isDefault: boolean;
    readonly headerCount: number;

    toString(): string;
    equivalentTo(other: Column): boolean;

    /**
     * indicate the collapsed state based on outlineLevel
     */
    readonly collapsed: boolean;

    /**
     * Iterate over all current cells in this column
     */
    eachCell(callback: (cell: Cell, rowNumber: number) => void): void;

    /**
     * Iterate over the cells of this column in the rows that have values; with `includeEmpty: true`, in every row up to
     * the last one (creating the missing rows and cells)
     */
    eachCell(opt: { readonly includeEmpty?: boolean } | null | undefined, callback: (cell: Cell, rowNumber: number) => void): void;

    /**
     * The definition of the column (what `worksheet.columns = [...]` sets): setting it sets the key, width (default 9),
     * outline level, style, header and hidden state; `null` / `undefined` reset the column (no header, key, width, style
     * nor outline level).
     */
    get defn(): ColumnDefinition;
    set defn(value: ColumnDefinition | null | undefined);
}

/** The definition of a column: an element of `worksheet.columns = [...]`, and `column.defn`. */
export interface ColumnDefinition {
    header?: string | string[];
    key?: string;
    width?: number;
    /** `null`: no style. */
    style?: Partial<Style> | null;
    hidden?: boolean;
    outlineLevel?: number;
}
export interface PageSetup {
    /**
     * Whitespace on the borders of the page. Units are inches.
     *
     * `null` for a worksheet read from a file without `<pageMargins>` (no margins are then written).
     */
    margins: Margins | null;

    /**
     * Orientation of the page - i.e. taller (`'portrait'`) or wider (`'landscape'`).
     *
     * `'portrait'` by default
     */
    orientation: "portrait" | "landscape";

    /**
     * Horizontal Dots per Inch. Default value is 4294967295
     */
    horizontalDpi: number;

    /**
     * Vertical Dots per Inch. Default value is 4294967295
     */
    verticalDpi: number;

    /**
     * Whether to use fitToWidth and fitToHeight or scale settings.
     *
     * Default is based on presence of these settings in the pageSetup object - if both are present,
     * scale wins (i.e. default will be false)
     */
    fitToPage: boolean;

    /**
     * How many pages wide the sheet should print on to. Active when fitToPage is true
     *
     * Default is 1
     */
    fitToWidth: number;

    /**
     * How many pages high the sheet should print on to. Active when fitToPage is true
     *
     * Default is 1
     */
    fitToHeight: number;

    /**
     * Percentage value to increase or reduce the size of the print. Active when fitToPage is false
     *
     * Default is 100
     */
    scale: number;

    /**
     * Which order to print the pages.
     *
     * Default is `downThenOver`
     */
    pageOrder: "downThenOver" | "overThenDown";

    /**
     * Print without colour
     *
     * false by default
     */
    blackAndWhite: boolean;

    /**
     * Print with less quality (and ink)
     *
     * false by default
     */
    draft: boolean;

    /**
     * Where to print the comments: at the end of the sheet, as displayed, or not at all (`"None"`, the default; a file
     * may also give its own spelling, `"none"`). `"asDisplyed"` (sic), the value of earlier versions, is still accepted
     * and written as `asDisplayed`.
     */
    cellComments: "atEnd" | "asDisplayed" | "None" | "none";

    /**
     * Where to show errors
     *
     * Default is `displayed`
     */
    errors: "dash" | "blank" | "NA" | "displayed";

    /**
     * 	What paper size to use (see below)
     *
     * | Name                          | Value       |
     * | ----------------------------- | ---------   |
     * | Letter                        | `undefined` |
     * | Legal                         |  `5`        |
     * | Executive                     |  `7`        |
     * | A4                            |  `9`        |
     * | A5                            |  `11`       |
     * | B5 (JIS)                      |  `13`       |
     * | Envelope #10                  |  `20`       |
     * | Envelope DL                   |  `27`       |
     * | Envelope C5                   |  `28`       |
     * | Envelope B5                   |  `34`       |
     * | Envelope Monarch              |  `37`       |
     * | Double Japan Postcard Rotated |  `82`       |
     * | 16K 197x273 mm                |  `119`      |
     */
    paperSize: PaperSize | (number & {});

    /**
     * Whether to show the row numbers and column letters, `false` by default
     */
    showRowColHeaders: boolean;

    /**
     * Whether to show grid lines, `false` by default
     */
    showGridLines: boolean;

    /**
     * Which number to use for the first page; absent (the default) for automatic numbering. When set, the pages are
     * numbered from it unless `useFirstPageNumber` is `false`. A worksheet read from a file has none when the file gives
     * none and does not use one; `1` when it uses one without giving it (the default of the format).
     */
    firstPageNumber: number;

    /**
     * Whether the pages are numbered from `firstPageNumber` (written as `useFirstPageNumber`): when absent, they are as
     * soon as `firstPageNumber` is set. A worksheet read from a file that gives a `firstPageNumber` has `false` when the
     * file does not use it (automatic numbering, as Excel does), `true` when it does.
     */
    useFirstPageNumber: boolean;

    /**
     * Whether to print with the printer's default settings (written only when `true`; read from a file: `true` only
     * when the file says so)
     */
    usePrinterDefaults: boolean;

    /**
     * How many copies to print. `1` by default
     */
    copies: number;

    /**
     * 	Whether to center the sheet data horizontally, `false` by default
     */
    horizontalCentered: boolean;

    /**
     * 	Whether to center the sheet data vertically, `false` by default
     */
    verticalCentered: boolean;

    /**
     * Set Print Area for a sheet, e.g. `'A1:G20'`
     */
    printArea: string;

    /**
     * Repeat specific rows on every printed page, e.g. `'1:3'`
     */
    printTitlesRow: string;

    /**
     * Repeat specific columns on every printed page, e.g. `'A:C'`
     */
    printTitlesColumn: string;
}

export interface HeaderFooter {
    /**
     * Set the value of differentFirst as true, which indicates that headers/footers for first page are different from the other pages, `false` by default
     */
    differentFirst: boolean;
    /**
     * Set the value of differentOddEven as true, which indicates that headers/footers for odd and even pages are different, `false` by default
     */
    differentOddEven: boolean;
    /**
     * Set header string for odd pages, could format the string and `null` by default
     */
    oddHeader: string;
    /**
     * Set footer string for odd pages, could format the string and `null` by default
     */
    oddFooter: string;
    /**
     * Set header string for even pages, could format the string and `null` by default
     */
    evenHeader: string;
    /**
     * Set footer string for even pages, could format the string and `null` by default
     */
    evenFooter: string;
    /**
     * Set header string for the first page, could format the string and `null` by default
     */
    firstHeader: string;
    /**
     * Set footer string for the first page, could format the string and `null` by default
     */
    firstFooter: string;
}

export type AutoFilter =
    | string
    | {
          from: string | { row: number; column: number };
          to: string | { row: number; column: number };
      };

export interface WorksheetProtection {
    objects: boolean;
    scenarios: boolean;
    selectLockedCells: boolean;
    selectUnlockedCells: boolean;
    formatCells: boolean;
    formatColumns: boolean;
    formatRows: boolean;
    insertColumns: boolean;
    insertRows: boolean;
    insertHyperlinks: boolean;
    deleteColumns: boolean;
    deleteRows: boolean;
    sort: boolean;
    autoFilter: boolean;
    pivotTables: boolean;
    spinCount: number;
}

/**
 * The protection of a worksheet as `worksheet.sheetProtection` holds it: the {@link WorksheetProtection} options given
 * to `protect()` (or read from the file), plus `sheet` (`true` once protected) and, with a password, its SHA-512 hash
 * (`algorithmName` `"SHA-512"`, base64 `hashValue` and `saltValue`, `spinCount` iterations).
 */
export type WorksheetProtectionModel = Partial<WorksheetProtection> & {
    /** Whether the sheet is protected. */
    sheet?: boolean;
    /** The hash algorithm of the password (`"SHA-512"`). */
    algorithmName?: string;
    /** The hash of the password, base64. */
    hashValue?: string;
    /** The random salt of the hash, base64. */
    saltValue?: string;
};

/**
 * An image to add to a workbook (`workbook.addImage`): its extension and its data, given by exactly one of `filename`,
 * `buffer` or `base64` (checked in this order when the workbook is written).
 */
export interface Image {
    /**
     * The extension of the image file (`xl/media/image1.png`), also its content type (`image/<extension>`): Excel shows
     * `jpeg`, `png`, `gif`... Writing rejects an extension that is not 1 to 5 letters or digits.
     */
    extension: "jpeg" | "png" | "gif" | (string & {});
    /** The image as base64: plain, or a data URL (`data:image/png;base64,...`: the text up to the first comma is dropped). */
    base64?: string;
    /** The path of the image file, read when the workbook is written (Node.js and Bun only). */
    filename?: string;
    /** The bytes of the image (a Node Buffer is a Uint8Array). */
    buffer?: ArrayBuffer | Uint8Array;
}
export interface IAnchor {
    col: number;
    row: number;
    nativeCol: number;
    nativeRow: number;
    nativeColOff: number;
    nativeRowOff: number;
}
/**
 * The corner of an image (`tl` / `br` of `worksheet.getImages()` ranges): the anchor cell (`nativeCol` / `nativeRow`,
 * 0-based) and the offset inside it (in EMU), or the same position as a fractional 0-based `col` / `row`. Not exported
 * at runtime: use a plain {@link IAnchor} position to place an image.
 */
export interface Anchor extends IAnchor {
    /** The width of the anchor column, in EMU (the default for a column without custom width). */
    readonly colWidth: number;
    /** The height of the anchor row, in EMU (the default for a row without height). */
    readonly rowHeight: number;
    /** The worksheet that gives the sizes of the cells. */
    worksheet: Worksheet;
}
/**
 * A corner of an image given to `worksheet.addImage`: a fractional 0-based `col` / `row` (`{ col: 1.5, row: 0 }`: the
 * middle of column B, row 1), or the anchor cell (`nativeCol` / `nativeRow`, 0-based) and the offsets inside it, in EMU
 * (0 by default). An {@link Anchor} of `getImages()` is both.
 */
export type ImageAnchorPosition =
    { col: number; row: number } | { nativeCol: number; nativeRow: number; nativeColOff?: number; nativeRowOff?: number };

/** The position of an image given by two corners (a two-cell anchor). */
export interface ImageRange {
    /** The top-left corner. */
    tl: ImageAnchorPosition;
    /** The bottom-right corner. */
    br: ImageAnchorPosition;
}

/** An image of a worksheet (`worksheet.getImages()`). */
export interface WorksheetImage {
    /** Always `"image"` for `getImages()` (`"background"` is the background image). */
    type: "image" | "background";
    /** The id given by `workbook.addImage`. */
    imageId: number | string;
    /** Where the image is: its top-left corner, and its bottom-right corner or its size. */
    range?: {
        tl: Anchor;
        br?: Anchor;
        ext?: { width: number; height: number };
        editAs?: ImageEditAs;
        /**
         * The link of the image. Its `hyperlink` is absent only for an image read from a malformed file whose link has no
         * relationship (`<a:hlinkClick>` without `r:id`): the type does not say so.
         */
        hyperlinks?: ImageHyperlinkValue;
    };
}

/** The position of an image given by its top-left corner and its size (a one-cell anchor). */
export interface ImagePosition {
    /** The top-left corner. */
    tl: ImageAnchorPosition;
    /** The size of the image, in pixels. */
    ext: { width: number; height: number };
}

export interface ImageHyperlinkValue {
    hyperlink: string;
    tooltip?: string;
}

export interface Range extends Location {
    /** The worksheet name of a reference such as `Sheet1!A1:B2`, `undefined` without one. */
    sheetName: string | undefined;

    /** The top-left cell address (`A1`). */
    readonly tl: string;
    /** The top-left cell address, absolute (`$A$1`). */
    readonly $t$l: string;

    /** The bottom-right cell address (`B2`). */
    readonly br: string;
    /** The bottom-right cell address, absolute (`$B$2`). */
    readonly $b$r: string;

    /** The reference, with the sheet name if any (`Sheet1!A1:B2`). */
    readonly range: string;
    /** The absolute reference (`Sheet1!$A$1:$B$2`). */
    readonly $range: string;

    /** The reference, as a single address for a single cell (`A1:B2`, `A1`; with the sheet name if any). */
    readonly shortRange: string;
    /** The absolute reference, as a single address for a single cell (`$A$1:$B$2`, `$A$1`). */
    readonly $shortRange: string;

    /** The number of cells of the range. */
    readonly count: number;

    /**
     * Sets the range from the list of arguments of a range (as `worksheet.mergeCells` takes them), given as ONE array:
     * `decode(["A1:B2"])`, `decode(["A1", "B2"])`, `decode([1, 1, 2, 2])`.
     */
    decode(args: RangeArguments): void;

    expand(top: number, left: number, bottom: number, right: number): void;

    /** Grows the range to include the non-empty cells of a row (nothing for an empty row or no row). */
    expandRow(row: Row | null | undefined): void;

    expandToAddress(addressStr: string): void;

    toString(): string;

    /** Whether the ranges share a cell (ranges of two different named sheets never do). */
    intersects(other: Range | (Location & { sheetName?: string })): boolean;

    contains(addressStr: string): boolean;

    /** Whether the range contains a cell (1-based numbers; a cell of another named sheet is not contained). */
    containsEx(address: { row: number; col: number; sheetName?: string }): boolean;

    /**
     * Sets the bounds from two addresses (`setTLBR("B2", "A1")`; the third argument is ignored), or from four numbers and
     * a sheet name. The bounds are normalised (top <= bottom, left <= right).
     */
    setTLBR(tl: string, br: string, sheetName?: string): void;
    setTLBR(top: number, left: number, bottom: number, right: number, sheetName?: string): void;

    /** Calls `callback` for each cell of the range, column by column (top to bottom in each column). */
    forEachAddress(callback: (address: string, row: number, col: number) => void): void;
}

/** The header and footer of a worksheet: a {@link HeaderFooter} whose fields may also be `null` (the defaults). */
export type WorksheetHeaderFooter = { [K in keyof HeaderFooter]?: HeaderFooter[K] | null };

/** A manual page break after a row (`row.addPageBreak()`; `worksheet.rowBreaks`). */
export interface RowBreak {
    /** The row number (the break is after it). */
    id: number;
    /** The last column of the break (0-based). */
    max: number;
    /** The first column of the break (0-based); absent for the first column. */
    min?: number;
    /** `1`: a manual break (`row.addPageBreak()`); `0`: an automatic one (read from a file). */
    man: number;
}

/**
 * A medium of a worksheet as `worksheet.model` holds it: the background image, or an image placed over the cells
 * (anchor cells 0-based, offsets in EMU; `ext` in pixels). `imageId` is the id of `workbook.addImage`.
 */
export type WorksheetMediumModel =
    | {
          type: "background";
          /** The id of the image in the workbook. */
          imageId: number | string;
      }
    | {
          type: "image";
          /** The id of the image in the workbook. */
          imageId: number | string;
          /** Where the image is: its top-left corner, and its bottom-right corner (two-cell anchor) or its size. */
          range: {
              tl: Pick<IAnchor, "nativeCol" | "nativeColOff" | "nativeRow" | "nativeRowOff">;
              br?: Pick<IAnchor, "nativeCol" | "nativeColOff" | "nativeRow" | "nativeRowOff">;
              ext?: { width: number; height: number };
              editAs?: ImageEditAs;
          };
          /**
           * The hyperlink of the image (`hyperlink` is absent for an image read from a malformed file whose link has no
           * relationship).
           */
          hyperlinks?: { hyperlink?: string; tooltip?: string };
      };

/**
 * A table as `worksheet.model` holds it: the {@link TableProperties} given to `addTable` once validated, with the
 * ranges of the table (`tableRef`, e.g. `"A1:C5"`) and of its auto filter. A table read from a file has no `ref` nor
 * `rows` (its cells hold them), and a `null` style when the file has none.
 */
export interface WorksheetTableModel {
    /** The name of the table. */
    name: string;
    /** The display name. */
    displayName?: string;
    /** The top-left cell (a table added by `addTable`). */
    ref?: string;
    /** The range of the table, e.g. `"A1:C5"`. */
    tableRef?: string;
    /** The range of the auto filter of the table. */
    autoFilterRef?: string;
    /** Whether the table has a header row. */
    headerRow?: boolean;
    /** Whether the table has a totals row. */
    totalsRow?: boolean;
    /** The style of the table. */
    style?: TableStyleProperties | null;
    /** The columns. */
    columns: WorksheetTableColumnModel[];
    /** The rows of data (a table added by `addTable`). */
    rows?: CellValue[][];
}

/**
 * A column of a table as `worksheet.model` holds it: the {@link TableColumnProperties} given to `addTable`, or read from
 * a file (then its `style` is the differential format of the column).
 */
export interface WorksheetTableColumnModel extends Omit<TableColumnProperties, "style"> {
    /**
     * The style of the column: as given to `addTable`, or the differential format read from a file (`null` for the parts
     * the file does not set: see {@link TableColumnStyle}).
     */
    style?: TableColumnStyle;
}

/**
 * A range of columns (`min`..`max`, 1-based) sharing one definition, as `worksheet.model` holds them (`<col>` of the
 * file). A column read from a file has no `width` (`undefined`) when the file gives none, and a `null` style for an
 * unknown style.
 */
export interface WorksheetColumnModel {
    /** The first column (1-based). */
    min: number;
    /** The last column (1-based). */
    max: number;
    /** The width, in characters. */
    width?: number;
    /** The style of the cells of the columns. */
    style?: Partial<Style> | null;
    /** Whether the columns are hidden. */
    hidden?: boolean;
    /** The outline level. */
    outlineLevel?: number;
    /** Whether the outline of the columns is collapsed. */
    collapsed?: boolean;
}

/**
 * The model of a worksheet (`worksheet.model`): its settings (shared with the worksheet, not copied), its cells (the
 * rows and column ranges as models), its images, tables and conditional formattings. Setting `worksheet.model` replaces
 * the worksheet with a model of this shape ({@link WorksheetModelInput}: without the getter-only fields), so that
 * `worksheet2.model = worksheet1.model` copies a worksheet (but its name, id and state).
 */
export interface WorksheetModel {
    /** The id of the worksheet in the workbook. */
    id: number;
    /** The name of the worksheet. */
    name: string;
    /** Whether the worksheet is visible (not read by the setter). */
    state: WorksheetState;
    /** The data validations, by cell address or range (see {@link DataValidations}). */
    dataValidations: DataValidations["model"];
    /** Tab color, default row height, outline levels, etc. */
    properties: Partial<WorksheetProperties>;
    /** The print settings. */
    pageSetup: Partial<PageSetup>;
    /** Headers and footers of the printed pages (`null` texts by default). */
    headerFooter: WorksheetHeaderFooter;
    /** The page breaks after rows. */
    rowBreaks: RowBreak[];
    /** The views (`[]` by default). */
    views: WorksheetViewModel[];
    /** The auto filter: `null` or `undefined` when there is none. */
    autoFilter?: AutoFilter | null;
    /** The images (and background image). */
    media: WorksheetMediumModel[];
    /** The protection: `null` or `undefined` when the worksheet is not protected. */
    sheetProtection?: WorksheetProtectionModel | null;
    /** The tables. */
    tables: WorksheetTableModel[];
    /** The conditional formattings. */
    conditionalFormattings: ConditionalFormattingOptions[];
    /** The column definitions, by ranges of columns; `undefined` when there are none. */
    cols?: WorksheetColumnModel[];
    /** The rows that have cells or properties, in order. */
    rows: RowModel[];
    /** The used range (the getter only; not read by the setter). */
    dimensions: Range;
    /** The merged ranges, e.g. `"A1:B2"` (the setter reads them when `mergeCells` is absent). */
    merges: string[];
    /** The merged ranges of a model read from a file by a {@link ModelContainer} (read first); absent from the getter. */
    mergeCells?: string[];
}

/**
 * What setting `worksheet.model` reads: a {@link WorksheetModel} (such as the one of another worksheet, or one of a
 * file read by a {@link ModelContainer}) whose fields that the setter does not need may be absent: `id` and `state` (the
 * worksheet keeps its own; `workbook.model =` reads them), `rowBreaks` (none when absent), `dimensions` (not read: a
 * Range from the getter, the `ref` text in a model read from a file) and `merges` (the merged cells are read from
 * `mergeCells`, else from `merges`).
 *
 * @throws Error (the setter) for a merged range that is invalid or names another sheet
 */
export interface WorksheetModelInput extends Omit<WorksheetModel, "id" | "state" | "rowBreaks" | "dimensions" | "merges"> {
    /** The id of the worksheet in the workbook (read by `workbook.model =` only). */
    id?: number;
    /** Whether the worksheet is visible (read by `workbook.model =` only: visible when absent). */
    state?: WorksheetState;
    /** The page breaks after rows (none when absent). */
    rowBreaks?: RowBreak[];
    /** Not read. */
    dimensions?: Range | string;
    /** The merged ranges, read when `mergeCells` is absent. */
    merges?: string[];
}

/**
 * What setting `workbook.model` reads, and the model of a {@link ModelContainer}: a {@link WorkbookModel} whose worksheet
 * models may lack the fields that only the worksheet getter gives (see {@link WorksheetModelInput}). A file read by a
 * ModelContainer gives worksheet models without `merges` (`mergeCells` instead), their `dimensions` as text.
 */
export interface WorkbookModelInput extends Omit<WorkbookModel, "worksheets"> {
    /** The worksheets, in their order (each one with its `id`). */
    worksheets: (WorksheetModelInput & { id: number })[];
}

export type WorksheetState = "visible" | "hidden" | "veryHidden";

/** The comparison of a `cellIs` rule. */
export type CellIsOperators =
    "equal" | "notEqual" | "greaterThan" | "greaterThanOrEqual" | "lessThan" | "lessThanOrEqual" | "between" | "notBetween";

/** The test of a `containsText` rule. */
export type ContainsTextOperators =
    | "containsText"
    | "notContainsText"
    | "beginsWith"
    | "endsWith"
    | "containsBlanks"
    | "notContainsBlanks"
    | "containsErrors"
    | "notContainsErrors";

/** The period of a `timePeriod` rule. */
export type TimePeriodTypes =
    "lastWeek" | "thisWeek" | "nextWeek" | "yesterday" | "today" | "tomorrow" | "last7Days" | "lastMonth" | "thisMonth" | "nextMonth";

/**
 * An icon set (ST_IconSetType). `3Triangles`, `3Stars`, `5Boxes` and `NoIcons` (Excel 2010) are written in the x14
 * extension of the worksheet. `NoIcons` has a single "icon": none.
 */
export type IconSetTypes =
    | "5Arrows"
    | "5ArrowsGray"
    | "5Boxes"
    | "5Quarters"
    | "5Rating"
    | "4Arrows"
    | "4ArrowsGray"
    | "4Rating"
    | "4RedToBlack"
    | "4TrafficLights"
    | "NoIcons"
    | "3Arrows"
    | "3ArrowsGray"
    | "3Flags"
    | "3Signs"
    | "3Stars"
    | "3Symbols"
    | "3Symbols2"
    | "3TrafficLights1"
    | "3TrafficLights2"
    | "3Triangles";

/** The kind of a threshold ({@link Cvfo}); `autoMin` / `autoMax` are data bar thresholds (Excel 2010). */
export type CfvoTypes = "percentile" | "percent" | "num" | "min" | "max" | "formula" | "autoMin" | "autoMax";

/** A threshold of a color scale, a data bar or an icon set (`<cfvo>`, conditional format value object). */
export interface Cvfo {
    type: CfvoTypes;
    /**
     * A number, or a formula or cell reference such as "$A$1" (type formula, or num / percent / percentile); read from a
     * file, a number when the text is a number.
     */
    value?: number | string;
}

/** A formula of a conditional formatting rule (read from a file: the text of the formula). */
export type CfFormula = string | number | boolean;

/** The fields of every conditional formatting rule. */
export interface ConditionalFormattingBaseRule {
    /**
     * Rules without a priority are given one (after the highest existing priority) when written
     */
    priority?: number;
    /**
     * The format applied to the cells (`expression`, `cellIs`, `top10`, `aboveAverage`, `containsText` and `timePeriod`
     * rules). A rule read from a file has the format of the file (see {@link DifferentialStyle}); writing the workbook
     * sets `numFmtId` on it when it has a `numFmt`.
     */
    style?: DifferentialStyle;
    /**
     * When true, lower-priority rules are not evaluated for a cell this rule applies to
     */
    stopIfTrue?: boolean;
}

/** A rule that formats the cells for which a formula is true. */
export interface ExpressionRuleType extends ConditionalFormattingBaseRule {
    type: "expression";
    /** The formula (only the first one is written), relative to the top-left cell of the range. */
    formulae?: CfFormula[];
    /**
     * Write the rule in the x14 extension of the worksheet, as Excel does for formulae that reference
     * other sheets (always done when a formula contains a sheet reference such as Other!A1).
     * Set on the rules read from that extension only.
     */
    x14?: boolean;
}

/** A rule that compares the value of the cells with one or two formulae. */
export interface CellIsRuleType extends ConditionalFormattingBaseRule {
    type: "cellIs";
    /** The values compared with (two for `between` / `notBetween`). */
    formulae?: CfFormula[];
    operator?: CellIsOperators;
    /**
     * Write the rule in the x14 extension of the worksheet, as Excel does for formulae that reference
     * other sheets (always done when a formula contains a sheet reference such as Other!A1).
     * Set on the rules read from that extension only.
     */
    x14?: boolean;
}

/** A rule that formats the top (or bottom) values of the range. */
export interface Top10RuleType extends ConditionalFormattingBaseRule {
    type: "top10";
    /** The number (or the percentage, with `percent`) of values; 10 when absent. */
    rank?: number;
    /** Whether `rank` is a percentage (`false` when absent). */
    percent?: boolean;
    /** Whether the bottom values are formatted instead of the top ones (`false` when absent). */
    bottom?: boolean;
}

/** A rule that formats the values above (or below) the average of the range. */
export interface AboveAverageRuleType extends ConditionalFormattingBaseRule {
    type: "aboveAverage";
    /** `false` formats the values below the average (`true` when absent). */
    aboveAverage?: boolean;
}

/** A rule that colors the cells on a scale of 2 or 3 colors. */
export interface ColorScaleRuleType extends ConditionalFormattingBaseRule {
    type: "colorScale";
    /** The thresholds of the scale, from the lowest (required: writing a rule without them throws). */
    cfvo: Cvfo[];
    /**
     * The color of each threshold (required: writing a rule without them throws). `undefined` is written as an empty
     * `<color/>` (no color attribute), and a `<color>` without color attribute read from a file is `undefined`.
     */
    color: (Partial<Color> | undefined)[];
}

/**
 * A custom icon of an icon set (Excel 2010): the icon of another icon set shown for a threshold, e.g.
 * `{ iconSet: "3Flags", iconId: 0 }`.
 */
export interface CfIcon {
    /** The icon set the icon comes from (`"NoIcons"`: no icon). */
    iconSet: IconSetTypes;
    /**
     * The 0-based index of the icon in `iconSet` (from 0 to its number of icons - 1), written as is. Required by the file
     * format but for `"NoIcons"` (written as 0 when absent): writing a rule whose icon has no valid `iconId` throws.
     */
    iconId?: number;
}

/** A rule that shows an icon in the cells, by threshold. */
export interface IconSetRuleType extends ConditionalFormattingBaseRule {
    type: "iconSet";
    /** Whether the value is shown with the icon (`true` when absent). */
    showValue?: boolean;
    /** Whether the order of the icons is reversed. */
    reverse?: boolean;
    /**
     * Whether the icon set shows custom `icons` (Excel 2010). A rule with `custom` or `icons` is written in the x14
     * extension of the worksheet only, with `custom="1"` when it has `icons`. A rule read from the extension has the
     * `custom` attribute of the file (`true` with its `icons`).
     */
    custom?: boolean;
    /**
     * The icon set (`"3TrafficLights1"`, the default, when absent; a rule read from a file without icon set has
     * `"3TrafficLights1"`). Writing a rule whose icon set is not one of {@link IconSetTypes} throws.
     */
    iconSet?: IconSetTypes;
    /** The thresholds, one per icon, from the lowest (required: writing a rule without them throws). */
    cfvo: Cvfo[];
    /**
     * The custom icons, one per threshold (writing a rule with another number of icons throws): written in the x14
     * extension, with `custom="1"`. Set on the rules read from it that have some.
     */
    icons?: CfIcon[];
}

/** A rule that formats the cells whose text contains a text, begins or ends with it, or that are blank or errors. */
export interface ContainsTextRuleType extends ConditionalFormattingBaseRule {
    type: "containsText";
    /**
     * Defaults to "containsText"
     */
    operator?: ContainsTextOperators;
    /**
     * The text compared by the containsText, notContainsText, beginsWith and endsWith operators
     */
    text?: string;
    /**
     * Written instead of the formula derived from the operator and text
     */
    formulae?: CfFormula[];
    /**
     * Write the rule in the x14 extension of the worksheet, as Excel does for formulae that reference
     * other sheets (always done when a formula contains a sheet reference such as Other!A1).
     * Set on the rules read from that extension only.
     */
    x14?: boolean;
}

/** A rule that formats the dates of a period relative to today. */
export interface TimePeriodRuleType extends ConditionalFormattingBaseRule {
    type: "timePeriod";
    timePeriod?: TimePeriodTypes;
    /**
     * Written instead of the formula derived from the time period
     */
    formulae?: CfFormula[];
    /**
     * Write the rule in the x14 extension of the worksheet (set on the rules read from that extension only)
     */
    x14?: boolean;
}

/** A rule that draws a bar in the cells, proportional to their value. */
export interface DataBarRuleType extends ConditionalFormattingBaseRule {
    type: "dataBar";
    /** Whether the bars are a gradient (`true` when absent). */
    gradient?: boolean;
    /** The length of the shortest bar, in percent of the cell (0 when absent). */
    minLength?: number;
    /** The length of the longest bar, in percent of the cell (100 when absent). */
    maxLength?: number;
    /** Whether the value is shown with the bar (`true` when absent). */
    showValue?: boolean;
    /** Whether the bars have a border (`false` when absent). */
    border?: boolean;
    /** Whether negative bars have the color of the positive ones (`false` when absent). */
    negativeBarColorSameAsPositive?: boolean;
    /** Whether negative bars have the border color of the positive ones (`true` when absent). */
    negativeBarBorderColorSameAsPositive?: boolean;
    /** The position of the axis between negative and positive bars (`"auto"` when absent). */
    axisPosition?: "auto" | "middle" | "none";
    direction?: "context" | "leftToRight" | "rightToLeft";
    /** The thresholds of the shortest and longest bars (`autoMin` / `autoMax` when absent). */
    cfvo?: Cvfo[];
    /**
     * Colour of the bars
     */
    color?: Partial<Color>;
    borderColor?: Partial<Color>;
    negativeFillColor?: Partial<Color>;
    negativeBorderColor?: Partial<Color>;
    axisColor?: Partial<Color>;
}

/** A conditional formatting rule, by `type`. */
export type ConditionalFormattingRule =
    | ExpressionRuleType
    | CellIsRuleType
    | Top10RuleType
    | AboveAverageRuleType
    | ColorScaleRuleType
    | IconSetRuleType
    | ContainsTextRuleType
    | TimePeriodRuleType
    | DataBarRuleType;

/**
 * The values of a row: a sparse array indexed by column number (index 0 unused), a contiguous array (index 0 set: the
 * values of columns 1, 2...), or an object keyed by column key; `null` / `undefined`: no value.
 */
export type RowValues = readonly CellValue[] | { readonly [key: string]: CellValue } | undefined | null;

/** The conditional formatting of a range (`worksheet.addConditionalFormatting`). */
export interface ConditionalFormattingOptions {
    /** The range (`"A1:E7"`; several ranges separated by spaces). */
    ref: string;
    rules: ConditionalFormattingRule[];
}

export interface Worksheet {
    readonly id: number;
    name: string;
    readonly workbook: Workbook;
    readonly hasMerges: boolean;

    /**
     * The data validations of the worksheet (`cell.dataValidation` reads and writes them cell by cell)
     */
    readonly dataValidations: DataValidations;

    readonly dimensions: Range;
    /**
     * Contains information related to how a worksheet is printed
     */
    pageSetup: Partial<PageSetup>;

    /**
     * Headers and footers of the printed pages (`null` texts by default, also for a worksheet read from a file without
     * `<headerFooter>`).
     */
    headerFooter: WorksheetHeaderFooter;

    /**
     * Worksheet State
     */
    state: WorksheetState;

    /**
     * Worksheet Properties
     */
    properties: Partial<WorksheetProperties>;

    /**
     * The views of the worksheet (frozen or split panes, zoom, grid lines...): empty by default, also for a worksheet
     * read from a file without `<sheetViews>`.
     */
    views: WorksheetViewModel[];

    /** The auto filter of the worksheet, `null` or `undefined` when there is none. */
    autoFilter?: AutoFilter | null;

    /**
     * The protection of the worksheet (set by `protect()`, cleared by `unprotect()`): `null` or `undefined` when the
     * worksheet is not protected.
     */
    sheetProtection: WorksheetProtectionModel | null | undefined;

    /** The page breaks after rows, added by `row.addPageBreak()` or read from a file. */
    rowBreaks: RowBreak[];

    /** The conditional formattings (`addConditionalFormatting`), in order. */
    conditionalFormattings: ConditionalFormattingOptions[];

    destroy(): void;

    /**
     * A count of the number of rows that have values. If a mid-document row is empty, it will not be included in the count.
     */
    readonly actualRowCount: number;

    /**
     * The total column size of the document. Equal to the maximum cell count from all of the rows
     */
    readonly columnCount: number;

    /**
     * Get the last column in a worksheet
     */
    readonly lastColumn: Column | undefined;

    /**
     * A count of the number of columns that have values.
     */
    readonly actualColumnCount: number;

    /** The column of a key (`column.key`), `undefined` for an unknown key. */
    getColumnKey(key: string): Column | undefined;

    setColumnKey(key: string, value: Column): void;

    deleteColumnKey(key: string): void;

    /** Calls `callback` with each keyed column and its key. */
    eachColumnKey(callback: (col: Column, key: string) => void): void;

    /**
     * A column by key (`column.key`), letters (`"B"`) or 1-based number, created if needed.
     *
     * Returns `undefined` for a number that is not an integer >= 1 (an invalid argument: the type does not say so).
     *
     * @throws Error for a string that is neither a column key nor valid column letters (`"A"` to `"XFD"`)
     */
    getColumn(indexOrKey: number | string): Column;

    /**
     * Removes `count` columns from column `start` (1-based; the columns to the right shift left) and inserts columns
     * there. The cells, column definitions, merges and defined names move; a merge spanning the inserted columns widens.
     *
     * @param insert - the values of each inserted column, indexed by row number - 1
     * @throws Error when a value would be inserted inside a merge, or the merges would exceed the cap
     */
    spliceColumns(start: number, count: number, ...insert: (readonly CellValue[])[]): void;

    /**
     * The columns up to the last one defined or created (all of them created by this getter); `null` when there are none.
     */
    get columns(): Column[] | null;
    /**
     * Replaces the columns: defines their headers (written in the first rows, over the existing values), keys, widths and
     * styles. Keys are a workbook-building convenience: they are not saved in the file.
     */
    set columns(value: readonly ColumnDefinition[]);

    /**
     * The total row size of the document. Equal to the row number of the last row that has values.
     */
    readonly rowCount: number;

    /**
     * Get the last editable row in a worksheet (or undefined if there are none)
     */
    readonly lastRow: Row | undefined;

    /**
     * Tries to find and return row for row no, else undefined
     *
     * @param row The 1-index row number
     */
    findRow(row: number): Row | undefined;

    /**
     * Tries to find and return rows for row no start and length, else undefined
     *
     * @param start The 1-index starting row number
     * @param length The length of the expected array
     */
    findRows(start: number, length: number): (Row | undefined)[];

    /**
     * Removes `count` rows from row `start` (1-based; the rows below shift up) and inserts rows with the `insert` values
     * there. The cells (values, styles, notes), merges and defined names move; a merge spanning the inserted rows widens.
     *
     * @throws Error when a value would be inserted inside a merge, or the merges would exceed the cap
     */
    spliceRows(start: number, count: number, ...insert: RowValues[]): void;

    /**
     * Adds a row after the last one: values by column key (`{ id: 1 }`, see `columns`), or by column (`[1, "a"]` from
     * column A; a sparse array `[, 1, "a"]` is indexed by column number).
     *
     * @param style - `"i"` / `"i+"` copies the style of the row above (`"o"` / `"o+"` are ignored here)
     */
    addRow(data: RowValues, style?: RowStyleOption): Row;

    /** Adds rows after the last one (see `addRow`). */
    addRows(rows: readonly RowValues[], style?: RowStyleOption): Row[];

    /**
     * Inserts a row at `pos` (1-based), shifting the rows below down (values as for `addRow`; see {@link RowStyleOption}).
     *
     * @throws Error when a value would be inserted inside a merge, or the merges would exceed the cap
     */
    insertRow(pos: number, value: RowValues, style?: RowStyleOption): Row;

    /**
     * Inserts rows at `pos` (1-based), shifting the rows below down (see `insertRow`).
     *
     * @returns the inserted rows; `undefined` when `values` is empty
     * @throws Error when a value would be inserted inside a merge, or the merges would exceed the cap
     */
    insertRows(pos: number, values: readonly RowValues[], style?: RowStyleOption): Row[] | undefined;

    /**
     * Duplicates row `rowNum` `count` times below it (values, styles, height and its single-row merges): the copies are
     * inserted (`insert`) or overwrite the rows below (the default).
     *
     * @throws Error when the merges would exceed the cap
     */
    duplicateRow(rowNum: number, count: number, insert?: boolean): void;

    /**
     * Get or create row by 1-based index
     *
     * @throws Error for a number that is not a row (an integer from 1 to 1,048,576)
     */
    getRow(index: number): Row;

    /**
     * The rows `start`..`start + length - 1` (1-based), created if needed.
     *
     * @returns `undefined` when `length` < 1
     * @throws Error for a number that is not a row (see `getRow`)
     */
    getRows(start: number, length: number): Row[] | undefined;

    /**
     * Iterate over all rows that have values in a worksheet
     */
    eachRow(callback: (row: Row, rowNumber: number) => void): void;

    /**
     * Iterates over the rows that have values; with `{ includeEmpty: true }`, over every row up to the last one (creating
     * the missing ones).
     */
    eachRow(options: { includeEmpty?: boolean } | null | undefined, callback: (row: Row, rowNumber: number) => void): void;

    /**
     * The values of the rows (each one is `row.values`), as a sparse array indexed by row number: `undefined` at index 0
     * and for the rows that do not exist.
     */
    getSheetValues(): (CellValue[] | undefined)[];

    /**
     * The cell at an address (`"B2"`, `"$B$2"`, `"Sheet1!B2"`) if it exists: `undefined` for an address that is not a
     * cell of this worksheet (invalid, out of bounds, or with the name of another sheet), as for a cell that does not
     * exist.
     */
    findCell(address: string): Cell | undefined;
    /**
     * The cell at (row, col) (1-based numbers) if it exists: `undefined` for numbers out of bounds, as for a cell that
     * does not exist.
     */
    findCell(row: number, col: number): Cell | undefined;

    /**
     * The cell at an address (`"B2"`, `"$B$2"`, `"Sheet1!B2"`: a sheet name must be this worksheet's, case-insensitive),
     * created if needed.
     *
     * @throws Error for an invalid address (the column letters upper case, the rows from 1 to 1,048,576, the columns from
     * `A` to `XFD`), or the address of a cell of another sheet
     */
    getCell(address: string): Cell;
    /**
     * The cell at (row, col) (1-based numbers), created if needed.
     *
     * @throws Error for a row (1 to 1,048,576) or a column (1 to 16,384) out of bounds, or a number that is not an
     * integer
     */
    getCell(row: number, col: number): Cell;

    /**
     * Merges a range into its top-left cell, which gives its style to the other cells. The range is given as for a
     * {@link Range} (see {@link RangeArguments}), but not empty: `"A4:B5"`, `"G10", "H11"`, `10, 11, 12, 13` (top,
     * left, bottom, right), a {@link Location} or a Range. A sheet name, wherever given, must be this worksheet's
     * (case-insensitive): the merge has none.
     *
     * @throws Error for an invalid range (no arguments, bounds that are not rows 1 to 1,048,576 and columns 1 to 16,384,
     * an unparsable reference) or a range of another sheet, nothing being merged; when a cell is already merged, or the
     * merges would exceed the cap
     */
    mergeCells(...cells: Exclude<RangeArguments, []>): void;
    /**
     * Merges a range into its top-left cell, the cells keeping their own style (see `mergeCells`).
     *
     * @throws Error as `mergeCells`
     */
    mergeCellsWithoutStyle(...cells: Exclude<RangeArguments, []>): void;

    /**
     * Unmerges the merges that have a cell in the range (given as for `mergeCells`): a merge is unmerged as a whole.
     * Unmerging breaks the style links.
     *
     * @throws Error for an invalid range, or a range of another sheet (nothing is unmerged)
     */
    unMergeCells(...cells: Exclude<RangeArguments, []>): void;

    /**
     * Sets a formula on a range (`"A1:B3"`, or a single cell `"C3"`; a sheet name must be this worksheet's): the
     * top-left cell gets it (a shared formula, or an array formula with `shareType` `"array"`) with the range as `ref`
     * (`"A1:B3"`, `"C3"`: without sheet name nor `$`), the other cells are its clones (shared formula) or get their
     * result (array formula).
     *
     * @param results - the cached results of the cells: by cell (1-based row and column), by row then column of the
     *   range, or flat (row by row)
     * @param shareType - `"shared"` (the default) or `"array"`
     * @throws Error for an invalid range (as `mergeCells`) or a range of another sheet: nothing is filled
     */
    fillFormula(
        range: string,
        formula: string,
        results?:
            | ((row: number, col: number) => CellFormulaResult | undefined)
            | readonly (CellFormulaResult | undefined)[]
            | readonly (readonly (CellFormulaResult | undefined)[])[],
        shareType?: "shared" | "array"
    ): void;

    /**
     * Using the image id from `Workbook.addImage`, set the background to the worksheet
     */
    addBackgroundImage(imageId: number): void;

    getBackgroundImageId(): number | string | undefined;

    /**
     * Using the image id from `Workbook.addImage`,
     * embed an image within the worksheet to cover a range
     */
    addImage(
        imageId: number,
        range:
            | string
            | ({ editAs?: ImageEditAs } & ImageRange & { hyperlinks?: ImageHyperlinkValue })
            | ({ editAs?: ImageEditAs } & ImagePosition & { hyperlinks?: ImageHyperlinkValue })
    ): void;

    /** The images of the worksheet (not the background image). */
    getImages(): WorksheetImage[];

    /** The model of the worksheet (see {@link WorksheetModel}). */
    get model(): WorksheetModel;
    /**
     * Replaces the worksheet with a model (see {@link WorksheetModelInput}: the getter-only fields may be absent).
     * @throws Error for an invalid name, for overlapping ranges in `mergeCells`, or when the merged cells exceed the caps
     * (1,048,576 per worksheet, `maxMergedCells` of the workbook)
     */
    set model(value: WorksheetModelInput);

    /**
     * Worksheet protection
     */
    protect(password?: string, options?: Partial<WorksheetProtection>): Promise<void>;
    unprotect(): void;

    /**
     * Add a new table and return a reference to it
     */
    addTable(tableProperties: TableProperties): Table;
    /**
     * fetch table by name or id
     */
    getTable(name: string): Table | undefined;
    /**
     * delete table by name or id
     */
    removeTable(name: string): void;
    /**
     *  fetch table
     */
    getTables(): Table[];
    /**
     * add conditionalFormattingOptions
     */
    addConditionalFormatting(cf: ConditionalFormattingOptions): void;

    /**
     * Removes conditional formattings: the one at an index (of `conditionalFormattings`), or those for which a filter
     * returns a falsy value (it keeps the others); with no argument (or any other value), all of them.
     */
    removeConditionalFormatting(
        filter?: number | ((cf: ConditionalFormattingOptions, index: number, all: ConditionalFormattingOptions[]) => unknown) | null
    ): void;
}

/**
 * The calculation properties of a workbook (`workbook.calcProperties`, `<calcPr>`): written, never read from a file (a
 * loaded workbook has `{}`).
 */
export interface CalculationProperties {
    /**
     * Whether the application shall perform a full recalculation when the workbook is opened
     */
    fullCalcOnLoad: boolean;
}

export interface WorksheetProperties {
    /**
     * Color of the tab (a color read from a file may also have a `tint`, or be an `indexed` palette color)
     */
    tabColor: Partial<Color>;

    /**
     * The worksheet column outline level (default: 0)
     */
    outlineLevelCol: number;

    /**
     * The worksheet row outline level (default: 0)
     */
    outlineLevelRow: number;

    /**
     * Where the summary rows and columns of the outlines are (below / right of the details when absent). Written when one
     * of them is set; read from a file, an attribute that the file does not give is `undefined`.
     */
    outlineProperties: {
        /** Whether the summary rows are below their details. */
        summaryBelow?: boolean;
        /** Whether the summary columns are right of their details. */
        summaryRight?: boolean;
    };
    /**
     * Default row height (default: 15)
     */
    defaultRowHeight: number;

    /**
     * Default column width (optional)
     */
    defaultColWidth?: number;

    /**
     * default: 55
     */
    dyDescent: number;
}

/** The options of `workbook.addWorksheet`: the initial settings of the worksheet. */
export interface AddWorksheetOptions {
    /** Tab color, default row height, outline levels, etc. */
    properties: Partial<WorksheetProperties>;
    /** The print settings (merged with the defaults). */
    pageSetup: Partial<PageSetup>;
    /** Headers and footers of the printed pages (merged with the defaults). */
    headerFooter: WorksheetHeaderFooter;
    /** The views (frozen or split panes, zoom, grid lines...). */
    views: WorksheetViewModel[];
    /** Whether the worksheet is visible (`"visible"` by default). */
    state: WorksheetState;
    /** The auto filter, e.g. `"A1:C1"`. */
    autoFilter: AutoFilter | null;
}

/**
 * The properties of a workbook (`workbook.properties`, `<workbookPr>`): empty for a new workbook; a loaded workbook has
 * `date1904` when the file has a `<workbookPr>`.
 */
export interface WorkbookProperties {
    /**
     * Set workbook dates to 1904 date system
     */
    date1904: boolean;
}

/**
 * Zip options of the writers (`xlsx.write`, `writeBuffer`, `writeFile` and the streaming WorkbookWriter)
 */
export interface ZipOptions {
    /**
     * DEFLATE level from 1 (fastest) to 9 (smallest); 0 stores the parts uncompressed
     * @default 6
     */
    level: number;
    /**
     * Write ZIP64 records for every entry. They are added automatically when a size or an offset
     * exceeds 4 GiB or the archive has 65535 entries or more.
     * @default false
     */
    forceZip64: boolean;
    /**
     * Archive comment
     */
    comment: string;
    /**
     * @deprecated JSZip option (ExcelJS < 5): 'STORE' is `level: 0`
     */
    compression: "STORE" | "DEFLATE";
    /**
     * @deprecated JSZip option (ExcelJS < 5): use `level`
     */
    compressionOptions: null | { level: number };
    /**
     * @deprecated archiver option (ExcelJS < 5): `store: true` is `level: 0`
     */
    store: boolean;
    /**
     * @deprecated archiver option (ExcelJS < 5): only `level` is used
     */
    zlib: { level?: number };
}

/**
 * @deprecated use ZipOptions
 */
export type JSZipGeneratorOptions = ZipOptions;

/**
 * Decompression limits of the readers. A file is rejected as soon as one of them is exceeded
 * while it is inflated.
 */
export interface ZipReadLimits {
    /**
     * Maximum decompressed size of one zip entry (part), in bytes.
     * @default 1 GiB (xlsx.load / read / readFile), 16 GiB (streaming WorkbookReader)
     */
    maxEntrySize: number;
    /**
     * Maximum decompressed size of all the entries read, in bytes.
     * @default 2 GiB (xlsx.load / read / readFile), 32 GiB (streaming WorkbookReader)
     */
    maxUncompressedSize: number;
    /**
     * Maximum compression ratio of an entry larger than 10 MiB.
     * @default 1000
     */
    maxCompressionRatio: number;
}

/** The options of `xlsx.load`, `read` and `readFile`. */
export interface XlsxReadOptions extends ZipReadLimits {
    /**
     * The names of the worksheet XML elements to skip while parsing (e.g. `["dataValidations"]`)
     */
    ignoreNodes: string[];
    /**
     * Maximum number of cells covered by merges in all the worksheets of the file (every merged cell is created in
     * memory). One worksheet holds at most 1,048,576 merged cells whatever this option. The loaded workbook keeps this
     * value as its `maxMergedCells`.
     * @default workbook.maxMergedCells (1048576, one full worksheet)
     */
    maxMergedCells: number;
}

/** The options of `xlsx.load`: {@link XlsxReadOptions}, and `base64` for data given as base64 text. */
export interface XlsxLoadOptions extends XlsxReadOptions {
    /**
     * The data is the base64 text of the file: a string, or a Node `Buffer` holding that text (a plain `Uint8Array` is
     * not decoded as text).
     */
    base64: boolean;
}

/** The options of `xlsx.write`, `writeBuffer` and `writeFile`. */
export interface XlsxWriteOptions {
    /**
     * Whether to write the strings of the cells in a shared strings table (else inline in the cells)
     * @default true
     */
    useSharedStrings: boolean;
    /**
     * Whether to write the styles of the cells (else every cell has the default style)
     * @default true
     */
    useStyles: boolean;
    /**
     * Zip options
     */
    zip: Partial<ZipOptions>;
}

/**
 * An output of `xlsx.write`: a Node.js writable stream, or any object with `write(chunk)` and `end()` (`on` /
 * `removeListener`, when present, are used to wait for `"drain"` and to report errors).
 */
export interface XlsxOutput {
    /**
     * Writes a chunk of the file (a Node Buffer, or the Buffer polyfill in the browser bundle); returning `false` (with
     * `on`) waits for `"drain"`.
     */
    write(chunk: globalThis.Buffer): unknown;
    /** Ends the output (`xlsx.write` does not call it). */
    end(): unknown;
}

/**
 * The xlsx reader and writer of a workbook (`workbook.xlsx`) or of a {@link ModelContainer} (`container.xlsx`).
 * @typeParam T - the object read into and written from, which the read methods resolve with
 */
export interface Xlsx<T = Workbook> {
    /**
     * Reads a file into the workbook (Node.js and Bun only).
     * @returns the workbook, once read
     * @throws (rejects) when the file does not exist, or as {@link Xlsx.load}
     */
    readFile(path: string, options?: Partial<XlsxReadOptions>): Promise<T>;

    /**
     * Reads the bytes of a file from a stream (anything with `pipe`), or from an async iterable of byte chunks (e.g. a
     * web `ReadableStream`), into the workbook.
     * @returns the workbook, once read
     * @throws (rejects) on an error of the stream, or as {@link Xlsx.load}
     */
    read(
        stream: import("stream").Stream | NodeJS.ReadableStream | AsyncIterable<Uint8Array>,
        options?: Partial<XlsxReadOptions>
    ): Promise<T>;

    /**
     * Loads a file into the workbook from its bytes (a Buffer, a Uint8Array or any ArrayBufferView, an ArrayBuffer, an
     * array of byte values) or a Blob (e.g. the `File` of an `<input type="file">`).
     * @returns the workbook, once read
     * @throws (rejects) on an invalid zip or XML, a decompression limit or a limit of the worksheets exceeded
     */
    load(data: ArrayBuffer | ArrayBufferView | Blob | number[], options?: Partial<XlsxReadOptions>): Promise<T>;
    /**
     * Loads a file into the workbook from its base64 text (`base64: true`): a string, or a Node `Buffer` of that text.
     * @returns the workbook, once read
     * @throws (rejects) as the other form
     */
    load(data: string | globalThis.Buffer, options: Partial<XlsxLoadOptions> & { base64: true }): Promise<T>;

    /**
     * Writes the workbook as an xlsx file in memory.
     * @returns the bytes of the file
     * @throws (rejects) as {@link Xlsx.write}
     */
    writeBuffer(options?: Partial<XlsxWriteOptions>): Promise<Buffer>;

    /**
     * Writes the workbook to a file (Node.js and Bun only); resolves once the file is closed.
     * @throws (rejects) as {@link Xlsx.write}, or on an error of the file
     */
    writeFile(path: string, options?: Partial<XlsxWriteOptions>): Promise<void>;

    /**
     * Writes the workbook as an xlsx file to an output; resolves with this object once the archive is written (the
     * output is not ended: call `end()` afterwards).
     * @throws (rejects) for an image extension that is not 1 to 5 letters or digits (before anything is written), an
     * image without data, or a failure of the output
     */
    write(stream: XlsxOutput | NodeJS.WritableStream, options?: Partial<XlsxWriteOptions>): Promise<this>;
}

// https://c2fo.github.io/fast-csv/docs/parsing/options

/** The header names of the CSV parser: `undefined` (or `null`) skips the column. */
type HeaderArray = (string | undefined | null)[];
/** Renames the header names read from the first row (`parserOptions.headers`). */
type HeaderTransformFunction = (headers: HeaderArray) => HeaderArray;
/** The options of the fast-csv parser (`csv.read` / `readFile` `parserOptions`). */
export interface FastCsvParserOptionsArgs {
    objectMode: boolean;
    delimiter: string;
    quote: string | null;
    escape: string;
    /**
     * `true`: the first row holds the header names; an array: the header names; a function: renames the header names of
     * the first row. The names become the first row of the worksheet and the column keys (made unique).
     */
    headers: boolean | HeaderTransformFunction | HeaderArray;
    renameHeaders: boolean;
    ignoreEmpty: boolean;
    comment: string;
    strictColumnHandling: boolean;
    discardUnmappedColumns: boolean;
    trim: boolean;
    ltrim: boolean;
    rtrim: boolean;
    encoding: string;
    maxRows: number;
    skipLines: number;
    skipRows: number;
}

/** Columns to quote, by header name. */
interface QuoteColumnMap {
    [s: string]: boolean;
}
/** Whether to quote all the columns, or the ones of an array (by index) or of a map (by header name). */
declare type QuoteColumns = boolean | boolean[] | QuoteColumnMap;

/** A row of the CSV formatter as an object: its values by header name. */
interface RowMap {
    [key: string]: unknown;
}
/** A row of the CSV formatter as `[header, value]` pairs. */
declare type RowHashArray = [string, unknown][];
/** A row of the CSV formatter as its values: what `csv.write` gives (the results of `map`). */
declare type RowArray = unknown[];
/** A row of the CSV formatter. */
declare type Rows = RowArray | RowMap | RowHashArray;
/** The callback of an asynchronous row transform: the transformed row, or an error. */
declare type RowTransformCallback = (error?: Error | null, row?: Rows) => void;
/**
 * Transforms each row before it is formatted: synchronously (one parameter: returns the row) or asynchronously (two
 * parameters: calls back with the row). `csv.write` gives each row as the array of its mapped values.
 */
declare type RowTransformFunction = ((row: RowArray) => Rows) | ((row: RowArray, callback: RowTransformCallback) => void);

// https://c2fo.github.io/fast-csv/docs/formatting/options/
/** The options of the fast-csv formatter (`csv.write` / `writeFile` / `writeBuffer` `formatterOptions`). */
export interface FastCsvFormatterOptionsArgs {
    objectMode: boolean;
    delimiter: string;
    rowDelimiter: string;
    quote: string | boolean;
    escape: string;
    quoteColumns: QuoteColumns;
    quoteHeaders: QuoteColumns;
    headers: null | boolean | string[];
    writeHeaders: boolean;
    includeEndRowDelimiter: boolean;
    writeBOM: boolean;
    transform: RowTransformFunction;
    alwaysWriteHeaders: boolean;
}

/** The options of `csv.read` and `readFile`. */
export interface CsvReadOptions {
    /**
     * The formats (dayjs, strict parsing) of the dates of the default `map`
     * @default ["YYYY-MM-DD[T]HH:mm:ssZ", "YYYY-MM-DD[T]HH:mm:ss", "MM-DD-YYYY", "YYYY-MM-DD"]
     */
    dateFormats: string[];
    /**
     * Maps each field to the value of its cell. By default an empty field is null, a number a number (an integer that
     * would lose digits stays text), a date in one of `dateFormats` a Date, `true` / `false` a boolean, `#N/A`... an
     * error value, anything else text.
     * @param value - the text of the field
     * @param index - the 0-based index of the column
     */
    map(value: string, index: number): CellValue;
    /** The name of the new worksheet (default `sheet<id>`). */
    sheetName: string;
    /** The options of the parser. */
    parserOptions: Partial<FastCsvParserOptionsArgs>;
}

/** The options of `csv.write`, `writeFile` and `writeBuffer`. */
export interface CsvWriteOptions {
    /** The dayjs format of the dates of the default `map` (default ISO 8601). */
    dateFormat: string;
    /** Whether the default `map` writes the dates in UTC (else in local time). */
    dateUTC: boolean;
    /** The worksheet to write, by name (else `sheetId`, else the first worksheet). */
    sheetName: string;
    /** The worksheet to write, by id. */
    sheetId: number;
    /**
     * The encoding of the file (`writeFile` only)
     * @default "utf8"
     */
    encoding: string;
    /**
     * Maps each cell value to the field written. By default a hyperlink is its link, a formula its result, a date ISO
     * 8601 (or `dateFormat`), an error its text, another object its JSON.
     * @param value - the value of the cell
     * @param index - the 0-based index of the column
     */
    map(value: CellValue, index: number): unknown;
    /**
     * Whether to write an empty line for each row without values between two rows
     * @default true
     */
    includeEmptyRows: boolean;
    /** The options of the formatter. */
    formatterOptions: Partial<FastCsvFormatterOptionsArgs>;
}

/** The CSV reader and writer of a workbook (`workbook.csv`): one worksheet per file. */
export interface Csv {
    /**
     * Reads a CSV file into a new worksheet (Node.js and Bun only).
     * @returns the new worksheet
     * @throws (rejects) when the file does not exist, or as {@link Csv.read}
     */
    readFile(path: string, options?: Partial<CsvReadOptions>): Promise<Worksheet>;

    /**
     * Reads CSV text from a readable stream into a new worksheet.
     * @returns the new worksheet
     * @throws (rejects) on an error of the stream or of the parser, or a header of more than 16,384 names
     */
    read(stream: NodeJS.ReadableStream, options?: Partial<CsvReadOptions>): Promise<Worksheet>;

    /**
     * Writes a worksheet as CSV in memory.
     * @returns the CSV bytes
     */
    writeBuffer(options?: Partial<CsvWriteOptions>): Promise<Buffer>;

    /**
     * Writes a worksheet to a CSV file (Node.js and Bun only); resolves once the file is written.
     * @throws (rejects) on an error of the file or of the formatter
     */
    writeFile(path: string, options?: Partial<CsvWriteOptions>): Promise<void>;

    /**
     * Writes a worksheet as CSV to a writable stream, which is ended; resolves on its `"finish"` event.
     * @throws (rejects) on an error of the stream or of the formatter
     */
    write(stream: NodeJS.WritableStream, options?: Partial<CsvWriteOptions>): Promise<void>;
}

/**
 * An image of a workbook (`workbook.getImage(id)`, `workbook.model.media`): an {@link Image} as added, or as read from a
 * file (`name`, `extension` and `buffer`).
 */
export interface Media {
    /** Always `"image"`. */
    type: "image";
    /** The name of the image part (`image1`): set when the workbook is read or written. */
    name?: string;
    /** The extension of the image file (`png`, `jpeg`...; as named in the file for a loaded workbook). */
    extension: string;
    /** The bytes of the image; `null` for an empty image part of a loaded file. */
    buffer?: ArrayBuffer | Uint8Array | null;
    /** The image as base64 (see {@link Image.base64}). */
    base64?: string;
    /** The path of the image file (see {@link Image.filename}). */
    filename?: string;
    /** The index of the image part, set when a workbook is read. */
    index?: number;
}

export interface Address {
    sheetName?: string;
    address: string;
    col: number;
    row: number;
    $col$row: string;
}

/**
 * The arguments that define a range: nothing (A1:A1), a range, a location or a reference (`"Sheet1!A1:B2"`), two
 * addresses and a sheet name, or top, left, bottom, right (1-based) and a sheet name; or such a list as one array.
 */
export type RangeArguments =
    | []
    | [range: Range | Location | string | RangeArguments]
    | [tl: string, br: string, sheetName?: string]
    | [top: number, left: number, bottom: number, right: number, sheetName?: string];

/** The bounds of a rectangle of cells (1-based row and column numbers). */
export interface Location {
    top: number;
    left: number;
    bottom: number;
    right: number;
}

/** A defined name and its ranges (`definedNames.getRanges(name)`, an entry of `definedNames.model`). */
export interface DefinedNamesRanges {
    /** The name. */
    name: string;
    /** Its cells and ranges, as `$`-absolute references (`"Sheet1!$A$1:$B$2"`). */
    ranges: string[];
    /** The 0-based index of the worksheet a name is local to (absent: a workbook-level name). */
    localSheetId?: number;
}

/** The defined names of a workbook (`definedNames.model`); setting it loads only the `$`-absolute references. */
export type DefinedNamesModel = DefinedNamesRanges[];

/** A cell of a defined name: its worksheet and address (`row` / `col` 1-based, `address` without `$`). */
export interface DefinedNameCell {
    sheetName?: string;
    address: string;
    row: number;
    col: number;
    $col$row?: string;
}

/**
 * The defined names of a workbook (`workbook.definedNames`). The cells of a name follow the rows and columns inserted or
 * removed; a range of more than 100,000 cells is kept as a rectangle (removing cells from it cuts it into the rectangles
 * that remain).
 */
export interface DefinedNames {
    /**
     * Adds a name to a cell or a range, given as `"Sheet1!$A$1"` or `"Sheet1!$A$1:$B$2"`.
     * @throws Error when a column or a row is out of bounds
     */
    add(locStr: string, name: string): void;
    /**
     * Adds a name to a cell, or to the cells of a range given by its 1-based bounds (`top`, `left`, `bottom`, `right`,
     * e.g. a decoded range reference).
     */
    addEx(location: DefinedNameCell | (Location & { sheetName?: string }), name: string): void;

    /**
     * Removes a name from a cell or from the cells of a range, given as a reference (`"Sheet1!$A$1"` or
     * `"Sheet1!$A$1:$B$2"`).
     * @throws Error when a column or a row is out of bounds
     */
    remove(locStr: string, name: string): void;
    /**
     * Removes a name from a cell, or from the cells of a range given by its 1-based bounds (`top`, `left`, `bottom`,
     * `right`, e.g. a decoded range reference).
     * @throws Error for a range whose bounds are not rows and columns of the sheet, or are reversed
     */
    removeEx(location: DefinedNameCell | (Location & { sheetName?: string }), name: string): void;
    /**
     * Removes every name of a cell, or of the cells of a range given by its 1-based bounds.
     * @throws Error for a range whose bounds are not rows and columns of the sheet, or are reversed
     */
    removeAllNames(location: DefinedNameCell | (Location & { sheetName?: string })): void;

    /** Calls `callback` with each name and each of its cells (the cells of the very large ranges are not enumerated). */
    forEach(callback: (name: string, cell: DefinedNameCell) => void): void;

    /** The names of a cell, given as a reference (`"Sheet1!$A$1"`). */
    getNames(addressStr: string): string[];
    /** The names of a cell. */
    getNamesEx(address: DefinedNameCell): string[];

    /** The ranges of a name, as references (`"Sheet1!$A$1:$B$2"`); none for an unknown name. */
    getRanges(name: string): DefinedNamesRanges;

    /** The names with their ranges (names without range are left out); setting it replaces every name. */
    model: DefinedNamesModel;
}

/**
 * A `<sheet>` of the workbook part: the `sheets` of the model that `xlsx` reads into a {@link ModelContainer}.
 */
export interface WorkbookSheetModel {
    /** The sheet id (`NaN` when the file's is not a number). */
    id: number;
    /** The worksheet name. */
    name: string;
    /** `"visible"` (absent), `"hidden"` or `"veryHidden"`. */
    state?: string;
    /** The relationship id of the worksheet part. */
    rId?: string;
}

/**
 * The model of a workbook (`workbook.model`, `ModelContainer.model`): the document properties, the worksheet models and
 * the settings. The document properties are absent when unset (the getter gives `"Unknown"` for an unset `creator` /
 * `lastModifiedBy`); setting `workbook.model` reads `sheets` only to order the worksheets.
 */
export interface WorkbookModel {
    /** Document property: the author. */
    creator?: string;
    /** Document property: the author of the last change. */
    lastModifiedBy?: string;
    /** Document property: the date of the last print. */
    lastPrinted?: Date;
    /** Document property: the creation date. */
    created?: Date;
    /** Document property: the date of the last change. */
    modified?: Date;
    /** The workbook properties. */
    properties: Partial<WorkbookProperties>;
    /** The calculation properties. */
    calcProperties: Partial<CalculationProperties>;
    /** The worksheets, in their order. */
    worksheets: WorksheetModel[];
    /** The worksheet models again (getter), or the `<sheet>` elements read by a {@link ModelContainer}. */
    sheets?: WorksheetModel[] | WorkbookSheetModel[];
    /** The defined names. */
    definedNames: DefinedNamesModel;
    /** The workbook windows. */
    views: WorkbookView[];
    /** Document property: the company. */
    company?: string;
    /** Document property: the manager. */
    manager?: string;
    /** Document property: the title. */
    title?: string;
    /** Document property: the subject. */
    subject?: string;
    /** Document property: the keywords. */
    keywords?: string;
    /** Document property: the category. */
    category?: string;
    /** Document property: the description. */
    description?: string;
    /** Document property: the identifier (read and written by a {@link ModelContainer} only). */
    identifier?: string;
    /** Document property: the language. */
    language?: string;
    /** Document property: the revision number (`NaN` when the file's is not a number). */
    revision?: number;
    /** Document property: the content status (`"Draft"`, `"Final"`...). */
    contentStatus?: string;
    /** Document property: the content type (read and written by a {@link ModelContainer} only). */
    contentType?: string;
    /** The themes read from a file: part name (`theme1`) → XML, written back as they are. */
    themes?: Record<string, string>;
    /** The images, by image id. */
    media: Media[];
}

/**
 * A workbook: its worksheets, defined names, images, properties and document properties; `xlsx` and `csv` read and
 * write it.
 */
export class Workbook {
    /** Document property: the category (`""` for a new workbook; `undefined` after loading a file without it). */
    category: string | undefined;
    /** Document property: the company (`""` for a new workbook; `undefined` after loading a file without it). */
    company: string | undefined;
    /** Document property: the author (written as `"Unknown"` when unset). */
    creator?: string;
    /** Document property: the description (`""` for a new workbook; `undefined` after loading a file without it). */
    description: string | undefined;
    /** Document property: the keywords (`""` for a new workbook; `undefined` after loading a file without it). */
    keywords: string | undefined;
    /** Document property: the author of the last change (written as `"Unknown"` when unset). */
    lastModifiedBy?: string;
    /** Document property: the creation date (now for a new workbook; `undefined` after loading a file without it). */
    created: Date | undefined;
    /** Document property: the manager (`""` for a new workbook; `undefined` after loading a file without it). */
    manager: string | undefined;
    /** Document property: the date of the last change (now for a new workbook; `undefined` after loading a file without it). */
    modified: Date | undefined;
    /** Document property: the date of the last print. */
    lastPrinted?: Date;
    /** Document property: the subject (`""` for a new workbook; `undefined` after loading a file without it). */
    subject: string | undefined;
    /** Document property: the title (`""` for a new workbook; `undefined` after loading a file without it). */
    title: string | undefined;
    /** Document property: the language. */
    language?: string;
    /** Document property: the revision number (`NaN` after loading a file whose revision is not a number). */
    revision?: number;
    /** Document property: the content status (`"Draft"`, `"Final"`...). */
    contentStatus?: string;
    /** The workbook properties (`{}` for a new workbook). */
    properties: Partial<WorkbookProperties>;

    /**
     * Workbook calculation Properties (`{}` for a new or a loaded workbook)
     */
    calcProperties: Partial<CalculationProperties>;

    /**
     * xlsx file format operations
     */
    readonly xlsx: Xlsx;

    /**
     * csv file format operations
     */
    readonly csv: Csv;

    /** The id that the next worksheet added gets: the first free id (from 1). */
    readonly nextId: number;

    /** The defined names. */
    readonly definedNames: DefinedNames;

    /**
     * Maximum number of cells covered by merges in all the worksheets (`mergeCells` throws beyond it; one worksheet holds
     * at most 1,048,576). Also the default `maxMergedCells` of `xlsx.load` / `read` / `readFile`.
     * @default 1048576 (one full worksheet)
     */
    maxMergedCells: number;

    /** The model of the workbook (the worksheet models and the settings). */
    get model(): WorkbookModel;
    /**
     * Replaces the workbook with a model (see {@link WorkbookModelInput}: such as the model of another workbook, or of a
     * {@link ModelContainer} that read a file).
     * @throws Error for an invalid worksheet name, or when the merges exceed `maxMergedCells`
     */
    set model(value: WorkbookModelInput);

    /**
     * The Workbook views controls how many separate windows Excel will open when viewing the workbook. `[]` for a new
     * workbook, also after loading a file without `<bookViews>` (such as one written with no views).
     */
    views: WorkbookView[];

    /** The images (`addImage`, or read from a file), by image id. */
    media: Media[];

    /**
     * The worksheets, in their order (a new array)
     */
    readonly worksheets: Worksheet[];

    /**
     * Add a new worksheet and return a reference to it
     * @param name - the name (default `sheet<id>`)
     * @throws Error for an invalid or duplicate name
     */
    addWorksheet(name?: string, options?: Partial<AddWorksheetOptions>): Worksheet;

    /** Removes a worksheet (a worksheet removed twice does not remove the worksheet that took its id). */
    removeWorksheetEx(worksheet: Worksheet): void;
    /** Removes a worksheet by id or name (nothing for an unknown one). */
    removeWorksheet(indexOrName: number | string): void;

    /**
     * fetch sheet by name or id; without argument, the first worksheet (by id)
     */
    getWorksheet(indexOrName?: number | string): Worksheet | undefined;

    /**
     * Iterate over all sheets.
     *
     * Note: `workbook.worksheets.forEach` will still work but this is better.
     */
    eachSheet(callback: (worksheet: Worksheet, id: number) => void): void;

    /** Drops the themes read from a file (the default theme is then written). */
    clearThemes(): void;

    /**
     * Add Image to Workbook and return the id (a copy of `img`, checked only when the workbook is written)
     */
    addImage(img: Image): number;

    /** An image by id: as added (with `type`), or as read from a file; `undefined` for an unknown id. */
    getImage(id: number): Media | undefined;
}

/**
 * A workbook model without a Workbook (Node.js and Bun only: not in the browser bundles): `xlsx` reads a file into
 * `model` without creating worksheets, and writes `model`.
 */
export class ModelContainer {
    /** @param model - the model to write (omitted to read a file into the container) */
    constructor(model?: WorkbookModel);
    /**
     * The workbook model: what `xlsx` writes, and what it sets when reading (then the parsed model, whose worksheets
     * hold their parsed rows and cells); `undefined` before a file is read into a container created without one.
     */
    model: WorkbookModelInput | undefined;
    /** The xlsx reader and writer of the model (its read methods resolve with the container). */
    readonly xlsx: Xlsx<ModelContainer>;
}

/** The style of a table (`TableProperties.style`). */
export interface TableStyleProperties {
    /**
     * The table style: a built-in one (listed), or the name of a custom table style of the workbook (a table read from a
     * file keeps it; the custom styles themselves are not written back). `null` for no table style.
     * @default 'TableStyleMedium2'
     */
    theme?:
        | null
        | (string & Record<never, never>)
        | "TableStyleDark1"
        | "TableStyleDark10"
        | "TableStyleDark11"
        | "TableStyleDark2"
        | "TableStyleDark3"
        | "TableStyleDark4"
        | "TableStyleDark5"
        | "TableStyleDark6"
        | "TableStyleDark7"
        | "TableStyleDark8"
        | "TableStyleDark9"
        | "TableStyleLight1"
        | "TableStyleLight10"
        | "TableStyleLight11"
        | "TableStyleLight12"
        | "TableStyleLight13"
        | "TableStyleLight14"
        | "TableStyleLight15"
        | "TableStyleLight16"
        | "TableStyleLight17"
        | "TableStyleLight18"
        | "TableStyleLight19"
        | "TableStyleLight2"
        | "TableStyleLight20"
        | "TableStyleLight21"
        | "TableStyleLight3"
        | "TableStyleLight4"
        | "TableStyleLight5"
        | "TableStyleLight6"
        | "TableStyleLight7"
        | "TableStyleLight8"
        | "TableStyleLight9"
        | "TableStyleMedium1"
        | "TableStyleMedium10"
        | "TableStyleMedium11"
        | "TableStyleMedium12"
        | "TableStyleMedium13"
        | "TableStyleMedium14"
        | "TableStyleMedium15"
        | "TableStyleMedium16"
        | "TableStyleMedium17"
        | "TableStyleMedium18"
        | "TableStyleMedium19"
        | "TableStyleMedium2"
        | "TableStyleMedium20"
        | "TableStyleMedium21"
        | "TableStyleMedium22"
        | "TableStyleMedium23"
        | "TableStyleMedium24"
        | "TableStyleMedium25"
        | "TableStyleMedium26"
        | "TableStyleMedium27"
        | "TableStyleMedium28"
        | "TableStyleMedium3"
        | "TableStyleMedium4"
        | "TableStyleMedium5"
        | "TableStyleMedium6"
        | "TableStyleMedium7"
        | "TableStyleMedium8"
        | "TableStyleMedium9";
    /**
     * Highlight the first column (bold)
     * @default false
     */
    showFirstColumn?: boolean;
    /**
     * Highlight the last column (bold)
     * @default false
     */
    showLastColumn?: boolean;
    /**
     * Alternate rows shown with background colour
     * @default false
     */
    showRowStripes?: boolean;
    /**
     * Alternate rows shown with background colour
     * @default false
     */
    showColumnStripes?: boolean;
}

/** A column of a new table (`TableProperties.columns`, `table.addColumn`). */
export interface TableColumnProperties {
    /**
     * The name of the column, also used in the header
     */
    name: string;
    /**
     * Switches the filter control in the header
     * @default false
     */
    filterButton?: boolean;
    /**
     * Label to describe the totals row (first column)
     * @default 'Total'
     */
    totalsRowLabel?: string;
    /**
     * Name of the totals function
     * @default 'none'
     */
    totalsRowFunction?: "none" | "average" | "countNums" | "count" | "max" | "min" | "stdDev" | "var" | "sum" | "custom";
    /**
     * Optional formula for custom functions
     */
    totalsRowFormula?: string;
    /**
     * The cached result of the totals row formula (`totalsRowFormula`), written as the result of the formula cell
     */
    totalsRowResult?: CellFormulaResult;

    /**
     * Styles applied to the column
     */
    style?: Partial<Style>;
}

/** The definition of a new table (`worksheet.addTable`). */
export interface TableProperties {
    /**
     * The name of the table
     */
    name: string;
    /**
     * The display name of the table
     */
    displayName?: string;
    /**
     * Top left cell of the table
     */
    ref: string;
    /**
     * Show headers at top of table
     * @default true
     */
    headerRow?: boolean;
    /**
     * Show totals at bottom of table
     * @default false
     */
    totalsRow?: boolean;
    /**
     * Extra style properties
     * @default {}
     */
    style?: TableStyleProperties;
    /**
     * Column definitions
     */
    columns: TableColumnProperties[];
    /**
     * Rows of data: the values of the cells of each row, by column
     */
    rows: CellValue[][];
}

/**
 * The style of a table column as the table holds it: the `style` of its {@link TableColumnProperties}, or the
 * differential format read from a file (`null` for the parts the file does not set).
 */
export type TableColumnStyle = DifferentialStyle;

/** A column of a table (`table.getColumn(i)`): changes to `name` and to the totals row need `table.commit()`. */
export interface TableColumn {
    name: string;
    filterButton: boolean | undefined;
    style: TableColumnStyle | undefined;
    totalsRowLabel: string | undefined;
    totalsRowFunction: TableColumnProperties["totalsRowFunction"];
    /** The cached result of the totals row formula. */
    totalsRowResult: CellFormulaResult | undefined;
    totalsRowFormula: string | undefined;
}

/**
 * A table of a worksheet (`worksheet.addTable`, `worksheet.getTable`). Its rows and columns live in the cells of the
 * worksheet: change them with the methods below, then `commit()`. A table read from a file has no data rows of its own
 * (they are the cells of its worksheet): the methods that change its rows, columns or position throw an Error.
 */
export interface Table {
    /** The table name (unique in the workbook, used in structured references). */
    name: string;
    /** The display name (the name when there is none). */
    get displayName(): string;
    set displayName(value: string | undefined);
    /** The top-left cell of the table (`"A1"`); changing it moves the table on `commit()`. */
    ref: string | undefined;
    /** Whether the table has a header row; needs `commit()` once changed. */
    headerRow: boolean | undefined;
    /** Whether the table has a totals row; needs `commit()` once changed. */
    totalsRow: boolean | undefined;
    /** The table style (`"TableStyleMedium2"`), or `null` for none (also for a table read from a file without one). */
    theme: TableStyleProperties["theme"];
    showFirstColumn: boolean | undefined;
    showLastColumn: boolean | undefined;
    showRowStripes: boolean | undefined;
    showColumnStripes: boolean | undefined;
    /** The number of columns. */
    readonly width: number;
    /** The number of data rows (of a table read from a file: those of its range). */
    readonly height: number;

    /**
     * Commit changes (writes the table to the cells of the worksheet); does nothing without a change.
     * @throws Error for invalid table properties (e.g. a row with more values than the table has columns, a table that
     *   does not fit in the sheet): the cells are not changed then
     */
    commit(): void;
    /**
     * Remove rows of data (0-based index; 1 row by default)
     * @throws Error for a table read from a file
     */
    removeRows(rowIndex: number, count?: number): void;
    /**
     * Add a row of data, either insert at rowNumber (0-based) or append
     * @throws Error for more values than the table has columns (the table is not changed); for a table read from a file
     */
    addRow(values: CellValue[], rowNumber?: number): void;
    /**
     * Get column (0-based index)
     * @throws Error for an index that is not a column of the table
     */
    getColumn(colIndex: number): TableColumn;
    /**
     * Add a new column, including column defn and values (one per data row); inserts at colIndex (0-based) or adds to
     * the right
     * @throws Error for more values than the table has rows (the table is not changed); for a table read from a file
     */
    addColumn(column: TableColumnProperties, values: CellValue[], colIndex?: number): void;
    /**
     * Remove columns with their data (0-based index; 1 column by default)
     * @throws Error for a table read from a file
     */
    removeColumns(colIndex: number, count?: number): void;
}

/**
 * The streaming API (Node.js and Bun only: not in the browser bundles). `stream.xlsx.WorkbookWriter` writes a workbook
 * row by row without holding it in memory; `stream.xlsx.WorkbookReader` reads one in a single pass. Their worksheets are
 * not document {@link Worksheet}s: see {@link stream.xlsx.WorksheetWriter} and {@link stream.xlsx.WorksheetReader}.
 */
export namespace stream {
    namespace xlsx {
        /**
         * The output and the content options of the streaming {@link WorkbookWriter} (also the base of
         * {@link XlsxWriteOptions}).
         */
        interface WorkbookWriterOptions {
            /**
             * The writable stream the archive is written to. Without `stream` nor `filename`, the archive is written to
             * an in-memory stream (`writer.stream`).
             */
            stream: NodeJS.WritableStream;

            /**
             * The path of the file to write the archive to, when there is no `stream`.
             */
            filename: string;

            /**
             * Write the strings to the shared strings table (smaller file, more memory) instead of inline.
             * @default false (streaming WorkbookWriter), true (`xlsx.write`, `writeBuffer`, `writeFile`)
             */
            useSharedStrings: boolean;

            /**
             * Write the styles of the cells (some performance overhead).
             * @default false (streaming WorkbookWriter), true (`xlsx.write`, `writeBuffer`, `writeFile`)
             */
            useStyles: boolean;
        }

        /**
         * @deprecated use {@link ZipOptions} (archiver is no longer used)
         */
        type ArchiverZipOptions = ZipOptions;

        /**
         * @deprecated archiver's `zlib` option (ExcelJS < 5): only `level` is read, see {@link ZipOptions.zlib}
         */
        type ZlibOptions = ZipOptions["zlib"];

        /** The options of the {@link WorkbookWriter} constructor (all optional). */
        interface WorkbookStreamWriterOptions extends WorkbookWriterOptions {
            /**
             * The zip options (compression level, ZIP64, archive comment).
             * @default { level: 6 }
             */
            zip: Partial<ZipOptions>;
            /**
             * Maximum number of cells covered by merges in all the worksheets (one worksheet holds at most 1,048,576).
             * Reading the file back with `xlsx.load` needs a `maxMergedCells` load option at least as large.
             * @default 1048576 (one full worksheet)
             */
            maxMergedCells: number;
            /**
             * Document property: the creation date.
             * @default now
             */
            created: Date;
            /**
             * Document property: the date of the last change.
             * @default `created`
             */
            modified: Date;
            /**
             * Document property: the author.
             * @default "ExcelJS"
             */
            creator: string;
            /**
             * Document property: the author of the last change.
             * @default "ExcelJS"
             */
            lastModifiedBy: string;
            /** Document property: the date of the last print. */
            lastPrinted: Date;
        }

        /** An image added with {@link WorkbookWriter.addImage}: the image given, with its part name (`image0.png`). */
        type WorkbookWriterImage = Image & {
            /** Always `"image"`. */
            type: "image";
            /** The name of its part in `xl/media/` (`image0.png`). */
            name: string;
        };

        /**
         * The streaming xlsx writer: each worksheet is written to the archive as its rows are committed, so that a large
         * workbook is never held in memory. {@link commit} commits the worksheets left, adds the workbook parts and
         * completes the archive. Node.js and Bun only.
         */
        class WorkbookWriter {
            /**
             * @throws TypeError for an invalid `maxMergedCells`
             */
            constructor(options?: Partial<WorkbookStreamWriterOptions>);

            /** Document property: the creation date (now by default). */
            created: Date;
            /** Document property: the date of the last change (`created` by default). */
            modified: Date;
            /** Document property: the author ("ExcelJS" by default). */
            creator: string;
            /** Document property: the author of the last change ("ExcelJS" by default). */
            lastModifiedBy: string;
            /** Document property: the date of the last print. */
            lastPrinted: Date | undefined;
            /** The default of the worksheets for writing strings to the shared strings table. */
            useSharedStrings: boolean;
            /** The workbook views (windows), written by {@link commit}. */
            views: WorkbookView[];
            /** The cap on the cells covered by merges in all the worksheets (`maxMergedCells` option). */
            maxMergedCells: number;
            /** Where the archive is written: the `stream` option, the `filename` file, or an in-memory stream. */
            stream: NodeJS.WritableStream;
            /** The defined names of the workbook, written by {@link commit}. */
            readonly definedNames: DefinedNames;
            /** The id that the next worksheet added gets: the first free one. */
            readonly nextId: number;

            /**
             * Commits the worksheets not committed yet, adds the workbook parts and completes the archive.
             *
             * @returns resolves with this writer once the output stream has finished
             * @throws (rejects) on a failure of the zip writer or of the output stream, or an image without data
             */
            commit(): Promise<this>;

            /**
             * Adds an image (`buffer`, `base64` or `filename`), written by {@link commit}.
             *
             * @returns its id, for {@link WorksheetWriter.addBackgroundImage}
             * @throws Error for an extension that is not 1 to 5 letters or digits
             */
            addImage(image: Image): number;
            /** An image added with {@link addImage}, `undefined` for an unknown id. */
            getImage(id: number): WorkbookWriterImage | undefined;

            /**
             * Adds a worksheet, written as its rows are committed.
             *
             * @param name - default `sheet<id>`
             */
            addWorksheet(name?: string, options?: Partial<AddWorksheetWriterOptions>): WorksheetWriter;
            /**
             * The first worksheet (no argument), the worksheet with this id, or the one with this name; `undefined` when
             * there is none.
             */
            getWorksheet(id?: number | string): WorksheetWriter | undefined;

            /** A step of {@link commit}: adds the styles part. Not meant to be called directly. */
            addStyles(): Promise<void>;
            /** A step of {@link commit} (done by the constructor): adds the theme part. Not meant to be called directly. */
            addThemes(): Promise<void>;
            /**
             * A step of {@link commit} (done by the constructor): adds the package relationships. Not meant to be called
             * directly.
             */
            addOfficeRels(): Promise<void>;
            /** A step of {@link commit}: adds `[Content_Types].xml`. Not meant to be called directly. */
            addContentTypes(): Promise<void>;
            /**
             * A step of {@link commit}: adds the images.
             *
             * @throws Error for an image without data (synchronously)
             */
            addMedia(): Promise<void[]>;
            /** A step of {@link commit}: adds the extended properties. Not meant to be called directly. */
            addApp(): Promise<void>;
            /** A step of {@link commit}: adds the core properties. Not meant to be called directly. */
            addCore(): Promise<void>;
            /** A step of {@link commit}: adds the shared strings part, when used. Not meant to be called directly. */
            addSharedStrings(): Promise<void>;
            /** A step of {@link commit}: adds the workbook relationships. Not meant to be called directly. */
            addWorkbookRels(): Promise<void>;
            /** A step of {@link commit}: adds the workbook part. Not meant to be called directly. */
            addWorkbook(): Promise<void>;
        }

        /** The options of {@link WorkbookWriter.addWorksheet}. */
        interface AddWorksheetWriterOptions extends AddWorksheetOptions {
            /**
             * Write the strings of this worksheet to the shared strings table.
             * @default the `useSharedStrings` option of the workbook writer
             */
            useSharedStrings: boolean;
            /** The auto filter of the worksheet. */
            autoFilter: AutoFilter | null;
            /**
             * @deprecated use `properties.tabColor`
             */
            tabColor: Partial<Color>;
        }

        /**
         * A worksheet of the streaming {@link WorkbookWriter} (`writer.addWorksheet()`): not a document
         * {@link Worksheet}. Rows are written, in order, when they are committed (`row.commit()`) and can no longer be
         * reached; the rest of the worksheet (merges, validations, conditional formatting, page setup...) is written when
         * the worksheet is committed.
         */
        interface WorksheetWriter {
            /** The id of the worksheet in the workbook. */
            readonly id: number;
            /** The worksheet name (`Sheet<id>` by default). */
            name: string;
            /** The visibility of the worksheet. */
            state: WorksheetState;
            /** The workbook writer. */
            readonly workbook: WorkbookWriter;
            /** True once {@link commit} was called: rows can no longer be added. */
            readonly committed: boolean;
            /** The data validations of the cells, written by {@link commit}. */
            dataValidations: DataValidations;
            /** The conditional formattings (see {@link addConditionalFormatting}). */
            conditionalFormatting: ConditionalFormattingOptions[];
            /** The manual page breaks (`row.addPageBreak()`). */
            rowBreaks: RowBreak[];
            /** The worksheet properties (default row height 15, dyDescent 55, outline levels 0, and those given). */
            properties: Partial<WorksheetProperties> &
                Pick<WorksheetProperties, "defaultRowHeight" | "dyDescent" | "outlineLevelCol" | "outlineLevelRow">;
            /** Headers and footers of the printed pages (`null` by default). */
            headerFooter: WorksheetHeaderFooter;
            /** For all things printing (margins, orientation, fit to page...): the defaults, and those given. */
            pageSetup: Partial<PageSetup>;
            /** Write the strings of this worksheet to the shared strings table. */
            useSharedStrings: boolean;
            /** The auto filter, `null` when there is none. */
            autoFilter: AutoFilter | null;
            /** The protection set by {@link protect}, `null` when the worksheet is not protected. */
            sheetProtection:
                | (Partial<WorksheetProtection> & {
                      /** Always true: the worksheet is protected. */
                      sheet?: boolean;
                      /** `"SHA-512"` when there is a password. */
                      algorithmName?: string;
                      /** The hash of the password (base64). */
                      hashValue?: string;
                      /** The salt of the hash (base64). */
                      saltValue?: string;
                  })
                | null;
            /** The range covering the rows added so far. */
            readonly dimensions: Range;
            /** The sheet views (the `views` option of `addWorksheet`). */
            readonly views: WorksheetViewModel[];

            /**
             * Not a valid operation for a streaming writer.
             *
             * @throws Error always
             */
            destroy(): never;

            /**
             * Writes the rows not committed yet and the rest of the worksheet. Rows can no longer be added.
             *
             * @returns a promise to await before writing more (backpressure)
             */
            commit(): Promise<void>;

            /** The columns defined or used, `null` when there is none. */
            get columns(): Column[] | null;
            /** Defines the columns (headers, keys, widths...). Any header defined overwrites the values of its cells. */
            set columns(value: readonly ColumnDefinition[]);
            /** The column of a key (`column.key`), `undefined` for an unknown key. */
            getColumnKey(key: string): Column | undefined;
            /** Registers the column of a key (done by the `column.key` setter). */
            setColumnKey(key: string, value: Column): void;
            /** Unregisters a key (done by the `column.key` setter). */
            deleteColumnKey(key: string): void;
            /** Calls `callback` with each keyed column and its key. */
            eachColumnKey(callback: (column: Column, key: string) => void): void;
            /**
             * A column by key, letters or 1-based number; it and the columns before it are created if needed.
             * `undefined` for a number below 1 (an invalid argument).
             *
             * @throws Error for letters that are not a column
             */
            getColumn(indexOrKey: number | string): Column;

            /** The last row not committed yet, if any. */
            readonly lastRow: Row | undefined;
            /** The row (1-based number) if it exists and is not committed. */
            findRow(rowNumber: number): Row | undefined;
            /**
             * The row (1-based number), created if needed.
             *
             * @throws Error for a number that is not a row (an integer from 1 to 1,048,576), or when the row was
             * committed
             */
            getRow(rowNumber: number): Row;
            /**
             * Adds a row after the last one, with values (an array by column number, or an object by column key).
             *
             * @throws Error when the worksheet was committed
             */
            addRow(values: RowValues): Row;
            /** Calls `callback` with each row not committed yet that has values. */
            eachRow(callback: (row: Row, rowNumber: number) => void): void;
            /** Calls `callback` with each row not committed yet (`includeEmpty`: each one, created if needed). */
            eachRow(options: { includeEmpty?: boolean } | undefined, callback: (row: Row, rowNumber: number) => void): void;

            /**
             * The cell at an address (`"B3"`, `"$B$3"`, `"Sheet1!B3"`), if it exists and its row is not committed:
             * `undefined` for an address that is not a cell of this worksheet (see the document `Worksheet.findCell`).
             */
            findCell(address: string): Cell | undefined;
            /** The cell at (row, column), if it exists and its row is not committed (`undefined` out of bounds). */
            findCell(row: number, col: number): Cell | undefined;
            /**
             * The cell at an address (`"B3"`, `"$B$3"`, `"Sheet1!B3"`: a sheet name must be this worksheet's), created
             * if needed.
             *
             * @throws Error for an invalid address or a cell of another sheet (see the document `Worksheet.getCell`), or
             * when its row was committed
             */
            getCell(address: string): Cell;
            /**
             * The cell at (row, column), created if needed.
             *
             * @throws Error for a row or a column out of bounds, or when its row was committed
             */
            getCell(row: number, col: number): Cell;

            /**
             * Merges a range of cells (`"A4:B5"`, a {@link Range} or {@link Location}, or an array of the arguments
             * below): every cell but the top-left one is merged into it. A sheet name, wherever given, must be this
             * worksheet's (case-insensitive): the merge has none.
             *
             * @throws Error for an invalid range or a range of another sheet (see the document
             * `Worksheet.mergeCells`), when a cell is already merged, a row was committed, or the merged cells exceed
             * the limits
             */
            mergeCells(range: string | Range | Location | [string, string, string?] | [number, number, number, number, string?]): void;
            /** Merges the range from `tl` to `br` (`"A4"`, `"B5"`). */
            mergeCells(tl: string, br: string, sheetName?: string): void;
            /** Merges the range of these bounds (1-based numbers). */
            mergeCells(top: number, left: number, bottom: number, right: number, sheetName?: string): void;

            /** Adds conditional formatting rules, written when the worksheet is committed. */
            addConditionalFormatting(cf: ConditionalFormattingOptions): void;
            /** Removes the conditional formatting at an index, those a filter rejects, or (no argument) all of them. */
            removeConditionalFormatting(filter?: number | ((value: ConditionalFormattingOptions, index: number) => unknown)): void;

            /** Sets the background image (an id of {@link WorkbookWriter.addImage}). */
            addBackgroundImage(imageId: number): void;
            /** The id of the background image, if any. */
            getBackgroundImageId(): number | undefined;

            /**
             * Protects the worksheet, with a password hashed with SHA-512 (`spinCount` iterations, 100000 by default).
             * Mutates `options`: `spinCount` is made an integer >= 0.
             */
            protect(password?: string, options?: Partial<WorksheetProtection>): Promise<void>;
            /** Removes the protection. */
            unprotect(): void;
        }

        /** The options of the {@link WorkbookReader}: what to do with each kind of part, and the decompression limits. */
        interface WorkbookStreamReaderOptions extends Partial<ZipReadLimits> {
            /**
             * `emit`: the worksheets are read (`worksheet` events, async iteration); `ignore`: skipped.
             * @default 'emit'
             */
            worksheets?: "emit" | "ignore";
            /**
             * `cache`: kept for the cells; `emit`: emitted one by one (`shared-strings` events: {@link EmittedSharedString}),
             * and the cells then hold their index ({@link CellSharedStringValue}); `ignore`: not read, the cells hold their
             * index.
             * @default 'cache'
             */
            sharedStrings?: "cache" | "emit" | "ignore";
            /**
             * The hyperlinks of the worksheets: `cache`: the `<hyperlink>` elements in `worksheetReader.hyperlinks`, with
             * the targets of the external links (read from the relationship parts: see {@link WorksheetHyperlink.target});
             * `emit`: `hyperlink` events of the worksheet readers, and `hyperlinks` events (a {@link HyperlinkReader} per
             * worksheet, for the targets) of the workbook reader; `ignore`. The cells of the rows do not get the
             * hyperlinks (a worksheet lists its hyperlinks after its rows).
             * @default 'ignore'
             */
            hyperlinks?: "cache" | "emit" | "ignore";
            /**
             * `cache`: the cells get their styles; `ignore`: the styles are not read (dates then come as serial numbers).
             * @default 'ignore'
             */
            styles?: "cache" | "ignore";
            /**
             * `emit`: an `entry` event for each part read.
             * @default 'ignore'
             */
            entries?: "emit" | "ignore";
        }

        /** The options of a {@link WorkbookReader}, defaults applied. */
        type WorkbookStreamReaderResolvedOptions = Required<Omit<WorkbookStreamReaderOptions, keyof ZipReadLimits>> &
            Partial<ZipReadLimits>;

        /**
         * A file name, a Node.js stream (or any object with a `pipe` method), a web ReadableStream or an
         * async iterable of byte chunks
         */
        type WorkbookReaderInput =
            string | import("stream").Stream | import("stream/web").ReadableStream<Uint8Array> | AsyncIterable<Uint8Array>;

        /** The payload of an `entry` event (`entries: "emit"`): the kind of part, and the number of a worksheet part. */
        interface WorkbookReaderEntry {
            /** The kind of part. */
            type: "workbook" | "shared-strings" | "styles" | "worksheet" | "hyperlinks";
            /** The number of the worksheet part (`sheet3.xml`: `"3"`), for `worksheet` and `hyperlinks`. */
            id?: string;
        }

        /** A shared string emitted by the reader (`sharedStrings: "emit"`): its index in the table and its value. */
        interface EmittedSharedString {
            /** The index of the string, as held by the cells ({@link CellSharedStringValue}). */
            index: number;
            /** The string or rich text. */
            text: string | CellRichTextValue;
        }

        /** The events of a {@link WorkbookReader} (emitted by {@link WorkbookReader.read}), and their arguments. */
        interface WorkbookReaderEvents {
            /** A worksheet (`worksheets: "emit"`): listen to its events; it is read once the listeners return. */
            worksheet: [worksheet: WorksheetReader];
            /** A shared string (`sharedStrings: "emit"`). */
            "shared-strings": [sharedString: EmittedSharedString];
            /** The hyperlink relationships of a worksheet (`hyperlinks: "emit"`); read once the listeners return. */
            hyperlinks: [hyperlinks: HyperlinkReader];
            /** A part is read (`entries: "emit"`). */
            entry: [entry: WorkbookReaderEntry];
            /** The file was read (followed by `finished`). */
            end: [];
            /** The file was read. */
            finished: [];
            /** The file could not be read (nothing follows). */
            error: [error: unknown];
        }

        /** An event of {@link WorkbookReader.parse}. */
        type WorkbookReaderEvent =
            | { readonly eventType: "shared-strings"; readonly value: EmittedSharedString }
            | { readonly eventType: "worksheet"; readonly value: WorksheetReader }
            | { readonly eventType: "hyperlinks"; readonly value: HyperlinkReader };

        /**
         * The streaming xlsx reader: reads the parts as they come out of the zip, in one pass, holding only the shared
         * strings (when cached), the styles (when cached) and a worksheet at a time (the rows are not kept). Node.js and
         * Bun only.
         *
         * Iterate it (`for await (const worksheet of reader)`), then each worksheet for its rows; or call {@link read}
         * and listen to its events ({@link WorkbookReaderEvents}). A worksheet that comes before the parts it depends on
         * is spooled to a temp file.
         */
        class WorkbookReader extends EventEmitter<WorkbookReaderEvents> {
            /** The valid values of the options. */
            static readonly Options: {
                /** The values of `worksheets`. */
                worksheets: string[];
                /** The values of `sharedStrings`. */
                sharedStrings: string[];
                /** The values of `hyperlinks`. */
                hyperlinks: string[];
                /** The values of `styles`. */
                styles: string[];
                /** The values of `entries`. */
                entries: string[];
            };

            /**
             * @param input - what to read (else the input given to {@link read} / {@link parse})
             */
            constructor(input?: WorkbookReaderInput, options?: Partial<WorkbookStreamReaderOptions>);

            /** The input given to the constructor (read when {@link read} / {@link parse} get none). */
            input: WorkbookReaderInput | undefined;
            /** The options, defaults applied. */
            options: WorkbookStreamReaderResolvedOptions;
            /** `sharedStrings: "cache"`: the shared strings table, once read. */
            sharedStrings?: (string | CellRichTextValue)[];

            /**
             * Reads the whole file, emitting the events ({@link WorkbookReaderEvents}); each worksheet and hyperlinks
             * part emitted is read once its listeners return.
             *
             * @param input - replaces the constructor's
             * @param options - replace the constructor's, on top of the defaults
             * @returns resolves once the file is read; errors are emitted (`error`), not thrown
             */
            read(input?: WorkbookReaderInput, options?: Partial<WorkbookStreamReaderOptions>): Promise<void>;
            /**
             * The worksheets of the file, each to be read (iterated) before the next one comes.
             *
             * @throws Error on an unrecognised input, an invalid zip, a decompression limit exceeded, or invalid XML
             */
            [Symbol.asyncIterator](): AsyncGenerator<WorksheetReader, void, undefined>;
            /**
             * Reads the parts of the file as they come.
             *
             * @param input - replaces the constructor's
             * @param options - replace the constructor's, on top of the defaults
             * @returns the worksheets, shared strings and hyperlinks, as the options select them
             * @throws Error on an unrecognised input, an invalid zip, a decompression limit exceeded, or invalid XML
             */
            parse(
                input?: WorkbookReaderInput,
                options?: Partial<WorkbookStreamReaderOptions>
            ): AsyncGenerator<WorkbookReaderEvent, void, undefined>;
        }

        /**
         * A `<hyperlink>` of a worksheet, as read (`hyperlinks: "cache"` or `"emit"`): the attributes present, and with
         * `hyperlinks: "cache"` the target of its relationship.
         */
        interface WorksheetHyperlink {
            /** The cell reference (`"A1"`). */
            ref: string | undefined;
            /** The relationship of an external link (`r:id`), see {@link HyperlinkReader}. */
            rId?: string;
            /** An internal link: a location in the workbook. */
            location?: string;
            /** The tooltip. */
            tooltip?: string;
            /** The text displayed. */
            display?: string;
            /**
             * `hyperlinks: "cache"`: the link of an external hyperlink (the `Target` of its relationship `rId`), set once
             * the relationships part of the worksheet is read: when the worksheet is read if that part comes first in the
             * file, else when the reader reaches it (at the latest when the iteration of the workbook ends). Absent when
             * the relationship is missing.
             */
            target?: string;
            /** `hyperlinks: "cache"`: the `TargetMode` of the relationship (`"External"` for a link outside the package). */
            targetMode?: string;
        }

        /** The events of a {@link WorksheetReader} (emitted by {@link WorksheetReader.read}), and their arguments. */
        interface WorksheetReaderEvents {
            /** A row, with its cells. */
            row: [row: Row];
            /** A hyperlink (`hyperlinks: "emit"`). */
            hyperlink: [hyperlink: WorksheetHyperlink];
            /** The worksheet was read. */
            finished: [];
            /** The worksheet could not be read (nothing follows). */
            error: [error: unknown];
        }

        /** An event of {@link WorksheetReader.parse}: a row, or a hyperlink (`hyperlinks: "emit"`). */
        type WorksheetReaderEvent =
            { readonly eventType: "row"; readonly value: Row } | { readonly eventType: "hyperlink"; readonly value: WorksheetHyperlink };

        /**
         * A worksheet of the streaming {@link WorkbookReader}: not a document {@link Worksheet}. Its rows come one at a
         * time (async iteration, or `row` events with {@link read}) and are not kept; a row read is read-only.
         */
        interface WorksheetReader extends EventEmitter<WorksheetReaderEvents> {
            /** The workbook reader. */
            readonly workbook: WorkbookReader;
            /**
             * The sheetId from the workbook, or the number of the part (a string) if the worksheet could not be matched to
             * the workbook
             */
            id: number | string;
            /** The worksheet name (`Sheet<id>` if the worksheet could not be matched to the workbook). */
            name: string;
            /** The visibility of the worksheet (from the workbook), if known. */
            state?: WorksheetState;
            /** The `worksheets` and `hyperlinks` options of the workbook reader. */
            readonly options: Partial<Pick<WorkbookStreamReaderOptions, "worksheets" | "hyperlinks">>;
            /**
             * `hyperlinks: "cache"`: the hyperlinks read so far, keyed by cell reference (an object without prototype),
             * with their targets (see {@link WorksheetHyperlink.target}).
             */
            hyperlinks?: Record<string, WorksheetHyperlink>;
            /** The sheet format properties (outline levels, default sizes), read from `<sheetFormatPr>`. */
            properties: Partial<WorksheetProperties> & Pick<WorksheetProperties, "outlineLevelCol" | "outlineLevelRow">;
            /** The range covering the rows read so far. */
            readonly dimensions: Range;
            /** Every column up to the last one read or used (created when it is read), `null` when there is none. */
            readonly columns: Column[] | null;

            /**
             * Not a valid operation for a streaming reader.
             *
             * @throws Error always
             */
            destroy(): never;

            /**
             * A column by key, letters or 1-based number, created if needed. `undefined` for a number that is not a
             * positive integer (an invalid argument).
             *
             * @throws Error for letters that are not a column
             */
            getColumn(indexOrKey: number | string): Column;
            /** The column of a key (`column.key`), `undefined` for an unknown key. */
            getColumnKey(key: string): Column | undefined;
            /** Registers the column of a key (done by the `column.key` setter). */
            setColumnKey(key: string, value: Column): void;
            /** Unregisters a key (done by the `column.key` setter). */
            deleteColumnKey(key: string): void;
            /** Calls `callback` with each keyed column and its key. */
            eachColumnKey(callback: (column: Column, key: string) => void): void;

            /**
             * Reads the whole worksheet, emitting its events ({@link WorksheetReaderEvents}).
             *
             * @returns resolves once the worksheet is read; errors are emitted (`error`), not thrown
             */
            read(): Promise<void>;
            /**
             * The rows of the worksheet.
             *
             * @throws Error on malformed XML or a row number outside the grid
             */
            [Symbol.asyncIterator](): AsyncGenerator<Row, void, undefined>;
            /**
             * Reads the worksheet.
             *
             * @returns its events, in batches (one per chunk of the part that has some)
             * @throws Error on malformed XML or a row number outside the grid
             */
            parse(): AsyncGenerator<WorksheetReaderEvent[], void, undefined>;
        }

        /** A hyperlink relationship of a worksheet, as read. The values come from the file; a missing one is `undefined`. */
        interface HyperlinkRelationship {
            /** Always `RelationshipType.Styles` (5), as in ExcelJS < 5 (kept for compatibility). */
            type: RelationshipType;
            /** The relationship id, referenced by the `rId` of a {@link WorksheetHyperlink}. */
            rId: string | undefined;
            /** The link. */
            target: string | undefined;
            /** `"External"` for a link outside the package. */
            targetMode: string | undefined;
        }

        /** The events of a {@link HyperlinkReader} (emitted by {@link HyperlinkReader.read}), and their arguments. */
        interface HyperlinkReaderEvents {
            /** A hyperlink relationship (`hyperlinks: "emit"`). */
            hyperlink: [relationship: HyperlinkRelationship];
            /** The part was read. */
            finished: [];
            /** The part could not be read. */
            error: [error: unknown];
        }

        /**
         * The hyperlink relationships (link targets) of a worksheet: the `hyperlinks` event of the {@link WorkbookReader},
         * with `hyperlinks: "emit"`. {@link read} emits them one by one (`hyperlink` events). The cache members
         * ({@link hyperlinks}, {@link count}, {@link each}) are filled only by a reader of the `"cache"` option, which the
         * workbook reader never hands out (with `hyperlinks: "cache"`, it gives the targets to the hyperlinks of the
         * worksheet readers instead): they stay empty.
         */
        interface HyperlinkReader extends EventEmitter<HyperlinkReaderEvents> {
            /** The workbook reader. */
            readonly workbook: WorkbookReader;
            /** The number of the worksheet part (`sheet3.xml.rels`: `"3"`). */
            readonly id: string;
            /** The `hyperlinks` option of the workbook reader. */
            readonly options: Pick<WorkbookStreamReaderResolvedOptions, "hyperlinks">;
            /** `hyperlinks: "cache"`: the relationships read, keyed by relationship id (an object without prototype). */
            hyperlinks?: Record<string, HyperlinkRelationship>;
            /** The number of relationships cached (0 before {@link read} and when they are not cached). */
            readonly count: number;
            /** Calls `callback` with each cached relationship, in the order of the part. */
            each(callback: (relationship: HyperlinkRelationship, index: number) => void): void;
            /**
             * Reads the relationships (once: later calls return the same promise).
             *
             * @returns resolves once the part is read; errors are emitted (`error`), not thrown
             */
            read(): Promise<void>;
        }
    }
}
