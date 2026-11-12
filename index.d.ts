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
     * Presentation style
     */
    style: "pageBreakPreview" | "pageLayout";
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
 * A view of `worksheet.views`: the fields of the normal, frozen and split views, all optional. A view read from a file
 * is not narrowed to the variant of its `state`; a {@link WorksheetView} can be assigned.
 */
export type WorksheetViewModel = Partial<WorksheetViewCommon & Omit<WorksheetViewFrozen, "state"> & Omit<WorksheetViewSplit, "state">> & {
    state?: WorksheetView["state"];
};

/**
 * The style option of `addRow` / `insertRow` and their plural forms: `"i"` copies the style of the row above (`"i+"`:
 * also its empty cells), `"o"` the style of the row below (`"o+"`, `insertRow` only), `"n"` (the default) none.
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

export interface FillPattern {
    type: "pattern";
    pattern: FillPatterns;
    fgColor?: Partial<Color>;
    bgColor?: Partial<Color>;
}

export interface GradientStop {
    position: number;
    color: Partial<Color>;
}

export interface FillGradientAngle {
    type: "gradient";
    gradient: "angle";

    /**
     * For 'angle' gradient, specifies the direction of the gradient. 0 is from the left to the right.
     * Values from 1 - 359 rotates the direction clockwise
     */
    degree: number;

    /**
     * Specifies the gradient colour sequence. Is an array of objects containing position and
     * color starting with position 0 and ending with position 1.
     * Intermediary positions may be used to specify other colours on the path.
     */
    stops: GradientStop[];
}

export interface FillGradientPath {
    type: "gradient";
    gradient: "path";

    /**
     * For 'path' gradient. Specifies the relative coordinates for the start of the path.
     * 'left' and 'top' values range from 0 to 1
     */
    center: { left: number; top: number };

    /**
     * Specifies the gradient colour sequence. Is an array of objects containing position and
     * color starting with position 0 and ending with position 1.
     * Intermediary positions may be used to specify other colours on the path.
     */
    stops: GradientStop[];
}

export type Fill = FillPattern | FillGradientAngle | FillGradientPath;

export interface Font {
    name: string;
    size: number;
    family: number;
    scheme: "minor" | "major" | "none";
    charset: number;
    color: Partial<Color>;
    bold: boolean;
    italic: boolean;
    underline: boolean | "none" | "single" | "double" | "singleAccounting" | "doubleAccounting";
    vertAlign: "superscript" | "subscript";
    strike: boolean;
    outline: boolean;
}

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

export interface Color {
    /**
     * Hex string for alpha-red-green-blue e.g. FF00FF00
     */
    argb: string;

    /**
     * Choose a theme by index
     */
    theme: number;
}

export interface Border {
    style: BorderStyle;
    color: Partial<Color>;
}

export interface BorderDiagonal extends Border {
    up: boolean;
    down: boolean;
}

export interface Borders {
    top: Partial<Border>;
    left: Partial<Border>;
    bottom: Partial<Border>;
    right: Partial<Border>;
    diagonal: Partial<BorderDiagonal>;
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

export interface Alignment {
    horizontal: "left" | "center" | "right" | "fill" | "justify" | "centerContinuous" | "distributed";
    vertical: "top" | "middle" | "bottom" | "distributed" | "justify";
    wrapText: boolean;
    shrinkToFit: boolean;
    indent: number;
    readingOrder: "rtl" | "ltr";
    textRotation: number | "vertical";
}

export interface Protection {
    locked: boolean;
    hidden: boolean;
}

export interface Style {
    numFmt: string;
    font: Partial<Font>;
    alignment: Partial<Alignment>;
    protection: Partial<Protection>;
    border: Partial<Borders>;
    fill: Fill;
}

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

export interface DataValidation {
    type: DataValidationType;
    /** The formulas or values of the rule (`["A1:A5"]`, `[1, 10]`, `[new Date()]`); absent for `"any"`. */
    formulae?: (string | number | Date)[];
    allowBlank?: boolean;
    operator?: DataValidationOperator;
    error?: string;
    errorTitle?: string;
    errorStyle?: string;
    prompt?: string;
    promptTitle?: string;
    showErrorMessage?: boolean;
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

export interface RichText {
    text: string;
    font?: Partial<Font>;
}

export interface CellRichTextValue {
    richText: RichText[];
}

export interface CellHyperlinkValue {
    text: string;
    hyperlink: string;
    tooltip?: string;
}

/** The cached result of a formula. */
export type CellFormulaResult = number | string | boolean | Date | CellErrorValue;

export interface CellFormulaValue {
    formula: string;
    result?: CellFormulaResult;
    date1904?: boolean;
}

export interface CellSharedFormulaValue {
    sharedFormula: string;
    readonly formula?: string;
    result?: number | string | boolean | Date | CellErrorValue;
    date1904?: boolean;
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

export interface CommentMargins {
    insetmode: "auto" | "custom";
    inset: Number[];
}

export interface CommentProtection {
    locked: "True" | "False";
    lockText: "True" | "False";
}

export type CommentEditAs = "twoCells" | "oneCells" | "absolute";

export type ImageEditAs = "twoCell" | "oneCell" | "absolute";

export interface Comment {
    texts?: RichText[];
    margins?: Partial<CommentMargins>;
    protection?: Partial<CommentProtection>;
    editAs?: CommentEditAs;
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
    readonly text: string;
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
     * Assign (or get) an array of names for a cell (cells can have more than one name)
     */
    names: string[];

    /**
     * Cells can define what values are valid or not and provide prompting to the user to help guide them; `undefined`
     * when the cell has no validation.
     */
    get dataValidation(): DataValidation | undefined;
    set dataValidation(value: DataValidation);

    /**
     * Value of the cell
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
     * The type of the cell's formula: Master or Shared for a formula cell, `undefined` for other cells
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
    merge(master: Cell, ignoreStyle?: boolean): void;
    unmerge(ignoreStyle?: boolean): void;
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

export interface Row extends StyleProperties {
    readonly worksheet: Worksheet;
    readonly hasValues: boolean;
    /**
     * The min and max column numbers of the non-empty cells, or null if the row is empty
     */
    readonly dimensions: { min: number; max: number } | null;
    /** The model of the row; `null` for a row without height and without cell. */
    model: RowModel | null;
    /**
     * Set a specific row height; `undefined` when the row has the default height
     */
    height: number | undefined;

    /**
     * Make row hidden
     */
    hidden: boolean;

    /**
     * Get a row as a sparse array (index 0 is unused: column A is at index 1)
     */
    get values(): CellValue[];
    /**
     * Set the row's values from a contiguous or sparse array, or from an object keyed by column key
     */
    set values(value: CellValue[] | { [key: string]: CellValue });

    /**
     * Set an outline level for rows
     */
    outlineLevel?: number;

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

    findCell(colNumber: number): Cell | undefined;

    getCellEx(address: Address): Cell;

    /**
     * Iterate over all non-null cells in a row
     */
    eachCell(callback: (cell: Cell, colNumber: number) => void): void;

    /**
     * Iterate over all cells in a row (including empty cells)
     */
    eachCell(opt: { includeEmpty: boolean }, callback: (cell: Cell, colNumber: number) => void): void;

    /**
     * Cut one or more cells (cells to the right are shifted left)
     *
     * Note: this operation will not affect other rows
     */
    splice(start: number, count: number, ...insert: any[]): void;

    /**
     * Commit a completed row (and the rows before it) to stream. In the streaming writer, returns a promise that resolves
     * once the output can take more data: awaiting it keeps memory bounded (backpressure). No effect on a document worksheet.
     */
    commit(): void | Promise<void>;
    destroy(): void;
    addPageBreak(lft?: number, rght?: number): void;
}

export interface Column {
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
     * Set an outline level for columns
     */
    outlineLevel: number;

    /**
     * Hides the column
     */
    hidden: boolean;

    /**
     * Styles applied to the column
     */
    style: Partial<Style>;

    /**
     * The cell values in the column, as a sparse array indexed by row number
     */
    values: CellValue[];

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

    /**
     * Below properties read from style
     */
    border?: Partial<Borders>;
    fill?: Fill;
    numFmt?: string;
    font?: Partial<Font>;
    alignment?: Partial<Alignment>;
    protection?: Partial<Protection>;

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
     * Iterate over all current cells in this column including empty cells
     */
    eachCell(opt: { includeEmpty: boolean }, callback: (cell: Cell, rowNumber: number) => void): void;

    /**
     * The definition of the column (what `worksheet.columns = [...]` sets): setting it sets the key, width (default 9),
     * outline level, style, header and hidden state; `null` / `undefined` reset nothing.
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
     */
    margins: Margins;

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
     * Where to place comments
     *
     * Default is `None`
     */
    cellComments: "atEnd" | "asDisplayed" | "None";

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
    paperSize: PaperSize;

    /**
     * Whether to show the row numbers and column letters, `false` by default
     */
    showRowColHeaders: boolean;

    /**
     * Whether to show grid lines, `false` by default
     */
    showGridLines: boolean;

    /**
     * Which number to use for the first page
     */
    firstPageNumber: number;

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
export interface ImageRange {
    tl: Anchor;
    br: Anchor;
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
        hyperlinks?: ImageHyperlinkValue;
    };
}

export interface ImagePosition {
    tl: { col: number; row: number };
    ext: { width: number; height: number };
}

export interface ImageHyperlinkValue {
    hyperlink: string;
    tooltip?: string;
}

export interface Range extends Location {
    /** The worksheet name of a reference such as `Sheet1!A1:B2`, `undefined` without one. */
    sheetName: string | undefined;

    tl: string;
    $t$l: string;

    br: string;
    $b$r: string;

    range: string;
    $range: string;

    shortRange: string;
    $shortRange: string;

    count: number;

    /**
     * Sets the range from the list of arguments of a range (as `worksheet.mergeCells` takes them), given as ONE array:
     * `decode(["A1:B2"])`, `decode(["A1", "B2"])`, `decode([1, 1, 2, 2])`.
     */
    decode(args: RangeArguments): void;

    expand(top: number, left: number, bottom: number, right: number): void;

    expandRow(row: Row): void;

    expandToAddress(addressStr: string): void;

    toString(): string;

    intersects(other: Range): boolean;

    contains(addressStr: string): boolean;

    containsEx(
        address: Partial<{
            sheetName: string;
            row: number;
            col: number;
        }>
    ): boolean;
}

/** The header and footer of a worksheet: a {@link HeaderFooter} whose fields may also be `null` (the defaults). */
export type WorksheetHeaderFooter = { [K in keyof HeaderFooter]?: HeaderFooter[K] | null };

export interface RowBreak {
    id: number;
    max: number;
    min: number;
    man: number;
}

export interface WorksheetModel {
    id: number;
    name: string;
    // dataValidations: this.dataValidations.model,
    properties: Partial<WorksheetProperties>;
    pageSetup: Partial<PageSetup>;
    /** Headers and footers of the printed pages (`null` texts by default). */
    headerFooter: WorksheetHeaderFooter;
    rowBreaks: RowBreak[];
    views: WorksheetView[];
    autoFilter: AutoFilter;
    media: Media[];
    merges: Range["range"][];
}
export type WorksheetState = "visible" | "hidden" | "veryHidden";

export type CellIsOperators =
    "equal" | "notEqual" | "greaterThan" | "greaterThanOrEqual" | "lessThan" | "lessThanOrEqual" | "between" | "notBetween";

export type ContainsTextOperators =
    | "containsText"
    | "notContainsText"
    | "beginsWith"
    | "endsWith"
    | "containsBlanks"
    | "notContainsBlanks"
    | "containsErrors"
    | "notContainsErrors";

export type TimePeriodTypes =
    "lastWeek" | "thisWeek" | "nextWeek" | "yesterday" | "today" | "tomorrow" | "last7Days" | "lastMonth" | "thisMonth" | "nextMonth";

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

export type CfvoTypes = "percentile" | "percent" | "num" | "min" | "max" | "formula" | "autoMin" | "autoMax";

export interface Cvfo {
    type: CfvoTypes;
    /**
     * A number, or a formula or cell reference such as "$A$1" (type formula, or num / percent / percentile)
     */
    value?: number | string;
}
export interface ConditionalFormattingBaseRule {
    /**
     * Rules without a priority are given one (after the highest existing priority) when written
     */
    priority?: number;
    style?: Partial<Style>;
    /**
     * When true, lower-priority rules are not evaluated for a cell this rule applies to
     */
    stopIfTrue?: boolean;
}
export interface ExpressionRuleType extends ConditionalFormattingBaseRule {
    type: "expression";
    formulae?: any[];
    /**
     * Write the rule in the x14 extension of the worksheet, as Excel does for formulae that reference
     * other sheets (always done when a formula contains a sheet reference such as Other!A1).
     * Set on the rules read from that extension only.
     */
    x14?: boolean;
}

export interface CellIsRuleType extends ConditionalFormattingBaseRule {
    type: "cellIs";
    formulae?: any[];
    operator?: CellIsOperators;
    /**
     * Write the rule in the x14 extension of the worksheet, as Excel does for formulae that reference
     * other sheets (always done when a formula contains a sheet reference such as Other!A1).
     * Set on the rules read from that extension only.
     */
    x14?: boolean;
}

export interface Top10RuleType extends ConditionalFormattingBaseRule {
    type: "top10";
    rank: number;
    percent: boolean;
    bottom: boolean;
}

export interface AboveAverageRuleType extends ConditionalFormattingBaseRule {
    type: "aboveAverage";
    aboveAverage: boolean;
}

export interface ColorScaleRuleType extends ConditionalFormattingBaseRule {
    type: "colorScale";
    cfvo?: Cvfo[];
    color?: Partial<Color>[];
}

export interface IconSetRuleType extends ConditionalFormattingBaseRule {
    type: "iconSet";
    showValue?: boolean;
    reverse?: boolean;
    custom?: boolean;
    iconSet?: IconSetTypes;
    cfvo?: Cvfo[];
}

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
    formulae?: any[];
    /**
     * Write the rule in the x14 extension of the worksheet, as Excel does for formulae that reference
     * other sheets (always done when a formula contains a sheet reference such as Other!A1).
     * Set on the rules read from that extension only.
     */
    x14?: boolean;
}

export interface TimePeriodRuleType extends ConditionalFormattingBaseRule {
    type: "timePeriod";
    timePeriod?: TimePeriodTypes;
    /**
     * Written instead of the formula derived from the time period
     */
    formulae?: any[];
    /**
     * Write the rule in the x14 extension of the worksheet (set on the rules read from that extension only)
     */
    x14?: boolean;
}

export interface DataBarRuleType extends ConditionalFormattingBaseRule {
    type: "dataBar";
    gradient?: boolean;
    minLength?: number;
    maxLength?: number;
    showValue?: boolean;
    border?: boolean;
    negativeBarColorSameAsPositive?: boolean;
    negativeBarBorderColorSameAsPositive?: boolean;
    axisPosition?: "auto" | "middle" | "none";
    direction?: "context" | "leftToRight" | "rightToLeft";
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

export type RowValues = CellValue[] | { [key: string]: CellValue } | undefined | null;

export interface ConditionalFormattingOptions {
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
     * Worksheet Header and Footer
     */
    /** Headers and footers of the printed pages (`null` by default). */
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
     * Open panes representing the sheet
     */
    views: WorksheetViewModel[];

    /**
     * Apply an auto filter to your worksheet.
     */
    /** The auto filter, `null` or `undefined` when there is none. */
    autoFilter?: AutoFilter | null;

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
     * Access an individual columns by key, letter and 1-based column number
     */
    getColumn(indexOrKey: number | string): Column;

    /**
     * Cut one or more columns (columns to the right are shifted left)
     * and optionally insert more
     *
     * If column properties have been defined, they will be cut or moved accordingly
     *
     * Known Issue: If a splice causes any merged cells to move, the results may be unpredictable
     *
     * Also: If the worksheet has more rows than values in the colulmn inserts,
     * the rows will still be shifted as if the values existed
     */
    spliceColumns(start: number, count: number, ...insert: CellValue[][]): void;

    /**
     * Add column headers and define column keys and widths.
     *
     * Note: these column structures are a workbook-building convenience only,
     * apart from the column width, they will not be fully persisted.
     */
    get columns(): Column[] | null;
    set columns(value: ColumnDefinition[]);

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
     * Cut one or more rows (rows below are shifted up)
     * and optionally insert more
     *
     * Known Issue: If a splice causes any merged cells to move, the results may be unpredictable
     */
    spliceRows(start: number, count: number, ...insert: RowValues[]): void;

    /**
     * Add a row by key-value, after the last current row, using the column keys,
     * or add a row by contiguous Array (assign to columns A, B & C)
     */
    addRow(data: RowValues, style?: RowStyleOption): Row;

    /**
     * Add multiple rows by providing an array of arrays or key-value pairs
     */
    addRows(rows: RowValues[], style?: RowStyleOption): Row[];

    /**
     * Insert a Row by key-value, at the position (shifiting down all rows from position),
     * using the column keys, or add a row by contiguous Array (assign to columns A, B & C)
     */
    insertRow(pos: number, value: RowValues, style?: RowStyleOption): Row;

    /**
     * Insert multiple rows at position (shifiting down all rows from position)
     * by providing an array of arrays or key-value pairs
     */
    insertRows(pos: number, values: RowValues[], style?: RowStyleOption): Row[] | undefined;

    /**
     * Duplicate rows and insert new rows
     */
    duplicateRow(rowNum: number, count: number, insert: boolean): void;

    /**
     * Get or create row by 1-based index
     */
    getRow(index: number): Row;

    /**
     * Get or create rows by 1-based index
     */
    getRows(start: number, length: number): Row[] | undefined;

    /**
     * Iterate over all rows that have values in a worksheet
     */
    eachRow(callback: (row: Row, rowNumber: number) => void): void;

    /**
     * Iterate over all rows (including empty rows) in a worksheet
     */
    eachRow(opt: { includeEmpty: boolean }, callback: (row: Row, rowNumber: number) => void): void;

    /**
     * return all rows as sparse array
     */
    /** The values of the rows (sparse: indexed by row number, then by column number; index 0 unused). */
    getSheetValues(): CellValue[][];

    /**
     * returns the cell at [r,c] or address given by r. If not found, return undefined
     */
    findCell(r: number | string, c: number | string): Cell | undefined;

    /**
     * Get or create cell
     */
    getCell(r: number | string, c?: number | string): Cell;

    /**
     * Merge cells, either:
     *
     * tlbr string, e.g. `'A4:B5'`
     *
     * tl string, br string, e.g. `'G10', 'H11'`
     *
     * t, l, b, r numbers, e.g. `10,11,12,13`
     */
    mergeCells(): void;
    mergeCells(v: Range): void;
    mergeCells(v: string): void;
    mergeCells(v: Location): void;
    mergeCells(top: number, left: number, bottom: number, right: number, sheetName?: string): void;
    mergeCells(tl: string, br: string, sheetName?: string): void;
    mergeCells(v: [string, string]): void;
    mergeCells(v: [string, string, string]): void;
    mergeCells(v: [number, number, number, number]): void;
    mergeCells(v: [number, number, number, number, string]): void;
    mergeCellsWithoutStyle(): void;
    mergeCellsWithoutStyle(v: Range): void;
    mergeCellsWithoutStyle(v: string): void;
    mergeCellsWithoutStyle(v: Location): void;
    mergeCellsWithoutStyle(top: number, left: number, bottom: number, right: number, sheetName?: string): void;
    mergeCellsWithoutStyle(tl: string, br: string, sheetName?: string): void;
    mergeCellsWithoutStyle(v: [string, string]): void;
    mergeCellsWithoutStyle(v: [string, string, string]): void;
    mergeCellsWithoutStyle(v: [number, number, number, number]): void;
    mergeCellsWithoutStyle(v: [number, number, number, number, string]): void;

    /**
     * unmerging the cells breaks the style links
     */
    unMergeCells(): void;
    unMergeCells(v: Range): void;
    unMergeCells(v: string): void;
    unMergeCells(v: Location): void;
    unMergeCells(top: number, left: number, bottom: number, right: number, sheetName?: string): void;
    unMergeCells(tl: string, br: string, sheetName?: string): void;
    unMergeCells(v: [string, string]): void;
    unMergeCells(v: [string, string, string]): void;
    unMergeCells(v: [number, number, number, number]): void;
    unMergeCells(v: [number, number, number, number, string]): void;

    fillFormula(
        range: Range | string | Location,
        formula: string,
        results?: ((r: number, c: number) => string | number) | number[] | number[][]
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

    model: WorksheetModel;

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
     * delete conditionalFormattingOptions
     */
    removeConditionalFormatting(filter: any): void;
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
     * Color of the tab
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
     * The outline properties which controls how it will summarize rows and columns
     */
    outlineProperties: {
        summaryBelow: boolean;
        summaryRight: boolean;
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
    showGridLines: boolean;
}

export interface AddWorksheetOptions {
    properties: Partial<WorksheetProperties>;
    pageSetup: Partial<PageSetup>;
    headerFooter: Partial<HeaderFooter>;
    views: Array<Partial<WorksheetView>>;
    state: WorksheetState;
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
    /** The data is the base64 text of the file (a string, or the bytes of that text). */
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
    /** Writes a chunk of the file; returning `false` (with `on`) waits for `"drain"`. */
    write(chunk: Buffer): unknown;
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
     * Loads a file into the workbook from its base64 text (`base64: true`).
     * @returns the workbook, once read
     * @throws (rejects) as the other form
     */
    load(data: string | Uint8Array, options: Partial<XlsxLoadOptions> & { base64: true }): Promise<T>;

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
     * Writes the workbook as an xlsx file to an output; resolves once the archive is written (the output is not ended:
     * call `end()` afterwards).
     * @throws (rejects) for an image extension that is not 1 to 5 letters or digits (before anything is written), an
     * image without data, or a failure of the output
     */
    write(stream: XlsxOutput | NodeJS.WritableStream, options?: Partial<XlsxWriteOptions>): Promise<void>;
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
 * removed; a range of more than 100,000 cells is kept as a whole (removing any of its cells removes the range).
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
     * Removes a name from a cell, given as a reference (`"Sheet1!$A$1"`). A range reference removes nothing.
     * @throws Error when a column or a row is out of bounds
     */
    remove(locStr: string, name: string): void;
    /** Removes a name from a cell (and the range of more than 100,000 cells of the name that contains it). */
    removeEx(location: DefinedNameCell, name: string): void;
    /** Removes every name of a cell. */
    removeAllNames(location: DefinedNameCell): void;

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
    views: WorkbookView[] | undefined;
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
    creator: string | undefined;
    /** Document property: the description (`""` for a new workbook; `undefined` after loading a file without it). */
    description: string | undefined;
    /** Document property: the keywords (`""` for a new workbook; `undefined` after loading a file without it). */
    keywords: string | undefined;
    /** Document property: the author of the last change (written as `"Unknown"` when unset). */
    lastModifiedBy: string | undefined;
    /** Document property: the creation date (now for a new workbook; `undefined` after loading a file without it). */
    created: Date | undefined;
    /** Document property: the manager (`""` for a new workbook; `undefined` after loading a file without it). */
    manager: string | undefined;
    /** Document property: the date of the last change (now for a new workbook; `undefined` after loading a file without it). */
    modified: Date | undefined;
    /** Document property: the date of the last print. */
    lastPrinted: Date | undefined;
    /** Document property: the subject (`""` for a new workbook; `undefined` after loading a file without it). */
    subject: string | undefined;
    /** Document property: the title (`""` for a new workbook; `undefined` after loading a file without it). */
    title: string | undefined;
    /** Document property: the language. */
    language: string | undefined;
    /** Document property: the revision number (`NaN` after loading a file whose revision is not a number). */
    revision: number | undefined;
    /** Document property: the content status (`"Draft"`, `"Final"`...). */
    contentStatus: string | undefined;
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

    /**
     * The model of the workbook (the worksheet models and the settings); setting it replaces the workbook.
     * @throws Error (setter) for an invalid worksheet name, or when the merges exceed `maxMergedCells`
     */
    model: WorkbookModel;

    /**
     * The Workbook views controls how many separate windows Excel will open when viewing the workbook. `[]` for a new
     * workbook; `undefined` after loading a file without `<bookViews>` (such as one written with no views).
     */
    views: WorkbookView[] | undefined;

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
    /** @param model - the model to write (`undefined` before reading a file into the container) */
    constructor(model: WorkbookModel);
    /**
     * The workbook model: what `xlsx` writes, and what it sets when reading (then the parsed model, whose worksheets
     * hold their parsed rows and cells).
     */
    model: WorkbookModel;
    /** The xlsx reader and writer of the model (its read methods resolve with the container). */
    readonly xlsx: Xlsx<ModelContainer>;
}

export interface TableStyleProperties {
    /**
     * The colour theme of the table, or null for no table style
     * @default 'TableStyleMedium2'
     */
    theme?:
        | null
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
     * Styles applied to the column
     */
    style?: Partial<Style>;
}

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
     * Rows of data
     */
    rows: any[][];
}

/**
 * The style of a table column as the table holds it: a differential format (the style parts it sets; `null` for a part a
 * loaded file does not set, `numFmtId` once registered).
 */
export type TableColumnStyle = { [K in keyof Style]?: Partial<Style>[K] | null } & { numFmtId?: number };

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
 * worksheet: change them with the methods below, then `commit()`.
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
    /** The table style (`"TableStyleMedium2"`), or `null` for none. */
    theme: TableStyleProperties["theme"];
    showFirstColumn: boolean | undefined;
    showLastColumn: boolean | undefined;
    showRowStripes: boolean | undefined;
    showColumnStripes: boolean | undefined;
    /** The number of columns. */
    readonly width: number;
    /** The number of data rows. */
    readonly height: number;

    /**
     * Commit changes (writes the table to the cells of the worksheet)
     */
    commit(): void;
    /**
     * Remove rows of data (0-based index; 1 row by default)
     */
    removeRows(rowIndex: number, count?: number): void;
    /**
     * Add a row of data, either insert at rowNumber (0-based) or append
     */
    addRow(values: CellValue[], rowNumber?: number): void;
    /**
     * Get column (0-based index)
     */
    getColumn(colIndex: number): TableColumn;
    /**
     * Add a new column, including column defn and values; inserts at colIndex (0-based) or adds to the right
     */
    addColumn(column: TableColumnProperties, values: CellValue[], colIndex?: number): void;
    /**
     * Remove columns with their data (0-based index; 1 column by default)
     */
    removeColumns(colIndex: number, count?: number): void;
}

export namespace stream {
    namespace xlsx {
        interface WorkbookWriterOptions {
            /**
             * Specifies a writable stream to write the XLSX workbook to.
             */
            stream: import("stream").Stream;

            /**
             * 	If stream not specified, this field specifies the path to a file to write the XLSX workbook to.
             */
            filename: string;

            /**
             * 	Specifies whether to use shared strings in the workbook. Default is false
             */
            useSharedStrings: boolean;

            /**
             * Specifies whether to add style information to the workbook.
             * Styles can add some performance overhead. Default is false
             */
            useStyles: boolean;
        }

        /**
         * @deprecated use ZipOptions
         */
        type ArchiverZipOptions = ZipOptions;

        interface ZlibOptions {
            /**
             * @default constants.Z_NO_FLUSH
             */
            flush: number;
            /**
             * @default constants.Z_FINISH
             */
            finishFlush: number;
            /**
             * @default 16*1024
             */
            chunkSize: number;
            windowBits: number;
            level: number; // compression only
            memLevel: number; // compression only
            strategy: number; // compression only
            dictionary: Buffer | NodeJS.TypedArray | DataView | ArrayBuffer; // deflate/inflate only, empty dictionary by default
        }

        interface WorkbookStreamWriterOptions extends WorkbookWriterOptions {
            /**
             * Specifies whether to add style information to the workbook.
             * Styles can add some performance overhead. Default is false
             */
            zip: Partial<ZipOptions>;
            /**
             * Maximum number of cells covered by merges in all the worksheets (one worksheet holds at most 1,048,576).
             * Reading the file back with `xlsx.load` needs a `maxMergedCells` load option at least as large.
             * @default 1048576 (one full worksheet)
             */
            maxMergedCells: number;
        }

        class WorkbookWriter extends Workbook {
            constructor(options: Partial<WorkbookStreamWriterOptions>);
            // commit all worksheets, then add suplimentary files
            commit(): Promise<void>;
            addStyles(): Promise<void>;
            addThemes(): Promise<void>;
            addOfficeRels(): Promise<void>;
            addContentTypes(): Promise<void>;
            addApp(): Promise<void>;
            addCore(): Promise<void>;
            addSharedStrings(): Promise<void>;
            addWorkbookRels(): Promise<void>;
            addWorkbook(): Promise<void>;
        }

        interface WorkbookStreamReaderOptions extends Partial<ZipReadLimits> {
            /**
             * @default 'emit'
             */
            worksheets?: "emit" | "ignore";
            /**
             * @default 'cache'
             */
            sharedStrings?: "cache" | "emit" | "ignore";
            /**
             * @default 'ignore'
             */
            hyperlinks?: "cache" | "emit" | "ignore";
            /**
             * @default 'ignore'
             */
            styles?: "cache" | "ignore";
            /**
             * @default 'ignore'
             */
            entries?: "emit" | "ignore";
        }

        /**
         * A file name, a Node.js stream (or any object with a `pipe` method), a web ReadableStream or an
         * async iterable of byte chunks
         */
        type WorkbookReaderInput =
            string | import("stream").Stream | import("stream/web").ReadableStream<Uint8Array> | AsyncIterable<Uint8Array>;

        class WorkbookReader extends Workbook {
            constructor(input: WorkbookReaderInput, options?: Partial<WorkbookStreamReaderOptions>);
            read(input?: WorkbookReaderInput, options?: Partial<WorkbookStreamReaderOptions>): Promise<void>;
            [Symbol.asyncIterator](): AsyncGenerator<WorksheetReader>;
            parse(input?: WorkbookReaderInput, options?: Partial<WorkbookStreamReaderOptions>): AsyncIterator<any>;
        }

        /**
         * At runtime the WorkbookReader is an EventEmitter ('worksheet', 'shared-strings',
         * 'hyperlinks', 'entry', 'end', 'finished', 'error')
         */
        interface WorkbookReader extends EventEmitter {}

        interface WorksheetReaderOptions {
            workbook: WorkbookReader;
            id: number | string;
            iterator: AsyncIterable<any>;
            options: WorkbookStreamReaderOptions;
        }

        class WorksheetReader extends EventEmitter {
            constructor(options: WorksheetReaderOptions);
            readonly workbook: WorkbookReader;
            /**
             * The sheetId from the workbook, or the sheet file number (a string) if the
             * worksheet could not be matched to the workbook
             */
            id: number | string;
            /**
             * The worksheet name (`Sheet<id>` if the worksheet could not be matched to the workbook)
             */
            name: string;
            state?: WorksheetState;
            /**
             * The dimensions of the rows read so far
             */
            readonly dimensions: Range;
            /**
             * The column definitions, or null if none have been read (or created by getColumn)
             */
            readonly columns: Column[] | null;
            read(): Promise<void>;
            [Symbol.asyncIterator](): AsyncGenerator<Row>;
            parse(): AsyncIterator<Array<any>>;
            getColumn(c: number | string): Column;
        }
    }
}
