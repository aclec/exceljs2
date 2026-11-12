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

export interface WorkbookView {
    x: number;
    y: number;
    width: number;
    height: number;
    firstSheet: number;
    activeTab: number;
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
export interface Image {
    extension: "jpeg" | "png" | "gif";
    base64?: string;
    filename?: string;
    buffer?: ArrayBuffer | globalThis.Buffer;
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

export interface XlsxReadOptions extends ZipReadLimits {
    /**
     * The list of XML node names to ignore while parsing an XLSX file
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

export interface XlsxWriteOptions extends stream.xlsx.WorkbookWriterOptions {
    /**
     * Zip options
     */
    zip: Partial<ZipOptions>;
}

export interface Xlsx {
    /**
     * read from a file
     */
    readFile(path: string, options?: Partial<XlsxReadOptions>): Promise<Workbook>;

    /**
     * read from a stream
     * @param stream
     */
    read(stream: import("stream").Stream, options?: Partial<XlsxReadOptions>): Promise<Workbook>;

    /**
     * load from a Buffer, an ArrayBuffer or a Uint8Array
     * @param buffer
     */
    load(buffer: ArrayBuffer | Uint8Array, options?: Partial<XlsxReadOptions>): Promise<Workbook>;

    /**
     * write to a buffer
     */
    writeBuffer(options?: Partial<XlsxWriteOptions>): Promise<Buffer>;

    /**
     * write to a file
     */
    writeFile(path: string, options?: Partial<XlsxWriteOptions>): Promise<void>;

    /**
     * write to a stream
     */
    write(stream: import("stream").Stream, options?: Partial<XlsxWriteOptions>): Promise<void>;
}

// https://c2fo.github.io/fast-csv/docs/parsing/options

type HeaderArray = (string | undefined | null)[];
type HeaderTransformFunction = (headers: HeaderArray) => HeaderArray;
export interface FastCsvParserOptionsArgs {
    objectMode: boolean;
    delimiter: string;
    quote: string | null;
    escape: string;
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

interface QuoteColumnMap {
    [s: string]: boolean;
}
declare type QuoteColumns = boolean | boolean[] | QuoteColumnMap;

interface RowMap {
    [key: string]: any;
}
declare type RowHashArray = [string, any][];
declare type RowArray = string[];
declare type Rows = RowArray | RowMap | RowHashArray;
declare type RowTransformCallback = (error?: Error | null, row?: Rows) => void;
interface RowTransformFunction {
    (row: Rows, callback: RowTransformCallback): void;
    (row: Rows): Rows;
}

// https://c2fo.github.io/fast-csv/docs/formatting/options/
export interface FastCsvFormatterOptionsArgs {
    objectMode: boolean;
    delimiter: string;
    rowDelimiter: string;
    quote: string | boolean;
    escape: string;
    quoteColumns: QuoteColumns;
    quoteHeaders: QuoteColumns;
    headers: null | boolean | string[];
    includeEndRowDelimiter: boolean;
    writeBOM: boolean;
    transform: RowTransformFunction;
    alwaysWriteHeaders: boolean;
}

export interface CsvReadOptions {
    dateFormats: string[];
    map(value: any, index: number): any;
    sheetName: string;
    parserOptions: Partial<FastCsvParserOptionsArgs>;
}

export interface CsvWriteOptions {
    dateFormat: string;
    dateUTC: boolean;
    sheetName: string;
    sheetId: number;
    encoding: string;
    map(value: any, index: number): any;
    includeEmptyRows: boolean;
    formatterOptions: Partial<FastCsvFormatterOptionsArgs>;
}

export interface Csv {
    /**
     * read from a file
     */
    readFile(path: string, options?: Partial<CsvReadOptions>): Promise<Worksheet>;

    /**
     * read from a stream
     */
    read(stream: import("stream").Stream, options?: Partial<CsvReadOptions>): Promise<Worksheet>;

    /**
     * write to a buffer
     */
    writeBuffer(options?: Partial<CsvWriteOptions>): Promise<Buffer>;

    /**
     * write to a file
     */
    writeFile(path: string, options?: Partial<CsvWriteOptions>): Promise<void>;

    /**
     * write to a stream
     */
    write(stream: import("stream").Stream, options?: Partial<CsvWriteOptions>): Promise<void>;
}

export interface Media {
    type: string; // image,background
    name: string;
    extension: string;
    buffer: Buffer;
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

export interface DefinedNamesRanges {
    name: string;
    ranges: string[];
    /** The 0-based index of the worksheet a name is local to (absent: a workbook-level name). */
    localSheetId?: number;
}

export type DefinedNamesModel = DefinedNamesRanges[];

/** A cell of a defined name: its worksheet and address (`row` / `col` 1-based, `address` without `$`). */
export interface DefinedNameCell {
    sheetName?: string;
    address: string;
    row: number;
    col: number;
    $col$row?: string;
}

/** The defined names of a workbook (`workbook.definedNames`). */
export interface DefinedNames {
    /** Adds a name to a cell or a range, given as `"Sheet1!$A$1"` or `"Sheet1!$A$1:$B$2"`. */
    add(locStr: string, name: string): void;
    /** Adds a name to a cell, or to the cells of a range given by its bounds (`top`, `left`, `bottom`, `right`). */
    addEx(location: DefinedNameCell | (Location & { sheetName?: string }), name: string): void;

    /** Removes a name from a cell or a range, given as a reference (`"Sheet1!$A$1"`). */
    remove(locStr: string, name: string): void;
    removeEx(location: DefinedNameCell | (Location & { sheetName?: string }), name: string): void;
    /** Removes every name of a cell. */
    removeAllNames(location: DefinedNameCell): void;

    /** Calls `callback` with each name and each of its cells (the cells of the very large ranges are not enumerated). */
    forEach(callback: (name: string, cell: DefinedNameCell) => void): void;

    /** The names of a cell, given as a reference (`"Sheet1!$A$1"`). */
    getNames(addressStr: string): string[];
    /** The names of a cell. */
    getNamesEx(address: DefinedNameCell): string[];

    /** The ranges of a name, as references (`"Sheet1!$A$1:$B$2"`). */
    getRanges(name: string): DefinedNamesRanges;

    model: DefinedNamesModel;
}

export interface WorkbookModel {
    creator: string;
    lastModifiedBy: string;
    lastPrinted: Date;
    created: Date;
    modified: Date;
    properties: WorkbookProperties;
    worksheets: Worksheet[];
    sheets: WorksheetModel[];
    definedNames: DefinedNamesModel;
    views: WorkbookView[];
    company: string;
    manager: string;
    title: string;
    subject: string;
    keywords: string;
    category: string;
    description: string;
    language: string;
    revision: Date;
    contentStatus: string;
    themes: string[];
    media: Media[];
}

export class Workbook {
    category: string;
    company: string;
    creator: string;
    description: string;
    keywords: string;
    lastModifiedBy: string;
    created: Date;
    manager: string;
    modified: Date;
    lastPrinted: Date;
    properties: WorkbookProperties;
    subject: string;
    title: string;

    /**
     * Workbook calculation Properties
     */
    calcProperties: CalculationProperties;

    /**
     * xlsx file format operations
     */
    readonly xlsx: Xlsx;

    /**
     * csv file format operations
     */
    readonly csv: Csv;

    readonly nextId: number;

    readonly definedNames: DefinedNames;

    /**
     * Maximum number of cells covered by merges in all the worksheets (`mergeCells` throws beyond it; one worksheet holds
     * at most 1,048,576). Also the default `maxMergedCells` of `xlsx.load` / `read` / `readFile`.
     * @default 1048576 (one full worksheet)
     */
    maxMergedCells: number;

    model: WorkbookModel;

    /**
     * The Workbook views controls how many separate windows Excel will open when viewing the workbook.
     */
    views: WorkbookView[];

    /**
     * return a clone of worksheets in order
     */
    worksheets: Worksheet[];

    /**
     * Add a new worksheet and return a reference to it
     */
    addWorksheet(name?: string, options?: Partial<AddWorksheetOptions>): Worksheet;

    removeWorksheetEx(worksheet: Worksheet): void;
    removeWorksheet(indexOrName: number | string): void;

    /**
     * fetch sheet by name or id
     */
    getWorksheet(indexOrName?: number | string): Worksheet | undefined;

    /**
     * Iterate over all sheets.
     *
     * Note: `workbook.worksheets.forEach` will still work but this is better.
     */
    eachSheet(callback: (worksheet: Worksheet, id: number) => void): void;

    clearThemes(): void;

    /**
     * Add Image to Workbook and return the id
     */
    addImage(img: Image): number;

    getImage(id: number): Image;
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
            readonly views: Partial<WorksheetView>[];

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
             * @throws Error when the row was committed
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

            /** The cell at an address (`"B3"`), if it exists and its row is not committed. */
            findCell(address: string): Cell | undefined;
            /** The cell at (row, column), if it exists and its row is not committed. */
            findCell(row: number, col: number): Cell | undefined;
            /**
             * The cell at an address (`"B3"`), created if needed.
             *
             * @throws Error when its row was committed
             */
            getCell(address: string): Cell;
            /**
             * The cell at (row, column), created if needed.
             *
             * @throws Error when its row was committed
             */
            getCell(row: number, col: number): Cell;

            /**
             * Merges a range of cells (`"A4:B5"`, a {@link Range} or {@link Location}, or an array of the arguments
             * below): every cell but the top-left one is merged into it.
             *
             * @throws Error when a cell is already merged, a row was committed, or the merged cells exceed the limits
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
             * The hyperlinks of the worksheets: `cache`: the `<hyperlink>` elements in `worksheetReader.hyperlinks` (the
             * targets of external links, in the relationship parts, are not read); `emit`: `hyperlink` events of the
             * worksheet readers, and `hyperlinks` events (a {@link HyperlinkReader} per worksheet, for the targets) of the
             * workbook reader; `ignore`.
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

        /** A `<hyperlink>` of a worksheet, as read (`hyperlinks: "cache"` or `"emit"`): the attributes present. */
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
             * `hyperlinks: "cache"`: the hyperlinks read so far, keyed by cell reference (an object without prototype).
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
         * workbook reader never hands out: they stay empty.
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
