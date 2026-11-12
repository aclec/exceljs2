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
 * also its empty cells), `"o"` the style of the row below (`"o+"`; `insertRow` / `insertRows` only, ignored by
 * `addRow`), `"n"` (the default) none.
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
     * of its cell, whatever its type). It must be truthy: a value with an empty or zero `text` is not a hyperlink (it is
     * held as a JSON value, see {@link CellValue}).
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
 * The kind of an object value is told by its truthy marker properties, in this order: `text` and `hyperlink`
 * (hyperlink), `formula` or `sharedFormula` (formula), `richText`, `sharedString`, `error`. Any other object (outside
 * this type: an untyped value, such as one returned by the CSV `map` option) is held as a JSON value: its type is
 * String, it is returned as is, and it is written as its JSON text (read back as a string).
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
     * Set a specific row height; `undefined` when the row has the default height
     */
    height: number | undefined;

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

/**
 * The values of a row: a sparse array indexed by column number (index 0 unused), a contiguous array (index 0 set: the
 * values of columns 1, 2...), or an object keyed by column key; `null` / `undefined`: no value.
 */
export type RowValues = readonly CellValue[] | { readonly [key: string]: CellValue } | undefined | null;

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
     */
    getRow(index: number): Row;

    /**
     * The rows `start`..`start + length - 1` (1-based), created if needed.
     *
     * @returns `undefined` when `length` < 1
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

    /** The cell at an address (`"B2"`, `"$B$2"`, `"Sheet1!B2"`) if it exists. */
    findCell(address: string): Cell | undefined;
    /**
     * The cell at (row, col) (1-based numbers) if it exists.
     *
     * @throws Error for a column out of bounds
     */
    findCell(row: number, col: number): Cell | undefined;

    /**
     * The cell at an address (`"B2"`, `"$B$2"`, `"Sheet1!B2"`), created if needed.
     *
     * @throws Error for an invalid address, or a row or a column out of bounds
     */
    getCell(address: string): Cell;
    /**
     * The cell at (row, col) (1-based numbers), created if needed.
     *
     * @throws Error for a row or a column out of bounds
     */
    getCell(row: number, col: number): Cell;

    /**
     * Merges a range into its top-left cell, which gives its style to the other cells. The range is given as for a
     * {@link Range} (see {@link RangeArguments}): `"A4:B5"`, `"G10", "H11"`, `10, 11, 12, 13` (top, left, bottom, right),
     * a {@link Location} or a Range.
     *
     * @throws Error when a cell is already merged, or the merges would exceed the cap
     */
    mergeCells(...cells: RangeArguments): void;
    /**
     * Merges a range into its top-left cell, the cells keeping their own style (see `mergeCells`).
     *
     * @throws Error when a cell is already merged, or the merges would exceed the cap
     */
    mergeCellsWithoutStyle(...cells: RangeArguments): void;

    /**
     * Unmerges the merges that have a cell in the range (given as for `mergeCells`): a merge is unmerged as a whole.
     * Unmerging breaks the style links.
     */
    unMergeCells(...cells: RangeArguments): void;

    /**
     * Sets a formula on a range (`"A1:B3"`): the top-left cell gets it (a shared formula, or an array formula with
     * `shareType` `"array"`), the other cells are its clones (shared formula) or get their result (array formula). A
     * single cell (`"A1"`), or a string that is not a range, fills nothing.
     *
     * @param results - the cached results of the cells: by cell (1-based row and column), by row then column of the
     *   range, or flat (row by row)
     * @param shareType - `"shared"` (the default) or `"array"`
     * @throws Error for a range out of bounds
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

/** The bounds of a rectangle of cells (1-based row and column numbers). */
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
