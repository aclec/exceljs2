// A cell of a worksheet, and the value types behind it (NullValue, NumberValue, StringValue, FormulaValue...).
//
// Cell requirements
//  Operate inside a worksheet
//  Store and retrieve a value with a range of types: text, number, date, hyperlink, reference, formula, etc.
//  Manage/use and manipulate cell format either as local to cell or inherited from column or row.

import type { CellErrorValue, Comment, Style } from "../../index";
import colCache from "../utils/col-cache";
import { escapeHtml } from "../utils/under-dash";
import { FormulaType, ValueType } from "./enums";
import { slideFormula, checkFormulaLength } from "../utils/shared-formula";
import Note from "./note";
import type DataValidations from "./data-validations";
import type DefinedNames from "./defined-names";
import type {
    BooleanCellModel,
    CellModel,
    DataValidationModel,
    DateCellModel,
    ErrorCellModel,
    FormulaCellModel,
    FormulaResult,
    HyperlinkCellModel,
    MergeCellModel,
    NullCellModel,
    NumberCellModel,
    RichTextValue as RichTextObject,
    SharedStringCellModel,
    StringCellModel,
    StyleModel,
} from "../types";

// =============================================================================
// Public types

/** The name of a style property of a cell, row or column (`cell.font = ...`). */
export type StyleName = keyof Style;

/** The value of a hyperlink cell: index.d.ts `CellHyperlinkValue` (a loaded cell's text may be any value). */
export interface HyperlinkValueObject {
    text: HyperlinkCellModel["text"];
    hyperlink: string;
    tooltip?: string | undefined;
}

/**
 * The value of a formula cell: a master formula (`formula`, and for a shared or array formula `shareType` and `ref`), or
 * a clone of a shared formula (`sharedFormula`: the address of its master); `result` is the cached result. The value
 * read from a cell has only the keys that are set.
 */
export interface FormulaValueObject {
    formula?: string | undefined;
    result?: FormulaResult | undefined;
    ref?: string | undefined;
    shareType?: "shared" | "array" | undefined;
    sharedFormula?: string | undefined;
}

/** A shared string reference value (`{ sharedString }`, legacy: held as is, type SharedString). */
export interface SharedStringRef {
    readonly sharedString: unknown;
}

/**
 * The value of a cell: index.d.ts `CellValue` (null, number, string, boolean, Date, error, rich text, hyperlink,
 * formula), a shared string, or any other object, held as a JSON value (its `type` is String, it is written as its JSON
 * text). Setting `undefined` empties the cell (the value is then `null`).
 *
 * The kind of value is told by its type, then by truthy marker properties, in this order: `text` and `hyperlink`
 * (hyperlink), `formula` or `sharedFormula` (formula), `richText`, `sharedString`, `error`.
 */
export type CellValue =
    | null
    | number
    | string
    | boolean
    | Date
    | CellErrorValue
    | RichTextObject
    | HyperlinkValueObject
    | FormulaValueObject
    | SharedStringRef
    | object;

/**
 * What a {@link Cell} reads of its worksheet: its name and workbook (defined names), its data validations, and the
 * other cells (the master of a merge, the master of a shared formula).
 */
export interface CellWorksheet {
    readonly name: string;
    readonly workbook: { readonly definedNames: DefinedNames };
    readonly dataValidations: DataValidations;
    /** The cell at an address, created if needed. */
    getCell(address: string): Cell;
    /** The cell at an address, if it exists. */
    findCell(address: string): Cell | undefined;
}

/** What a {@link Cell} reads of its row: the row number, the worksheet, and the row style (inherited by new cells). */
export interface CellRow {
    readonly number: number;
    readonly worksheet: CellWorksheet;
    readonly style?: StyleModel | undefined;
}

/** What a {@link Cell} reads of its column: the column number and letters, and the style (inherited by new cells). */
export interface CellColumn {
    readonly number: number;
    readonly letter: string;
    readonly style?: StyleModel | undefined;
}

// =============================================================================
// Cell

/**
 * A cell of a worksheet: a value (see {@link CellValue}), a style (own, or inherited from its row and column when it is
 * created), a note, names, a data validation. A merged cell (`type` Merge) shows and sets the value of its master.
 */
class Cell {
    /** The value types (`cell.type`): the ValueType enumeration. */
    static readonly Types = ValueType;

    /** The row. */
    _row: CellRow;
    /** The column. */
    _column: CellColumn;
    /** The address (`"B2"`). */
    _address: string;
    /** The value holder (one of the value classes below). */
    _value: AnyValue;
    /**
     * The style. Style objects may be shared (the cells read from a file share one object per style id, the cells of a
     * merge share the master's): setting a style property copies it (see `_applyStyle`).
     */
    style: StyleModel;
    /** The cells merged into this one (created on first merge). */
    _mergeSlaves: Set<Cell> | undefined;
    /** The note of the cell, if any. */
    declare _comment?: Note | undefined;

    /**
     * @param address - the address of the cell (`"B2"`)
     * @throws Error without a row or a column, or for an invalid address
     */
    constructor(row: CellRow | undefined, column: CellColumn | undefined, address: string) {
        if (!row || !column) {
            throw new Error("A Cell needs a Row");
        }

        this._row = row;
        this._column = column;

        colCache.validateAddress(address);
        this._address = address;

        // TODO: lazy evaluation of this._value
        this._value = new NullValue(this);

        this.style = this._mergeStyle(row.style, column.style, {});

        // cells merged into this one (created on first merge)
        this._mergeSlaves = undefined;
    }

    get worksheet(): CellWorksheet {
        return this._row.worksheet;
    }

    get workbook(): CellWorksheet["workbook"] {
        return this._row.worksheet.workbook;
    }

    /** Helps GC by removing cyclic (and other) references: the cell cannot be used afterwards. */
    destroy(): void {
        // the cell is unusable once destroyed: its fields are removed, not typed as optional
        const self = this as Partial<Pick<Cell, "style" | "_value" | "_row" | "_column" | "_address">>;
        delete self.style;
        delete self._value;
        delete self._row;
        delete self._column;
        delete self._address;
    }

    // =========================================================================
    // Styles stuff
    get numFmt(): StyleModel["numFmt"] {
        return this.style.numFmt;
    }

    set numFmt(value: StyleModel["numFmt"]) {
        this._applyStyle("numFmt", value);
    }

    get font(): StyleModel["font"] {
        return this.style.font;
    }

    set font(value: StyleModel["font"]) {
        this._applyStyle("font", value);
    }

    get alignment(): StyleModel["alignment"] {
        return this.style.alignment;
    }

    set alignment(value: StyleModel["alignment"]) {
        this._applyStyle("alignment", value);
    }

    get border(): StyleModel["border"] {
        return this.style.border;
    }

    set border(value: StyleModel["border"]) {
        this._applyStyle("border", value);
    }

    get fill(): StyleModel["fill"] {
        return this.style.fill;
    }

    set fill(value: StyleModel["fill"]) {
        this._applyStyle("fill", value);
    }

    get protection(): StyleModel["protection"] {
        return this.style.protection;
    }

    set protection(value: StyleModel["protection"]) {
        this._applyStyle("protection", value);
    }

    /**
     * Sets a style property (what the style setters do). Style objects may be shared: cells read from a file reference
     * one object per style id. So copy on write, but keep the cells of a merge sharing the master's style object
     * (README "Merged Cells").
     */
    _applyStyle<K extends StyleName>(name: K, value: StyleModel[K]): void {
        const previous = this.style;
        if (previous[name] === value) {
            // nothing changes; this also keeps row/column styling across a big merge linear
            return;
        }
        const style: StyleModel = { ...previous, [name]: value };
        const { master } = this;
        if (master._mergeSlaves && master.style === previous) {
            master.style = style;
            for (const slave of master._mergeSlaves) {
                if (slave.style === previous) slave.style = style;
            }
        }
        this.style = style;
    }

    /** Copies into `style` the style properties of the row, or else of the column (by reference). */
    _mergeStyle(rowStyle: StyleModel | undefined, colStyle: StyleModel | undefined, style: StyleModel): StyleModel {
        const numFmt = (rowStyle && rowStyle.numFmt) || (colStyle && colStyle.numFmt);
        if (numFmt) style.numFmt = numFmt;

        const font = (rowStyle && rowStyle.font) || (colStyle && colStyle.font);
        if (font) style.font = font;

        const alignment = (rowStyle && rowStyle.alignment) || (colStyle && colStyle.alignment);
        if (alignment) style.alignment = alignment;

        const border = (rowStyle && rowStyle.border) || (colStyle && colStyle.border);
        if (border) style.border = border;

        const fill = (rowStyle && rowStyle.fill) || (colStyle && colStyle.fill);
        if (fill) style.fill = fill;

        const protection = (rowStyle && rowStyle.protection) || (colStyle && colStyle.protection);
        if (protection) style.protection = protection;

        return style;
    }

    // =========================================================================
    /** The address of the cell (`"B2"`). */
    get address(): string {
        return this._address;
    }

    /** The 1-based row number. */
    get row(): number {
        return this._row.number;
    }

    /** The 1-based column number. */
    get col(): number {
        return this._column.number;
    }

    /** The absolute address (`"$B$2"`). */
    get $col$row(): string {
        return `$${this._column.letter}$${this.row}`;
    }

    // =========================================================================
    // Value stuff

    /** The type of the value (ValueType; a JSON value is String). */
    get type(): ValueType {
        return this._value.type;
    }

    /** The type of the value, or of the result of a formula, or of the master's value for a merged cell. */
    get effectiveType(): ValueType {
        return this._value.effectiveType;
    }

    /**
     * The value as CSV text (legacy, not used by the CSV writer): a quoted string, `1` / `0` for a boolean (a number),
     * the target of a hyperlink, the result of a formula, an ISO date...
     */
    toCsvString(): string | number | undefined {
        return this._value.toCsvString();
    }

    // =========================================================================
    // Merge stuff

    /** Records a cell merged into this one. */
    addMergeRef(slave: Cell): void {
        if (!this._mergeSlaves) {
            this._mergeSlaves = new Set();
        }
        this._mergeSlaves.add(slave);
    }

    /** Forgets a cell merged into this one (idempotent). */
    releaseMergeRef(slave: Cell): void {
        // idempotent: releasing a cell that no longer references this master is a no-op
        if (!this._mergeSlaves) {
            return;
        }
        this._mergeSlaves.delete(slave);
        if (!this._mergeSlaves.size) {
            this._mergeSlaves = undefined;
        }
    }

    /** Whether the cell is part of a merge (its master, or merged into another cell). */
    get isMerged(): boolean {
        return !!this._mergeSlaves || this.type === ValueType.Merge;
    }

    /** Merges this cell into `master`: its value is dropped, and it takes the master's style unless `ignoreStyle`. */
    merge(master: Cell, ignoreStyle?: boolean): void {
        this._value.release();
        this._value = new MergeValue(this, master);
        if (!ignoreStyle) {
            this.style = master.style;
        }
    }

    /** Unmerges a merged cell (it becomes empty; its style is reset from its row and column unless `ignoreStyle`). */
    unmerge(ignoreStyle?: boolean): void {
        if (this.type === ValueType.Merge) {
            this._value.release();
            this._value = new NullValue(this);
            if (!ignoreStyle) {
                this.style = this._mergeStyle(this._row.style, this._column.style, {});
            }
        }
    }

    /** Whether this cell is merged into `master`. */
    isMergedTo(master: Cell): boolean {
        if (!(this._value instanceof MergeValue)) return false;
        return this._value.isMergedTo(master);
    }

    /** The master of the merge this cell is merged into; an unmerged cell (or a master) is its own master. */
    get master(): Cell {
        if (this._value instanceof MergeValue) {
            return this._value.master;
        }
        return this; // an unmerged cell is its own master
    }

    /** Whether the value is a hyperlink. */
    get isHyperlink(): boolean {
        return this._value.type === ValueType.Hyperlink;
    }

    /** The target of a hyperlink value (undefined for another value). */
    get hyperlink(): string | undefined {
        return this._value instanceof HyperlinkValue ? this._value.hyperlink : undefined;
    }

    /** The value (see {@link CellValue}): a merged cell gives its master's value. */
    get value(): CellValue {
        return this._value.value;
    }

    /** Sets the value (see {@link CellValue}); a merged cell sets its master's value. */
    set value(v: CellValue | undefined) {
        // special case - merge cells set their master's value
        if (this._value instanceof MergeValue) {
            this._value.master.value = v;
            return;
        }

        this._value.release();

        // assign value
        this._value = createValue(getValueType(v), this, v);
    }

    /** The note: a string, or a comment. */
    get note(): string | Comment | undefined {
        return this._comment && this._comment.note;
    }

    set note(note: string | Comment) {
        this._comment = new Note(note);
    }

    /** The value as text: the text of a rich text, the result of a formula, a date's `toString()`... */
    get text(): string {
        return this._value.toString();
    }

    /** The text, HTML-escaped. */
    get html(): string {
        return escapeHtml(this.text);
    }

    toString(): string {
        return this.text;
    }

    /** Turns a string value into a hyperlink (the string is its text); does nothing for another value. */
    _upgradeToHyperlink(hyperlink: string): void {
        // if this cell is a string, turn it into a Hyperlink
        if (this._value.type === ValueType.String) {
            this._value = new HyperlinkValue(this, {
                // a JSON value (also of type String) gives its raw value
                text: this._value.value as HyperlinkValueObject["text"],
                hyperlink,
            });
        }
    }

    // =========================================================================
    // Formula stuff
    /** The formula (without `=`): a clone of a shared formula gives its translated formula. undefined for another value. */
    get formula(): string | undefined {
        return this._value instanceof FormulaValue ? this._value.formula : undefined;
    }

    /** The cached result of a formula (undefined for another value). */
    get result(): FormulaResult | undefined {
        return this._value instanceof FormulaValue ? this._value.result : undefined;
    }

    /** The kind of formula (Master, Shared, None); undefined for a value that is not a formula. */
    get formulaType(): FormulaType | undefined {
        return this._value instanceof FormulaValue ? this._value.formulaType : undefined;
    }

    // =========================================================================
    // Name stuff
    /** The location of the cell, with its sheet name (for the defined names). */
    get fullAddress(): { sheetName: string; address: string; row: number; col: number } {
        const { worksheet } = this._row;
        return {
            sheetName: worksheet.name,
            address: this.address,
            row: this.row,
            col: this.col,
        };
    }

    /** The first defined name of the cell; setting it replaces all the names of the cell. */
    get name(): string | undefined {
        return this.names[0];
    }

    set name(value: string) {
        this.names = [value];
    }

    /** The defined names that include the cell; setting them replaces them. */
    get names(): string[] {
        return this.workbook.definedNames.getNamesEx(this.fullAddress);
    }

    set names(value: readonly string[]) {
        const { definedNames } = this.workbook;
        definedNames.removeAllNames(this.fullAddress);
        value.forEach((name) => {
            definedNames.addEx(this.fullAddress, name);
        });
    }

    /** Adds the cell to a defined name. */
    addName(name: string): void {
        this.workbook.definedNames.addEx(this.fullAddress, name);
    }

    /** Removes the cell from a defined name (a large range kept as a rectangle is removed whole). */
    removeName(name: string): void {
        this.workbook.definedNames.removeEx(this.fullAddress, name);
    }

    /** Removes the cell from all the defined names. */
    removeAllNames(): void {
        this.workbook.definedNames.removeAllNames(this.fullAddress);
    }

    // =========================================================================
    // Data Validation stuff
    get _dataValidations(): DataValidations {
        return this.worksheet.dataValidations;
    }

    /** The data validation of the cell, from its own key or from a range that contains it. */
    get dataValidation(): DataValidationModel | undefined {
        return this._dataValidations.find(this.address);
    }

    set dataValidation(value: DataValidationModel) {
        this._dataValidations.add(this.address, value);
    }

    // =========================================================================
    // Model stuff

    /**
     * The model of the cell: the model of its value (the same object each time, not a copy) with `style` and `comment`
     * set on it.
     */
    get model(): CellModel {
        const { model } = this._value;
        model.style = this.style;
        if (this._comment) {
            model.comment = this._comment.model;
        }
        return model;
    }

    /**
     * Sets the cell from a model, adopted by reference. The model of a merged cell names its master by address (the
     * master is created if needed). The style is taken by reference (`{}` when there is none).
     * @throws Error for an unknown value type
     */
    set model(value: CellModel) {
        this._value.release();
        // the model of a merged cell names its master by address (see MergeValue)
        const master = value.type === ValueType.Merge && value.master ? this.worksheet.getCell(value.master) : undefined;
        const holder = createValue(value.type, this, master);
        // the holder was created for the type of the model, whose class takes it as its model
        (holder as { model: CellModel }).model = value;
        this._value = holder;

        if (value.comment) {
            switch (value.comment.type) {
                case "note":
                    this._comment = Note.fromModel(value.comment);
                    break;
            }
        }

        if (value.style) {
            this.style = value.style;
        } else {
            this.style = {};
        }
    }
}

// =============================================================================
// Internal Value Types
//
// Each holds the model of the cell value (`model`), created by its constructor. The Cell model setter creates the holder
// of the model's type, then replaces its model with the given one. The class names are visible (spec, debugging).

/** A cell without value. */
class NullValue {
    model: NullCellModel;

    constructor(cell: Cell) {
        this.model = {
            address: cell.address,
            type: ValueType.Null,
        };
    }

    get value(): null {
        return null;
    }

    set value(_value: unknown) {
        // nothing to do
    }

    get type(): typeof ValueType.Null {
        return ValueType.Null;
    }

    get effectiveType(): ValueType {
        return ValueType.Null;
    }

    get address(): string {
        return this.model.address;
    }

    set address(value: string) {
        this.model.address = value;
    }

    toCsvString(): string {
        return "";
    }

    release(): void {}

    toString(): string {
        return "";
    }
}

/** A number. */
class NumberValue {
    model: NumberCellModel;

    constructor(cell: Cell, value: number) {
        this.model = {
            address: cell.address,
            type: ValueType.Number,
            value,
        };
    }

    get value(): number {
        return this.model.value;
    }

    set value(value: number) {
        this.model.value = value;
    }

    get type(): typeof ValueType.Number {
        return ValueType.Number;
    }

    get effectiveType(): ValueType {
        return ValueType.Number;
    }

    get address(): string {
        return this.model.address;
    }

    set address(value: string) {
        this.model.address = value;
    }

    toCsvString(): string {
        return this.model.value.toString();
    }

    release(): void {}

    toString(): string {
        return this.model.value.toString();
    }
}

/** A string (a String object is held as it is). */
class StringValue {
    model: StringCellModel & { value: string };

    constructor(cell: Cell, value: string) {
        this.model = {
            address: cell.address,
            type: ValueType.String,
            value,
        };
    }

    get value(): string {
        return this.model.value;
    }

    set value(value: string) {
        this.model.value = value;
    }

    get type(): typeof ValueType.String {
        return ValueType.String;
    }

    get effectiveType(): ValueType {
        return ValueType.String;
    }

    get address(): string {
        return this.model.address;
    }

    set address(value: string) {
        this.model.address = value;
    }

    toCsvString(): string {
        return `"${this.model.value.replace(/"/g, '""')}"`;
    }

    release(): void {}

    toString(): string {
        return this.model.value;
    }
}

/**
 * A rich text (`{ richText: [{ text, font }] }`). NB: the model created from a value has the type String (a loaded one
 * has RichText); the cell's `type` is RichText either way.
 */
class RichTextValue {
    model: StringCellModel & { value: RichTextObject };

    constructor(cell: Cell, value: RichTextObject) {
        this.model = {
            address: cell.address,
            type: ValueType.String,
            value,
        };
    }

    get value(): RichTextObject {
        return this.model.value;
    }

    set value(value: RichTextObject) {
        this.model.value = value;
    }

    /** The text of the runs, joined. */
    toString(): string {
        return this.model.value.richText.map((t) => t.text).join("");
    }

    get type(): typeof ValueType.RichText {
        return ValueType.RichText;
    }

    get effectiveType(): ValueType {
        return ValueType.RichText;
    }

    get address(): string {
        return this.model.address;
    }

    set address(value: string) {
        this.model.address = value;
    }

    toCsvString(): string {
        return `"${this.toString().replace(/"/g, '""')}"`;
    }

    release(): void {}
}

/** A date. */
class DateValue {
    model: DateCellModel;

    constructor(cell: Cell, value: Date) {
        this.model = {
            address: cell.address,
            type: ValueType.Date,
            value,
        };
    }

    get value(): Date {
        return this.model.value;
    }

    set value(value: Date) {
        this.model.value = value;
    }

    get type(): typeof ValueType.Date {
        return ValueType.Date;
    }

    get effectiveType(): ValueType {
        return ValueType.Date;
    }

    get address(): string {
        return this.model.address;
    }

    set address(value: string) {
        this.model.address = value;
    }

    toCsvString(): string {
        return this.model.value.toISOString();
    }

    release(): void {}

    toString(): string {
        return this.model.value.toString();
    }
}

/** A hyperlink: a text, a target, and an optional tooltip. */
class HyperlinkValue {
    model: HyperlinkCellModel;

    /** @param value - nothing from the Cell model setter, which replaces the model right after */
    constructor(cell: Cell, value?: HyperlinkValueObject) {
        this.model = {
            address: cell.address,
            type: ValueType.Hyperlink,
            // without a value, the placeholder model has neither text nor hyperlink
            text: (value ? value.text : undefined) as HyperlinkCellModel["text"],
            hyperlink: (value ? value.hyperlink : undefined) as string,
        };
        if (value && value.tooltip) {
            this.model.tooltip = value.tooltip;
        }
    }

    /** `{ text, hyperlink, tooltip? }` (a new object). */
    get value(): HyperlinkValueObject {
        const v: HyperlinkValueObject = {
            text: this.model.text,
            hyperlink: this.model.hyperlink,
        };
        if (this.model.tooltip) {
            v.tooltip = this.model.tooltip;
        }
        return v;
    }

    /** NB: replaces the model with one without address and type (unused by the library). */
    set value(value: HyperlinkValueObject) {
        this.model = {
            text: value.text,
            hyperlink: value.hyperlink,
        } as HyperlinkCellModel;
        if (value.tooltip) {
            this.model.tooltip = value.tooltip;
        }
    }

    get text(): HyperlinkCellModel["text"] {
        return this.model.text;
    }

    set text(value: HyperlinkCellModel["text"]) {
        this.model.text = value;
    }

    get hyperlink(): string {
        return this.model.hyperlink;
    }

    set hyperlink(value: string) {
        this.model.hyperlink = value;
    }

    get type(): typeof ValueType.Hyperlink {
        return ValueType.Hyperlink;
    }

    get effectiveType(): ValueType {
        return ValueType.Hyperlink;
    }

    get address(): string {
        return this.model.address;
    }

    set address(value: string) {
        this.model.address = value;
    }

    toCsvString(): string {
        return this.model.hyperlink;
    }

    release(): void {}

    /** The text (as it is: a loaded hyperlink's text may be another value than a string). */
    toString(): string {
        return this.model.text as string;
    }
}

/** A cell merged into another one (its master): it shows and sets the master's value. */
class MergeValue {
    model: MergeCellModel;
    /** The merged cell. */
    _cell: Cell;
    /** The master of the merge. */
    _master: Cell;

    constructor(cell: Cell, master: Cell) {
        this.model = {
            address: cell.address,
            type: ValueType.Merge,
            master: master ? master.address : undefined,
        };
        this._cell = cell;
        this._master = master;
        if (master) {
            master.addMergeRef(cell);
        }
    }

    /** The master's value. */
    get value(): CellValue {
        return this._master.value;
    }

    /** Sets the master's value; a Cell given as the value becomes the new master. */
    set value(value: CellValue | Cell | undefined) {
        if (value instanceof Cell) {
            if (this._master) {
                this._master.releaseMergeRef(this._cell);
            }
            value.addMergeRef(this._cell);
            this._master = value;
        } else {
            this._master.value = value;
        }
    }

    isMergedTo(master: Cell): boolean {
        return master === this._master;
    }

    get master(): Cell {
        return this._master;
    }

    get type(): typeof ValueType.Merge {
        return ValueType.Merge;
    }

    get effectiveType(): ValueType {
        return this._master.effectiveType;
    }

    get address(): string {
        return this.model.address;
    }

    set address(value: string) {
        this.model.address = value;
    }

    toCsvString(): string {
        return "";
    }

    release(): void {
        this._master.releaseMergeRef(this._cell);
    }

    /** The text of the master. */
    toString(): string {
        return this._master.text;
    }
}

/** The text of a formula result: the code of an error, the `toString()` of another value. */
function resultText(result: FormulaResult): string {
    return typeof result === "object" && "error" in result ? result.error : result.toString();
}

/** A cell reference or range in a formula (for `dependencies`), with an optional sheet name. */
const formulaRangeRegex = /([a-zA-Z0-9]+!)?[A-Z]{1,3}\d{1,4}:[A-Z]{1,3}\d{1,4}/g;
const formulaCellRegex = /([a-zA-Z0-9]+!)?[A-Z]{1,3}\d{1,4}/g;

/** A formula: a master formula, or a clone of a shared formula, with its cached result. */
class FormulaValue {
    /** The cell (to find the master of a shared formula). */
    cell: Cell;
    model: FormulaCellModel;
    /** The translated formula of a shared formula clone, once computed. */
    declare _translatedFormula?: string | undefined;

    /** @param value - nothing from the Cell model setter, which replaces the model right after */
    constructor(cell: Cell, value?: FormulaValueObject) {
        this.cell = cell;

        this.model = {
            address: cell.address,
            type: ValueType.Formula,
            shareType: value ? value.shareType : undefined,
            ref: value ? value.ref : undefined,
            formula: value ? value.formula : undefined,
            sharedFormula: value ? value.sharedFormula : undefined,
            result: value ? value.result : undefined,
        };
    }

    /** The formula fields of a model or value that are set (not undefined: falsy results such as 0 are values). */
    _copyModel(model: FormulaValueObject): FormulaValueObject {
        const copy: FormulaValueObject = {};
        // falsy results (0, false, '') are values too
        if (model.formula !== undefined) copy.formula = model.formula;
        if (model.result !== undefined) copy.result = model.result;
        if (model.ref !== undefined) copy.ref = model.ref;
        if (model.shareType !== undefined) copy.shareType = model.shareType;
        if (model.sharedFormula !== undefined) copy.sharedFormula = model.sharedFormula;
        return copy;
    }

    /** The formula fields that are set (a new object). */
    get value(): FormulaValueObject {
        return this._copyModel(this.model);
    }

    /** NB: replaces the model with one without address and type (unused by the library). */
    set value(value: FormulaValueObject) {
        this.model = this._copyModel(value) as FormulaCellModel;
    }

    /**
     * Throws when a value cannot be the result of a formula (unused by the library).
     * @throws Error for a value that is not null, a string, a number or a date
     */
    validate(value: unknown): void {
        switch (getValueType(value)) {
            case ValueType.Null:
            case ValueType.String:
            case ValueType.Number:
            case ValueType.Date:
                break;
            case ValueType.Merge:
            case ValueType.Hyperlink:
            case ValueType.Formula:
            case ValueType.SharedString:
            case ValueType.RichText:
            case ValueType.Boolean:
            case ValueType.Error:
            case undefined:
                throw new Error("Cannot process that type of result value");
        }
    }

    /**
     * The ranges and cells that the formula mentions (`null` for none).
     * @throws Error for a formula longer than the Excel limit
     */
    get dependencies(): { ranges: RegExpMatchArray | null; cells: RegExpMatchArray | null } {
        // the formula of a master; a clone without master has none, and throws, as it always did
        const formula = this.formula as string;
        checkFormulaLength(formula);
        // find all the ranges and cells mentioned in the formula
        const ranges = formula.match(formulaRangeRegex);
        const cells = formula.replace(formulaRangeRegex, "").match(formulaCellRegex);
        return {
            ranges,
            cells,
        };
    }

    /** The formula, or for a shared formula clone the master's formula translated to this cell. */
    get formula(): string | undefined {
        return this.model.formula || this._getTranslatedFormula();
    }

    set formula(value: string | undefined) {
        this.model.formula = value;
    }

    get formulaType(): FormulaType {
        if (this.model.formula) {
            return FormulaType.Master;
        }
        if (this.model.sharedFormula) {
            return FormulaType.Shared;
        }
        return FormulaType.None;
    }

    get result(): FormulaResult | undefined {
        return this.model.result;
    }

    set result(value: FormulaResult | undefined) {
        this.model.result = value;
    }

    get type(): typeof ValueType.Formula {
        return ValueType.Formula;
    }

    /** The type of the result (Null when there is none, and for a boolean or an error result). */
    get effectiveType(): ValueType {
        const v = this.model.result;
        if (v === null || v === undefined) {
            return ValueType.Null;
        }
        if (v instanceof String || typeof v === "string") {
            return ValueType.String;
        }
        if (typeof v === "number") {
            return ValueType.Number;
        }
        if (v instanceof Date) {
            return ValueType.Date;
        }
        const probe = v as ValueProbe;
        if (probe.text && probe.hyperlink) {
            return ValueType.Hyperlink;
        }
        if (probe.formula) {
            return ValueType.Formula;
        }

        return ValueType.Null;
    }

    get address(): string {
        return this.model.address;
    }

    set address(value: string) {
        this.model.address = value;
    }

    /** The master's formula translated to this cell (computed once), for a shared formula clone. */
    _getTranslatedFormula(): string | undefined {
        if (!this._translatedFormula && this.model.sharedFormula) {
            const { worksheet } = this.cell;
            const master = worksheet.findCell(this.model.sharedFormula);
            // a master without formula throws in slideFormula, as it always did
            this._translatedFormula = master && slideFormula(master.formula as string, master.address, this.model.address);
        }
        return this._translatedFormula;
    }

    toCsvString(): string {
        const { result } = this.model;
        return result ? resultText(result) : "";
    }

    release(): void {}

    /** The text of the result ("" when there is none; the code of an error). */
    toString(): string {
        const { result } = this.model;
        return result ? resultText(result) : "";
    }
}

/** A shared string index (a loaded model of the stream reader), or a `{ sharedString }` value. */
class SharedStringValue {
    model: SharedStringCellModel;

    constructor(cell: Cell, value: SharedStringCellModel["value"]) {
        this.model = {
            address: cell.address,
            type: ValueType.SharedString,
            value,
        };
    }

    get value(): SharedStringCellModel["value"] {
        return this.model.value;
    }

    set value(value: SharedStringCellModel["value"]) {
        this.model.value = value;
    }

    get type(): typeof ValueType.SharedString {
        return ValueType.SharedString;
    }

    get effectiveType(): ValueType {
        return ValueType.SharedString;
    }

    get address(): string {
        return this.model.address;
    }

    set address(value: string) {
        this.model.address = value;
    }

    toCsvString(): string {
        return this.toString();
    }

    release(): void {}

    toString(): string {
        // eslint-disable-next-line @typescript-eslint/no-base-to-string -- a legacy { sharedString } value gives "[object Object]", as it always did
        return this.model.value.toString();
    }
}

/** A boolean. */
class BooleanValue {
    model: BooleanCellModel;

    constructor(cell: Cell, value: boolean) {
        this.model = {
            address: cell.address,
            type: ValueType.Boolean,
            value,
        };
    }

    get value(): boolean {
        return this.model.value;
    }

    set value(value: boolean) {
        this.model.value = value;
    }

    get type(): typeof ValueType.Boolean {
        return ValueType.Boolean;
    }

    get effectiveType(): ValueType {
        return ValueType.Boolean;
    }

    get address(): string {
        return this.model.address;
    }

    set address(value: string) {
        this.model.address = value;
    }

    /** 1 or 0 (numbers). */
    toCsvString(): number {
        return this.model.value ? 1 : 0;
    }

    release(): void {}

    toString(): string {
        return this.model.value.toString();
    }
}

/** An error value (`{ error: "#N/A" }`). */
class ErrorValue {
    model: ErrorCellModel;

    constructor(cell: Cell, value: CellErrorValue) {
        this.model = {
            address: cell.address,
            type: ValueType.Error,
            value,
        };
    }

    get value(): CellErrorValue {
        return this.model.value;
    }

    set value(value: CellErrorValue) {
        this.model.value = value;
    }

    get type(): typeof ValueType.Error {
        return ValueType.Error;
    }

    get effectiveType(): ValueType {
        return ValueType.Error;
    }

    get address(): string {
        return this.model.address;
    }

    set address(value: string) {
        this.model.address = value;
    }

    toCsvString(): string {
        return this.toString();
    }

    release(): void {}

    toString(): string {
        return this.model.value.error.toString();
    }
}

/**
 * Any other object: held with its JSON text (`model.value`) and itself (`model.rawValue`). NB: its type is String, in
 * the model too (there is no JSON value type).
 */
class JSONValue {
    model: StringCellModel & { value: string; rawValue: object };

    constructor(cell: Cell, value: object) {
        this.model = {
            address: cell.address,
            type: ValueType.String,
            value: JSON.stringify(value),
            rawValue: value,
        };
    }

    get value(): object {
        return this.model.rawValue;
    }

    set value(value: object) {
        this.model.rawValue = value;
        this.model.value = JSON.stringify(value);
    }

    get type(): typeof ValueType.String {
        return ValueType.String;
    }

    get effectiveType(): ValueType {
        return ValueType.String;
    }

    get address(): string {
        return this.model.address;
    }

    set address(value: string) {
        this.model.address = value;
    }

    toCsvString(): string {
        return this.model.value;
    }

    release(): void {}

    toString(): string {
        return this.model.value;
    }
}

/** The value holder of a cell. */
type AnyValue =
    | NullValue
    | NumberValue
    | StringValue
    | RichTextValue
    | DateValue
    | HyperlinkValue
    | MergeValue
    | FormulaValue
    | SharedStringValue
    | BooleanValue
    | ErrorValue
    | JSONValue;

/** The marker properties that tell the kind of an object value. */
interface ValueProbe {
    readonly text?: unknown;
    readonly hyperlink?: unknown;
    readonly formula?: unknown;
    readonly sharedFormula?: unknown;
    readonly richText?: unknown;
    readonly sharedString?: unknown;
    readonly error?: unknown;
}

/**
 * The type of a value (see {@link CellValue}): `undefined` for a JSON value (any other object), which has no ValueType
 * (`Cell.Types.JSON` is undefined).
 */
function getValueType(value: unknown): ValueType | undefined {
    if (value === null || value === undefined) {
        return ValueType.Null;
    }
    if (value instanceof String || typeof value === "string") {
        return ValueType.String;
    }
    if (typeof value === "number") {
        return ValueType.Number;
    }
    if (typeof value === "boolean") {
        return ValueType.Boolean;
    }
    if (value instanceof Date) {
        return ValueType.Date;
    }
    // any other value: an object, or a symbol / bigint / function (held as a JSON value)
    const probe = value as ValueProbe;
    if (probe.text && probe.hyperlink) {
        return ValueType.Hyperlink;
    }
    if (probe.formula || probe.sharedFormula) {
        return ValueType.Formula;
    }
    if (probe.richText) {
        return ValueType.RichText;
    }
    if (probe.sharedString) {
        return ValueType.SharedString;
    }
    if (probe.error) {
        return ValueType.Error;
    }
    return undefined;
}

/**
 * The value holder of a type (`undefined`: a JSON value).
 *
 * `value` is of the kind of `type` (as classified by {@link getValueType}), or comes from the Cell model setter: the
 * master of a merge or `undefined`, the holder's model being replaced right after. Hence the casts below.
 * @throws Error for an unknown type
 */
function createValue(type: ValueType | undefined, cell: Cell, value: unknown): AnyValue {
    switch (type) {
        case ValueType.Null:
            return new NullValue(cell);
        case ValueType.Number:
            return new NumberValue(cell, value as number);
        case ValueType.String:
            return new StringValue(cell, value as string);
        case ValueType.Date:
            return new DateValue(cell, value as Date);
        case ValueType.Hyperlink:
            return new HyperlinkValue(cell, value as HyperlinkValueObject | undefined);
        case ValueType.Formula:
            return new FormulaValue(cell, value as FormulaValueObject | undefined);
        case ValueType.Merge:
            return new MergeValue(cell, value as Cell);
        case undefined:
            return new JSONValue(cell, value as object);
        case ValueType.SharedString:
            return new SharedStringValue(cell, value as SharedStringCellModel["value"]);
        case ValueType.RichText:
            return new RichTextValue(cell, value as RichTextObject);
        case ValueType.Boolean:
            return new BooleanValue(cell, value as boolean);
        case ValueType.Error:
            return new ErrorValue(cell, value as CellErrorValue);
        default:
            throw new Error(`Could not create Value of type ${String(type)}`);
    }
}

export default Cell;
