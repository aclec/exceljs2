// A cell of a worksheet, and the value types behind it (NullValue, NumberValue, StringValue, FormulaValue...).
//
// Cell requirements
//  Operate inside a worksheet
//  Store and retrieve a value with a range of types: text, number, date, hyperlink, reference, formula, etc.
//  Manage/use and manipulate cell format either as local to cell or inherited from column or row.

import type { CellErrorValue, Comment, Style } from "../../index";
import colCache from "../utils/col-cache";
import { escapeHtml } from "../utils/under-dash";
import { ErrorValue as ErrorCode, FormulaType, ValueType } from "./enums";
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
    RichTextRun,
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
    readonly sharedString: number;
}

/**
 * The value of a cell: index.d.ts `CellValue` (null, number, string, boolean, Date, error, rich text, hyperlink,
 * formula), a shared string, or any other object, held as a JSON value (its `type` is String, it is written as its JSON
 * text). Setting `undefined` empties the cell (the value is then `null`).
 *
 * The kind of value is told by its type, then by its marker properties (their presence and type, not their truthiness),
 * in this order: a `text` (not null) and a string `hyperlink` (hyperlink), a string `formula` or `sharedFormula`
 * (formula), an array `richText`, a number `sharedString`, a string `error`.
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
        this._value = new NullValue({ address, type: ValueType.Null });

        this.style = this._mergeStyle(row.style, column.style, {});

        // cells merged into this one (created on first merge)
        this._mergeSlaves = undefined;
    }

    /** The worksheet of the cell. */
    get worksheet(): CellWorksheet {
        return this._row.worksheet;
    }

    /** The workbook of the cell. */
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
    /**
     * The number format (`"0.00%"`, `"dd/mm/yyyy"`…). Setting a style property copies the style object (copy on write; the cells of a
     * merge keep sharing it).
     */
    get numFmt(): StyleModel["numFmt"] {
        return this.style.numFmt;
    }

    set numFmt(value: StyleModel["numFmt"]) {
        this._applyStyle("numFmt", value);
    }

    /** The font. Setting it copies the style object (see {@link numFmt}). */
    get font(): StyleModel["font"] {
        return this.style.font;
    }

    set font(value: StyleModel["font"]) {
        this._applyStyle("font", value);
    }

    /** The alignment. Setting it copies the style object (see {@link numFmt}). */
    get alignment(): StyleModel["alignment"] {
        return this.style.alignment;
    }

    set alignment(value: StyleModel["alignment"]) {
        this._applyStyle("alignment", value);
    }

    /** The borders. Setting them copies the style object (see {@link numFmt}). */
    get border(): StyleModel["border"] {
        return this.style.border;
    }

    set border(value: StyleModel["border"]) {
        this._applyStyle("border", value);
    }

    /** The fill. Setting it copies the style object (see {@link numFmt}). */
    get fill(): StyleModel["fill"] {
        return this.style.fill;
    }

    set fill(value: StyleModel["fill"]) {
        this._applyStyle("fill", value);
    }

    /** The protection (locked, hidden). Setting it copies the style object (see {@link numFmt}). */
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
    toCsvString(): string | number {
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
        this._value = new MergeValue(this, master, { address: this.address, type: ValueType.Merge, master: master.address });
        if (!ignoreStyle) {
            this.style = master.style;
        }
    }

    /** Unmerges a merged cell (it becomes empty; its style is reset from its row and column unless `ignoreStyle`). */
    unmerge(ignoreStyle?: boolean): void {
        if (this.type === ValueType.Merge) {
            this._value.release();
            this._value = new NullValue({ address: this.address, type: ValueType.Null });
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

        // the holder is created first: an invalid value throws, leaving the cell as it was
        const holder = valueHolder(this, v);
        this._value.release();
        this._value = holder;
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

    /** The text of the cell (same as {@link text}). */
    toString(): string {
        return this.text;
    }

    /** Turns a string value into a hyperlink (the string is its text); does nothing for another value. */
    _upgradeToHyperlink(hyperlink: string): void {
        // if this cell is a string, turn it into a Hyperlink (not a JSON value, whose type is String too)
        if (this._value instanceof StringValue) {
            this._value = new HyperlinkValue({
                address: this.address,
                type: ValueType.Hyperlink,
                text: this._value.value,
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

    /** Removes the cell from a defined name (a large range kept as a rectangle is cut around the cell). */
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
     * @throws Error for an unknown value type, or the model of a merged cell without master
     */
    set model(value: CellModel) {
        if (value.type === ValueType.Merge) {
            // the model of a merged cell names its master by address (see MergeValue)
            if (!value.master) {
                throw new Error(`The model of the merged cell ${value.address} has no master`);
            }
            this._value.release();
            this._value = new MergeValue(this, this.worksheet.getCell(value.master), value);
        } else {
            // the holder is created first: an unknown type throws, leaving the cell as it was
            const holder = modelHolder(this, value);
            this._value.release();
            this._value = holder;
        }

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
// Each holds the model of the cell value (`model`), given to its constructor: built from a value by `valueHolder()`, or
// adopted by reference from the Cell model setter (`modelHolder()`). The class names are visible (spec, debugging).

/**
 * The contract of the value holders of a {@link Cell}: one class per value type, chosen by `valueHolder()` / `modelHolder()`. The cell
 * delegates `value`, `type`, `effectiveType`, `model`, `toCsvString()` and `toString()` to its holder; the members documented
 * here apply to every implementation.
 */
interface CellValueHolder {
    /** The cell model of the value: address, type, value and the fields of the type (what `cell.model` is built from). */
    readonly model: CellModel;
    /** The value, as `cell.value` returns it (a Merge holder returns its master's value). */
    readonly value: unknown;
    /** The value type ({@link ValueType}) of the holder. */
    readonly type: ValueType;
    /** The type the value displays as: the type of a formula's result, of a merge master's value, else {@link type}. */
    readonly effectiveType: ValueType;
    /** The address of the cell (`A1`); updated when rows or columns are spliced. */
    address: string;
    /** The value as written in a CSV file (a number for a boolean, as the JavaScript version always did). */
    toCsvString(): string | number;
    /** Releases what the holder references before it is replaced (a merge unregisters itself from its master). */
    release(): void;
    /** The value as text (`cell.text`). */
    toString(): string;
}

/** A cell without value. */
class NullValue implements CellValueHolder {
    model: NullCellModel;

    constructor(model: NullCellModel) {
        this.model = model;
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
class NumberValue implements CellValueHolder {
    model: NumberCellModel;

    constructor(model: NumberCellModel) {
        this.model = model;
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
class StringValue implements CellValueHolder {
    model: StringCellModel & { value: string };

    constructor(model: StringCellModel & { value: string }) {
        this.model = model;
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
class RichTextValue implements CellValueHolder {
    model: StringCellModel & { value: RichTextObject };

    constructor(model: StringCellModel & { value: RichTextObject }) {
        this.model = model;
    }

    get value(): RichTextObject {
        return this.model.value;
    }

    set value(value: RichTextObject) {
        this.model.value = value;
    }

    /** The text of the runs, joined. */
    toString(): string {
        return richTextText(this.model.value);
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
class DateValue implements CellValueHolder {
    model: DateCellModel;

    constructor(model: DateCellModel) {
        this.model = model;
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
class HyperlinkValue implements CellValueHolder {
    model: HyperlinkCellModel;

    constructor(model: HyperlinkCellModel) {
        this.model = model;
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

    /** Replaces the model with one of the value, at the same address (unused by the library). */
    set value(value: HyperlinkValueObject) {
        this.model = hyperlinkModel(this.model.address, value.text, value.hyperlink, value.tooltip);
    }

    /** The text shown for the link (as it is: a loaded hyperlink's text may be another value than a string). */
    get text(): HyperlinkCellModel["text"] {
        return this.model.text;
    }

    set text(value: HyperlinkCellModel["text"]) {
        this.model.text = value;
    }

    /** The target: a URL, or a place in the workbook (`#Sheet1!A1`). */
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

    /** The text shown: that of its text value (see {@link hyperlinkText}). */
    toString(): string {
        return hyperlinkText(this.model.text);
    }
}

/** A cell merged into another one (its master): it shows and sets the master's value. */
class MergeValue implements CellValueHolder {
    model: MergeCellModel;
    /** The merged cell. */
    _cell: Cell;
    /** The master of the merge. */
    _master: Cell;

    /** @param model - the model of the merged cell: it names the master by address */
    constructor(cell: Cell, master: Cell, model: MergeCellModel) {
        this.model = model;
        this._cell = cell;
        this._master = master;
        master.addMergeRef(cell);
    }

    /** The master's value. */
    get value(): CellValue {
        return this._master.value;
    }

    /** Sets the master's value; a Cell given as the value becomes the new master. */
    set value(value: CellValue | Cell | undefined) {
        if (value instanceof Cell) {
            this._master.releaseMergeRef(this._cell);
            value.addMergeRef(this._cell);
            this._master = value;
        } else {
            this._master.value = value;
        }
    }

    /** Whether this cell is merged into `master`. */
    isMergedTo(master: Cell): boolean {
        return master === this._master;
    }

    /** The master cell of the merge (whose value and style this cell shows). */
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

/** The text of a formula result: the code of an error, the `toString()` of another value; "" without result. */
function resultText(result: FormulaResult | undefined): string {
    if (result === undefined || result === null) {
        return "";
    }
    return typeof result === "object" && "error" in result ? result.error : result.toString();
}

/** The text of a rich text: the text of its runs, joined. */
function richTextText(value: RichTextObject): string {
    return value.richText.map((t) => t.text).join("");
}

/**
 * The text shown by a hyperlink: its text as the cell of that value would show it (a string as is, the `toString()` of
 * a number, a boolean or a date, the code of an error, the runs of a rich text).
 */
function hyperlinkText(text: HyperlinkCellModel["text"]): string {
    if (typeof text === "string") {
        return text;
    }
    if (typeof text === "number" || typeof text === "boolean" || text instanceof Date) {
        return text.toString();
    }
    return "richText" in text ? richTextText(text) : text.error;
}

/** A cell reference or range in a formula (for `dependencies`), with an optional sheet name. */
const formulaRangeRegex = /([a-zA-Z0-9]+!)?[A-Z]{1,3}\d{1,4}:[A-Z]{1,3}\d{1,4}/g;
const formulaCellRegex = /([a-zA-Z0-9]+!)?[A-Z]{1,3}\d{1,4}/g;

/** A formula: a master formula, or a clone of a shared formula, with its cached result. */
class FormulaValue implements CellValueHolder {
    /** The cell (to find the master of a shared formula). */
    cell: Cell;
    model: FormulaCellModel;
    /** The translated formula of a shared formula clone, once computed. */
    declare _translatedFormula?: string | undefined;

    constructor(cell: Cell, model: FormulaCellModel) {
        this.cell = cell;
        this.model = model;
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

    /** Replaces the model with one of the formula fields of the value that are set, at the same address (unused by the library). */
    set value(value: FormulaValueObject) {
        this.model = { address: this.model.address, type: ValueType.Formula, ...this._copyModel(value) };
    }

    /**
     * Throws when a value cannot be the result of a formula (unused by the library).
     * @throws Error for a value that is not null, a string, a number or a date
     */
    validate(value: unknown): void {
        const holder = valueHolder(this.cell, value);
        if (!(
            holder instanceof NullValue ||
            holder instanceof StringValue ||
            holder instanceof NumberValue ||
            holder instanceof DateValue
        )) {
            throw new Error("Cannot process that type of result value");
        }
    }

    /**
     * The ranges and cells that the formula mentions (`null` for none: also for a shared formula clone whose master is
     * missing or has no formula).
     * @throws Error for a formula longer than the Excel limit
     */
    get dependencies(): { ranges: RegExpMatchArray | null; cells: RegExpMatchArray | null } {
        const { formula } = this;
        if (formula === undefined) {
            return { ranges: null, cells: null };
        }
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

    /** Master (a formula of its own), Shared (a clone of a shared formula), or None. */
    get formulaType(): FormulaType {
        if (this.model.formula) {
            return FormulaType.Master;
        }
        if (this.model.sharedFormula) {
            return FormulaType.Shared;
        }
        return FormulaType.None;
    }

    /** The cached result of the formula (as last computed by Excel, or as set), if any. */
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
        // a boolean or an error result
        return ValueType.Null;
    }

    get address(): string {
        return this.model.address;
    }

    set address(value: string) {
        this.model.address = value;
    }

    /**
     * The master's formula translated to this cell (computed once), for a shared formula clone; undefined when the
     * master is missing or has no formula.
     */
    _getTranslatedFormula(): string | undefined {
        if (!this._translatedFormula && this.model.sharedFormula) {
            const { worksheet } = this.cell;
            const master = worksheet.findCell(this.model.sharedFormula);
            const masterFormula = master && master.formula;
            this._translatedFormula =
                master && masterFormula !== undefined ? slideFormula(masterFormula, master.address, this.model.address) : undefined;
        }
        return this._translatedFormula;
    }

    toCsvString(): string {
        return resultText(this.model.result);
    }

    release(): void {}

    /** The text of the result ("" when there is none; the code of an error). */
    toString(): string {
        return resultText(this.model.result);
    }
}

/** A shared string index (a loaded model of the stream reader), or a `{ sharedString }` value. */
class SharedStringValue implements CellValueHolder {
    model: SharedStringCellModel;

    constructor(model: SharedStringCellModel) {
        this.model = model;
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
class BooleanValue implements CellValueHolder {
    model: BooleanCellModel;

    constructor(model: BooleanCellModel) {
        this.model = model;
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
class ErrorValue implements CellValueHolder {
    model: ErrorCellModel;

    constructor(model: ErrorCellModel) {
        this.model = model;
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
class JSONValue implements CellValueHolder {
    model: StringCellModel & { value: string; rawValue: object };

    constructor(model: StringCellModel & { value: string; rawValue: object }) {
        this.model = model;
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

// =============================================================================
// Creation of the value holders

/** The error codes of a cell value (`{ error }`, index.d.ts `CellErrorValue`). */
const errorCodes: ReadonlySet<string> = new Set<string>(Object.values(ErrorCode));

/** Whether a string is an error code (`"#N/A"`, `"#REF!"`...). */
function isErrorCode(code: string): code is CellErrorValue["error"] {
    return errorCodes.has(code);
}

/** Whether a value is an error value: an object whose `error` is an error code. */
function isCellErrorValue(value: unknown): value is CellErrorValue {
    return typeof value === "object" && value !== null && "error" in value && typeof value.error === "string" && isErrorCode(value.error);
}

/**
 * Whether a value is a run of rich text: an object with a string `text`, and a `font` object or none (the font is a style
 * object, taken as it is, as `cell.font` takes one).
 */
function isRichTextRun(run: unknown): run is RichTextRun {
    return (
        typeof run === "object" &&
        run !== null &&
        "text" in run &&
        typeof run.text === "string" &&
        (!("font" in run) || run.font === undefined || (typeof run.font === "object" && run.font !== null))
    );
}

/** Whether an object is a rich text value: its `richText` is an array of runs. */
function isRichTextObject(value: object): value is RichTextObject {
    return "richText" in value && Array.isArray(value.richText) && value.richText.every(isRichTextRun);
}

/** Whether a value is a `{ sharedString }` value (a number index). */
function isSharedStringRef(value: object): value is SharedStringRef {
    return "sharedString" in value && typeof value.sharedString === "number";
}

/** Whether a value is a formula result: a string, a number, a boolean, a Date or an error value. */
function isFormulaResult(value: unknown): value is FormulaResult {
    return (
        typeof value === "string" ||
        typeof value === "number" ||
        typeof value === "boolean" ||
        value instanceof Date ||
        isCellErrorValue(value)
    );
}

/** Whether a value is the text of a hyperlink: a formula result (see {@link isFormulaResult}) or a rich text. */
function isHyperlinkText(value: unknown): value is HyperlinkCellModel["text"] {
    return isFormulaResult(value) || (typeof value === "object" && value !== null && isRichTextObject(value));
}

const isString = (value: unknown): value is string => typeof value === "string";
const isShareType = (value: unknown): value is "shared" | "array" => value === "shared" || value === "array";

/**
 * An optional field of a value object (`kind` value, field `name`): undefined when absent, undefined or null.
 * @throws Error when it is of another type than `expected`
 */
function optionalField<T>(kind: string, name: string, field: unknown, is: (value: unknown) => value is T, expected: string): T | undefined {
    if (field === undefined || field === null) {
        return undefined;
    }
    if (is(field)) {
        return field;
    }
    throw new Error(`Invalid ${kind} value: its ${name} must be ${expected}`);
}

/** The model of a hyperlink (the tooltip only when not empty). */
function hyperlinkModel(
    address: string,
    text: HyperlinkCellModel["text"],
    hyperlink: string,
    tooltip: string | undefined
): HyperlinkCellModel {
    const model: HyperlinkCellModel = { address, type: ValueType.Hyperlink, text, hyperlink };
    if (tooltip) {
        model.tooltip = tooltip;
    }
    return model;
}

const RESULT_TYPES = "a string, a number, a boolean, a Date or an error value";

/**
 * The holder of an object value (see {@link CellValue}): a Date, a String object (held as its string), or the kind told
 * by its markers (their presence and type: an empty text, formula or error code, a shared string 0 are values), in this
 * order: hyperlink, formula, rich text, shared string, error; any other object is a JSON value.
 * @throws Error for a value of a kind with a field of the wrong type (a hyperlink text, a formula result, a rich text
 *   run, an error code...)
 */
function objectHolder(cell: Cell, address: string, value: object): AnyValue {
    if (value instanceof Date) {
        return new DateValue({ address, type: ValueType.Date, value });
    }
    if (value instanceof String) {
        return new StringValue({ address, type: ValueType.String, value: value.valueOf() });
    }
    if ("hyperlink" in value && typeof value.hyperlink === "string" && "text" in value && value.text !== undefined && value.text !== null) {
        if (!isHyperlinkText(value.text)) {
            throw new Error(`Invalid hyperlink value: its text must be ${RESULT_TYPES}, or a rich text`);
        }
        const tooltip = optionalField("hyperlink", "tooltip", "tooltip" in value ? value.tooltip : undefined, isString, "a string");
        return new HyperlinkValue(hyperlinkModel(address, value.text, value.hyperlink, tooltip));
    }
    if (
        ("formula" in value && typeof value.formula === "string") ||
        ("sharedFormula" in value && typeof value.sharedFormula === "string")
    ) {
        return new FormulaValue(cell, {
            address,
            type: ValueType.Formula,
            shareType: optionalField(
                "formula",
                "shareType",
                "shareType" in value ? value.shareType : undefined,
                isShareType,
                '"shared" or "array"'
            ),
            ref: optionalField("formula", "ref", "ref" in value ? value.ref : undefined, isString, "a string"),
            formula: optionalField("formula", "formula", "formula" in value ? value.formula : undefined, isString, "a string"),
            sharedFormula: optionalField(
                "formula",
                "sharedFormula",
                "sharedFormula" in value ? value.sharedFormula : undefined,
                isString,
                "a string"
            ),
            result: optionalField("formula", "result", "result" in value ? value.result : undefined, isFormulaResult, RESULT_TYPES),
        });
    }
    if ("richText" in value && Array.isArray(value.richText)) {
        if (!isRichTextObject(value)) {
            throw new Error("Invalid rich text value: each run must be an object with a string text (and a font object)");
        }
        return new RichTextValue({ address, type: ValueType.String, value });
    }
    if (isSharedStringRef(value)) {
        return new SharedStringValue({ address, type: ValueType.SharedString, value });
    }
    if ("error" in value && typeof value.error === "string") {
        if (!isCellErrorValue(value)) {
            throw new Error(`Invalid error value: ${value.error} is not an error code (${[...errorCodes].join(", ")})`);
        }
        return new ErrorValue({ address, type: ValueType.Error, value });
    }
    return new JSONValue({ address, type: ValueType.String, value: JSON.stringify(value), rawValue: value });
}

/**
 * The holder of a value (see {@link CellValue}): null and undefined empty the cell.
 * @throws Error for a bigint, a symbol or a function, and for an invalid object value (see {@link objectHolder})
 */
function valueHolder(cell: Cell, value: unknown): AnyValue {
    const { address } = cell;
    if (value === null || value === undefined) {
        return new NullValue({ address, type: ValueType.Null });
    }
    if (typeof value === "string") {
        return new StringValue({ address, type: ValueType.String, value });
    }
    if (typeof value === "number") {
        return new NumberValue({ address, type: ValueType.Number, value });
    }
    if (typeof value === "boolean") {
        return new BooleanValue({ address, type: ValueType.Boolean, value });
    }
    if (typeof value === "object") {
        return objectHolder(cell, address, value);
    }
    // a bigint, a symbol or a function
    throw new Error(`Invalid cell value: a ${typeof value} is not a cell value`);
}

/** Whether the value of a string model is a string (else a rich text). */
function hasStringValue(model: StringCellModel): model is StringCellModel & { value: string } {
    return typeof model.value === "string";
}

/** Whether the value of a string model is a rich text (an object with a `richText` array: its runs are not checked). */
function hasRichTextValue(model: StringCellModel): model is StringCellModel & { value: RichTextObject } {
    const { value } = model;
    return typeof value === "object" && value !== null && Array.isArray(value.richText);
}

/**
 * The holder of a cell model (but a merged cell's, see the Cell model setter), which adopts the model by reference. A
 * String or RichText model holds a string or a rich text, whatever its type: the model of a rich text value has the type
 * String (see {@link RichTextValue}); with another value (a model of the application against its type: the reader
 * gives none), the cell is empty (a Null model, not adopted).
 * @throws Error for an unknown type
 */
function modelHolder(cell: Cell, model: Exclude<CellModel, MergeCellModel>): AnyValue {
    const { type } = model;
    switch (model.type) {
        case ValueType.Null:
            return new NullValue(model);
        case ValueType.Number:
            return new NumberValue(model);
        case ValueType.String:
        case ValueType.RichText:
            if (hasStringValue(model)) {
                return new StringValue(model);
            }
            if (hasRichTextValue(model)) {
                return new RichTextValue(model);
            }
            return new NullValue({ address: model.address, type: ValueType.Null });
        case ValueType.Date:
            return new DateValue(model);
        case ValueType.Hyperlink:
            return new HyperlinkValue(model);
        case ValueType.Formula:
            return new FormulaValue(cell, model);
        case ValueType.SharedString:
            return new SharedStringValue(model);
        case ValueType.Boolean:
            return new BooleanValue(model);
        case ValueType.Error:
            return new ErrorValue(model);
        default:
            throw new Error(`Could not create Value of type ${String(type)}`);
    }
}

export default Cell;
