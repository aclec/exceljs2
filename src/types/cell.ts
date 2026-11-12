// Internal models of cells, rows and columns: the JSON exchanged between src/doc (Cell / Row / Column `model`) and the
// sheet xforms (cell-xform, row-xform, col-xform), the worksheet xform and the stream writer / reader.
//
// The models are mutated in place along the way: `prepare()` adds the ids needed to render (styleId, ssId, si, ...),
// `reconcile()` turns the raw parsed values back into document values (and may change `type`). The types below describe
// the document-side models with the prepare-side additions as optional fields; the raw parsed cell (before reconcile) is
// {@link ParsedCellModel}.

import type { CellErrorValue, Font } from "../../index";
import type Range from "../doc/range";
import type { ValueType } from "../doc/enums";
import type { FontModel, StyleModel } from "./style";
import type { NoteModel } from "./sheet";

/** One run of rich text (rich-text-xform). Producer: the application, rich-text-xform parse. A run without text is not rendered. */
export interface RichTextRun {
    text: string;
    font?: FontModel | Partial<Font>;
}

/** A rich text value: `{ richText: [...] }`. */
export interface RichTextValue {
    richText: RichTextRun[];
}

/** The value of a formula result. A `t="d"` result that does not parse as a date stays a string. */
export type FormulaResult = string | number | boolean | Date | CellErrorValue;

/** Fields of every cell model. */
interface CellModelBase {
    /** The cell reference (`A1`). Absent from the models of hyperlink / formula values once their value was set. */
    address: string;
    /** The style: always set by `Cell.model` (`{}` when none); set by reconcile only for a known style id. */
    style?: StyleModel;
    /** The note of the cell. Producer: Cell.model (Note.model), reconcile (from the comments part). */
    comment?: NoteModel;
    /** prepare: the cell format id (only when non-zero). parse: the `s` attribute. reconcile: set to `undefined` (key kept). */
    styleId?: number;
}

/** An empty cell (rendered only when it has a style). */
export interface NullCellModel extends CellModelBase {
    type: typeof ValueType.Null;
}

/** A cell merged into another: `master` is the address of the top-left cell of the merge. */
export interface MergeCellModel extends CellModelBase {
    type: typeof ValueType.Merge;
    /** Absent in a parsed `<c>` without value and style (the merges are applied from `mergeCells`). */
    master?: string;
}

/** A number. */
export interface NumberCellModel extends CellModelBase {
    type: typeof ValueType.Number;
    value: number;
}

/**
 * A string, a rich text, or a JSON value (all have `type` String in their model; RichTextValue's `type` getter is RichText).
 * A JSON value has `value: JSON.stringify(rawValue)` and `rawValue`.
 */
export interface StringCellModel extends CellModelBase {
    type: typeof ValueType.String | typeof ValueType.RichText;
    value: string | RichTextValue;
    /** JSON values only: the value itself. */
    rawValue?: unknown;
    /** prepare: the shared string index, when shared strings are used. */
    ssId?: number;
}

/** A date. */
export interface DateCellModel extends CellModelBase {
    type: typeof ValueType.Date;
    value: Date;
    /** prepare: set (to `true`) for a workbook with the 1904 date system. */
    date1904?: true;
}

/** A hyperlink. */
export interface HyperlinkCellModel extends CellModelBase {
    type: typeof ValueType.Hyperlink;
    /** The value shown: a string from the application; on read, the value of the cell, whatever its type. */
    text: string | number | boolean | Date | CellErrorValue | RichTextValue;
    /** The link target. */
    hyperlink: string;
    tooltip?: string;
    /** prepare: the shared string index of a non-scalar text. */
    ssId?: number;
}

/**
 * A formula (master, shared formula clone, or array formula). All the keys are present on the document side, possibly
 * `undefined`.
 */
export interface FormulaCellModel extends CellModelBase {
    type: typeof ValueType.Formula;
    /** The formula text (without `=`); absent on a shared formula clone. */
    formula?: string;
    /** On a shared formula clone: the address of its master. */
    sharedFormula?: string;
    result?: FormulaResult;
    /** `"shared"` (master of shared formulas) or `"array"` (the parsed model of a cell copies `<f t>` as is, see ParsedCellModel). */
    shareType?: "shared" | "array";
    /** The range of a shared / array formula master. */
    ref?: string;
    /** prepare: the shared formula index. parse: the raw `si` attribute (a string); deleted by reconcile. */
    si?: number | string;
    /** prepare: on a master that gets its first clone, the range of its clones (rendered as `ref`). */
    range?: Range;
    /** prepare: set (to `true`) for a workbook with the 1904 date system. */
    date1904?: true;
}

/** A shared string index (stream reader only: prepare and render ignore this type, the cell renders as `<c r s/>`). */
export interface SharedStringCellModel extends CellModelBase {
    type: typeof ValueType.SharedString;
    value: number;
}

/** A boolean. */
export interface BooleanCellModel extends CellModelBase {
    type: typeof ValueType.Boolean;
    value: boolean;
}

/** An error value. */
export interface ErrorCellModel extends CellModelBase {
    type: typeof ValueType.Error;
    value: CellErrorValue;
}

/**
 * The model of a cell, discriminated by `type` (ValueType). Producer: `Cell.model` (src/doc/cell.ts), cell-xform reconcile.
 * Consumer: cell-xform prepare / render, `Cell.model =`. `Cell.model =` adopts the object by reference: later mutations
 * (prepare) show through the cell's value.
 */
export type CellModel =
    | NullCellModel
    | MergeCellModel
    | NumberCellModel
    | StringCellModel
    | DateCellModel
    | HyperlinkCellModel
    | FormulaCellModel
    | SharedStringCellModel
    | BooleanCellModel
    | ErrorCellModel;

/**
 * A cell as parsed by cell-xform, before reconcile (partial: the fields reconcile reads). `type` is the parsed type;
 * `value` is raw: a shared string index (number) for `t="s"`, a `{ richText }` for an inline rich string, a number, a
 * boolean, a Date (`t="d"`), an error, or a string. reconcile resolves it into a {@link CellModel} (String → RichText,
 * Number → Date for a date format, hyperlinks and comments attached).
 */
export interface ParsedCellModel {
    address?: string;
    type: ValueType;
    value?: string | number | boolean | Date | CellErrorValue | RichTextValue;
    styleId?: number;
    style?: StyleModel;
    formula?: string;
    sharedFormula?: string;
    shareType?: string;
    si?: string;
    ref?: string;
    result?: FormulaResult;
    text?: HyperlinkCellModel["text"];
    hyperlink?: string;
    tooltip?: string;
    comment?: NoteModel;
}

/**
 * The model of a row. Producer: `Row.model` (null when the row has no height and no cell), row-xform parse. Consumer:
 * row-xform prepare / render, `Row.model =` (which ignores `min`, `max`, `collapsed`).
 *
 * @typeParam TCell - the cells: {@link CellModel} (document side, prepare, render, after reconcile), or
 *   {@link ParsedCellModel} (`RowModel<ParsedCellModel>`, as parsed)
 */
export interface RowModel<TCell = CellModel> {
    /** 1-based row number. parse: `NaN` when `r` is missing (the worksheet xform infers it). */
    number: number;
    cells: TCell[];
    /** The first / last column (1-based) with a cell; 0 for a row without cells. parse: from `spans` (NaN if malformed), or undefined. */
    min?: number;
    max?: number;
    height?: number;
    style?: StyleModel | null;
    hidden?: boolean;
    outlineLevel?: number;
    collapsed?: boolean;
    /** prepare: the cell format id (when truthy). parse: `s`. reconcile: set to `undefined`. */
    styleId?: number;
    /** parse only (the `bestFit` attribute): never rendered, ignored by Row. */
    bestFit?: boolean;
}

/**
 * A `<col>` range (col-xform): columns `min`..`max` (1-based) sharing the same definition. Producer: `Column.toModel`
 * (worksheet `cols`), col-xform parse. Consumer: col-xform prepare / render, `Column.fromModel` / `Column.rangesFromModel`.
 */
export interface ColModel {
    min: number;
    max: number;
    /** Column.toModel: the width or 9; parse: undefined when absent (the key is present). */
    width?: number;
    style?: StyleModel;
    /** Column.toModel only, not rendered. */
    isCustomWidth?: boolean;
    hidden?: boolean;
    outlineLevel?: number;
    collapsed?: boolean;
    /** parse / render only. */
    bestFit?: boolean;
    /** prepare / parse: the cell format id. reconcile keeps it. */
    styleId?: number;
}
