// A row of a worksheet: its cells (a sparse array indexed by column number - 1), height, style, outline.

import { ValueType } from "./enums";
import colCache, { type CellAddress } from "../utils/col-cache";
import Cell from "./cell";
import type { CellValue, CellWorksheet, StyleName } from "./cell";
import type Column from "./column";
import { eachEntry, loopSteps, someEntry } from "../utils/sparse-array";
import type { CellModel, RowBreakModel, RowModel, StyleModel } from "../types";

/**
 * What a {@link Row} reads of its worksheet (a Worksheet, or the streaming WorksheetWriter / WorksheetReader): the
 * columns (for new cells and column keys), the outline level, the page breaks, and what its cells read.
 */
export interface RowWorksheet extends CellWorksheet {
    /** The column `col` (1-based), created if needed; undefined for an invalid number. */
    getColumn(col: number): Column | undefined;
    getColumnKey(key: string): Column | undefined;
    eachColumnKey(callback: (column: Column, key: string) => void): void;
    /** Commits the rows up to `row` (streaming writer: resolves when the output can take more data). */
    _commitRow(row: Row): void | Promise<void>;
    readonly properties: { readonly outlineLevelRow?: number | undefined };
    readonly rowBreaks: RowBreakModel[];
}

/** The location of a cell given to {@link Row.getCellEx}: its address (`"B2"`) and 1-based column. */
export interface RowCellAddress {
    readonly address: string;
    readonly col: number;
    readonly row?: number;
}

/**
 * The values of a row: a sparse array indexed by column number (index 0 unused), a contiguous array (index 0 set: the
 * values of columns 1, 2...), or an object keyed by column key.
 */
export type RowValues = readonly (CellValue | undefined)[] | { readonly [key: string]: CellValue | undefined };

/** Options of {@link Row.eachCell}. */
export interface EachCellOptions {
    /** Also visit the empty cells, up to the last one (creating the missing ones). */
    readonly includeEmpty?: boolean;
}

/** An iteratee of {@link Row.eachCell}: the cell and its 1-based column number. */
export type CellIteratee = (cell: Cell, colNumber: number) => void;

/**
 * A row of a worksheet. Its cells are created on demand (`getCell`); a new cell inherits the style of the row (else of
 * its column).
 */
class Row {
    /** The worksheet. */
    _worksheet: RowWorksheet;
    /** 1-based row number. */
    _number: number;
    /** The cells, indexed by column number - 1 (sparse). */
    _cells: (Cell | undefined)[];
    /** The style of the row: the default style of the cells created in it. Shared: copy on write. */
    style: StyleModel;
    declare _outlineLevel?: number | undefined;
    declare _hidden?: boolean | undefined;
    /** The height in points, when set. */
    declare height?: number;

    /** @param number - 1-based row number */
    constructor(worksheet: RowWorksheet, number: number) {
        this._worksheet = worksheet;
        this._number = number;
        this._cells = [];
        this.style = {};
        this.outlineLevel = 0;
    }

    /** The 1-based row number. */
    get number(): number {
        return this._number;
    }

    /** The worksheet of the row. */
    get worksheet(): RowWorksheet {
        return this._worksheet;
    }

    /**
     * Informs the streaming writer that this row (and all the rows before it) are complete and ready to write. Has no
     * effect on a Worksheet document.
     * @returns in the streaming writer, a promise to await before writing more rows (backpressure)
     */
    commit(): void | Promise<void> {
        return this._worksheet._commitRow(this);
    }

    /** Helps GC by breaking cyclic references: the row cannot be used afterwards. */
    destroy(): void {
        // the row is unusable once destroyed: its fields are removed, not typed as optional
        const self = this as Partial<Pick<Row, "_worksheet" | "_cells" | "style">>;
        delete self._worksheet;
        delete self._cells;
        delete self.style;
    }

    /** The cell of a column (1-based) if it exists. */
    findCell(colNumber: number): Cell | undefined {
        return this._cells[colNumber - 1];
    }

    /** The cell at a location, created if needed. */
    getCellEx(address: RowCellAddress): Cell {
        let cell = this._cells[address.col - 1];
        if (!cell) {
            const column = this._worksheet.getColumn(address.col);
            cell = new Cell(this, column, address.address);
            this._cells[address.col - 1] = cell;
        }
        return cell;
    }

    /** The cell of a column, created if needed: by 1-based number, column key, or letters (`"B"`). */
    getCell(col: number | string): Cell {
        if (typeof col === "string") {
            // is it a key?
            const column = this._worksheet.getColumnKey(col);
            if (column) {
                col = column.number;
            } else {
                col = colCache.l2n(col);
            }
        }
        return (
            this._cells[col - 1] ||
            this.getCellEx({
                address: colCache.encodeAddress(this._number, col),
                row: this._number,
                col,
            })
        );
    }

    /**
     * Removes `count` cells from column `start` (1-based), inserts the `inserts` values there, and shifts the cells on
     * the right (their value, style and note). Only this row changes.
     */
    splice(start: number, count: number, ...inserts: (CellValue | undefined)[]): void {
        const nKeep = start + count;
        const nExpand = inserts.length - count;
        const nEnd = this._cells.length;
        // the loops below only visit the cells i whose source or destination exists: the other steps do nothing,
        // and most of them are empty when the last cell is far
        let i;
        let cSrc;
        let cDst;

        if (nExpand < 0) {
            // remove cells
            for (i of loopSteps(this._cells, 1, start + inserts.length, nEnd, -nExpand)) {
                cDst = this._cells[i - 1];
                cSrc = this._cells[i - nExpand - 1];
                if (cSrc) {
                    cDst = this.getCell(i);
                    cDst.value = cSrc.value;
                    cDst.style = cSrc.style;

                    cDst._comment = cSrc._comment;
                } else if (cDst) {
                    cDst.value = null;
                    cDst.style = {};

                    cDst._comment = undefined;
                }
            }
        } else if (nExpand > 0) {
            // insert new cells
            for (i of loopSteps(this._cells, 1, nKeep, nEnd, nExpand).reverse()) {
                cSrc = this._cells[i - 1];
                if (cSrc) {
                    cDst = this.getCell(i + nExpand);
                    cDst.value = cSrc.value;
                    cDst.style = cSrc.style;

                    cDst._comment = cSrc._comment;
                } else {
                    this._cells[i + nExpand - 1] = undefined;
                }
            }
            // the cells are counted up to the last one moved, existing or not (cellCount)
            if (nEnd >= nKeep && this._cells.length < nEnd + nExpand) {
                this._cells.length = nEnd + nExpand;
            }
        }

        // now add the new values
        for (i = 0; i < inserts.length; i++) {
            cDst = this.getCell(start + i);
            cDst.value = inserts[i];
            cDst.style = {};

            cDst._comment = undefined;
        }
    }

    /**
     * Calls `iteratee` for each non-empty cell of the row, in column order (with `{ includeEmpty: true }`: for every cell
     * up to the last one, creating the missing ones).
     */
    eachCell(iteratee: CellIteratee): void;
    eachCell(options: EachCellOptions | null | undefined, iteratee: CellIteratee): void;
    eachCell(optionsOrIteratee: EachCellOptions | CellIteratee | null | undefined, iteratee?: CellIteratee): void {
        // eachCell(iteratee) or eachCell(options, iteratee)
        const callback = iteratee || (optionsOrIteratee as CellIteratee);
        const options = iteratee ? (optionsOrIteratee as EachCellOptions | null | undefined) : null;
        if (options && options.includeEmpty) {
            const n = this._cells.length;
            for (let i = 1; i <= n; i++) {
                callback(this.getCell(i), i);
            }
        } else {
            eachEntry(this._cells, (cell, index) => {
                if (cell.type !== ValueType.Null) {
                    callback(cell, index + 1);
                }
            });
        }
    }

    /**
     * Iterates over the cells of this row that exist, including empty ones, without creating the missing cells up to
     * the last one as eachCell({includeEmpty: true}) does (16,383 of them for a cell in column XFD).
     */
    _eachExistingCell(iteratee: CellIteratee): void {
        eachEntry(this._cells, (cell, index) => {
            iteratee(cell, index + 1);
        });
    }

    // ===========================================================================
    // Page Breaks
    /** Adds a manual page break after this row, from column `lft` to column `rght` (1-based; default: all of them). */
    addPageBreak(lft?: number, rght?: number): void {
        const ws = this._worksheet;
        const left = Math.max(0, (lft ?? 0) - 1) || 0;
        const right = Math.max(0, (rght ?? 0) - 1) || 16838;
        const pb: RowBreakModel = {
            id: this._number,
            max: right,
            man: 1,
        };
        if (left) pb.min = left;

        ws.rowBreaks.push(pb);
    }

    /** The values of the non-empty cells, as a sparse array indexed by column number (index 0 unused). */
    get values(): CellValue[] {
        const values: CellValue[] = [];
        eachEntry(this._cells, (cell) => {
            if (cell.type !== ValueType.Null) {
                values[cell.col] = cell.value;
            }
        });
        return values;
    }

    /**
     * Replaces the cells of the row (not additive: the existing cells are removed) with values (see {@link RowValues});
     * nothing empties the row.
     */
    set values(value: RowValues | null | undefined) {
        // this operation is not additive - any prior cells are removed
        this._cells = [];
        if (!value) {
            // empty row
        } else if (value instanceof Array) {
            const array: readonly (CellValue | undefined)[] = value;
            let offset = 0;
            if (Object.hasOwn(array, "0")) {
                // contiguous array - start at column 1
                offset = 1;
            }
            // a sparse array: row.values of a row with a cell in column XFD
            eachEntry(array, (item, index) => {
                this.getCellEx({
                    address: colCache.encodeAddress(this._number, index + offset),
                    row: this._number,
                    col: index + offset,
                }).value = item;
            });
        } else {
            const byKey = value;
            // assume object with column keys
            this._worksheet.eachColumnKey((column, key) => {
                // a key such as 'constructor' must not pick up an Object.prototype member
                if (key in Object.prototype && !Object.hasOwn(byKey, key)) {
                    return;
                }
                if (byKey[key] !== undefined) {
                    this.getCellEx({
                        address: colCache.encodeAddress(this._number, column.number),
                        row: this._number,
                        col: column.number,
                    }).value = byKey[key];
                }
            });
        }
    }

    /** Whether the row has a non-empty cell. */
    get hasValues(): boolean {
        return someEntry(this._cells, (cell) => cell.type !== ValueType.Null);
    }

    /** The number of cells, empty ones included: the column number of the last cell. */
    get cellCount(): number {
        return this._cells.length;
    }

    /** The number of non-empty cells. */
    get actualCellCount(): number {
        let count = 0;
        this.eachCell(() => {
            count++;
        });
        return count;
    }

    /** The first and last columns (1-based) of the non-empty cells, or null when there are none. */
    get dimensions(): { min: number; max: number } | null {
        let min = 0;
        let max = 0;
        eachEntry(this._cells, (cell) => {
            if (cell.type !== ValueType.Null) {
                if (!min || min > cell.col) {
                    min = cell.col;
                }
                if (max < cell.col) {
                    max = cell.col;
                }
            }
        });
        return min > 0
            ? {
                  min,
                  max,
              }
            : null;
    }

    // =========================================================================
    // styles
    /** Sets a style property of the row and of its existing cells. */
    _applyStyle<K extends StyleName>(name: K, value: StyleModel[K]): StyleModel[K] {
        // copy on write: the style object may be shared with other rows, columns or cells
        this.style = { ...this.style, [name]: value };
        eachEntry(this._cells, (cell) => {
            cell._applyStyle(name, value);
        });
        return value;
    }

    /** The number format of the row. Setting a style property sets it on the row and on its existing cells (copy on write). */
    get numFmt(): StyleModel["numFmt"] {
        return this.style.numFmt;
    }

    set numFmt(value: StyleModel["numFmt"]) {
        this._applyStyle("numFmt", value);
    }

    /** The font of the row. Setting it also sets it on the existing cells (see {@link numFmt}). */
    get font(): StyleModel["font"] {
        return this.style.font;
    }

    set font(value: StyleModel["font"]) {
        this._applyStyle("font", value);
    }

    /** The alignment of the row. Setting it also sets it on the existing cells (see {@link numFmt}). */
    get alignment(): StyleModel["alignment"] {
        return this.style.alignment;
    }

    set alignment(value: StyleModel["alignment"]) {
        this._applyStyle("alignment", value);
    }

    /** The protection of the row. Setting it also sets it on the existing cells (see {@link numFmt}). */
    get protection(): StyleModel["protection"] {
        return this.style.protection;
    }

    set protection(value: StyleModel["protection"]) {
        this._applyStyle("protection", value);
    }

    /** The borders of the row. Setting them also sets them on the existing cells (see {@link numFmt}). */
    get border(): StyleModel["border"] {
        return this.style.border;
    }

    set border(value: StyleModel["border"]) {
        this._applyStyle("border", value);
    }

    /** The fill of the row. Setting it also sets it on the existing cells (see {@link numFmt}). */
    get fill(): StyleModel["fill"] {
        return this.style.fill;
    }

    set fill(value: StyleModel["fill"]) {
        this._applyStyle("fill", value);
    }

    /** Whether the row is hidden. */
    get hidden(): boolean {
        return !!this._hidden;
    }

    set hidden(value: boolean | undefined) {
        this._hidden = value;
    }

    /** The outline level (0 when not set). */
    get outlineLevel(): number {
        return this._outlineLevel || 0;
    }

    set outlineLevel(value: number | undefined) {
        this._outlineLevel = value;
    }

    /** Whether the row is collapsed: its outline level is at least the worksheet's `outlineLevelRow`. */
    get collapsed(): boolean {
        const level = this._worksheet.properties.outlineLevelRow;
        // no level in the worksheet properties: never collapsed (`n >= undefined` is false)
        return !!(this._outlineLevel && level !== undefined && this._outlineLevel >= level);
    }

    // =========================================================================
    /**
     * The model of the row (the models of its cells: see `Cell.model`), or null when it has neither height nor cells
     * with a model.
     */
    get model(): RowModel | null {
        const cells: CellModel[] = [];
        let min = 0;
        let max = 0;
        eachEntry(this._cells, (cell) => {
            const cellModel = cell.model;
            if (cellModel) {
                if (!min || min > cell.col) {
                    min = cell.col;
                }
                if (max < cell.col) {
                    max = cell.col;
                }
                cells.push(cellModel);
            }
        });

        return this.height || cells.length
            ? {
                  cells,
                  number: this.number,
                  min,
                  max,
                  height: this.height,
                  style: this.style,
                  hidden: this.hidden,
                  outlineLevel: this.outlineLevel,
                  collapsed: this.collapsed,
              }
            : null;
    }

    /**
     * Sets the row from a model: its cells (merged cells are skipped: the worksheet merges them), height, hidden,
     * outline level, and a copy of its style. A cell without address is the cell after the previous one.
     * @throws Error when the model is of another row number
     */
    set model(value: RowModel) {
        if (value.number !== this._number) {
            throw new Error("Invalid row number in model");
        }
        this._cells = [];
        let previousAddress: CellAddress | undefined;
        for (const cellModel of value.cells) {
            if (cellModel.type === ValueType.Merge) {
                // special case - don't add this types (the worksheet merges the cells)
                continue;
            }
            let address: CellAddress | undefined;
            if (cellModel.address) {
                address = colCache.decodeAddress(cellModel.address);
            } else if (previousAddress) {
                // This is a <c> element without an r attribute
                // Assume that it's the cell for the next column
                const { row } = previousAddress;
                const col = previousAddress.col + 1;
                address = {
                    row,
                    col,
                    address: colCache.encodeAddress(row, col),
                    $col$row: `$${colCache.n2l(col)}$${row}`,
                };
            }
            previousAddress = address;
            // a first cell without address throws, as it always did
            const cell = this.getCellEx(address!);
            cell.model = cellModel;
        }

        if (value.height) {
            this.height = value.height;
        } else {
            delete this.height;
        }

        this.hidden = value.hidden;
        this.outlineLevel = value.outlineLevel || 0;

        this.style = (value.style && (JSON.parse(JSON.stringify(value.style)) as StyleModel)) || {};
    }
}

export default Row;
