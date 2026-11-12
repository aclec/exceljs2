// The corner of an image in a worksheet: a cell (0-based column and row) and an offset in it, in EMU.

import colCache from "../utils/col-cache";
import type { AnchorModel } from "../types";

/**
 * What an {@link Anchor} reads of its worksheet: the size of the anchor cell. The lookups must not create the row or the
 * column (the position may come from a file).
 */
export interface AnchorWorksheet {
    /** The column `col` (1-based) if it exists or is defined by a loaded `<col>`, without creating it. */
    _findColumn(col: number): { readonly isCustomWidth: boolean; readonly width?: number | undefined } | undefined;
    /** The row `row` (1-based) if it exists. */
    findRow(row: number): { readonly height?: number | undefined } | undefined;
}

/**
 * A position given to the {@link Anchor} constructor: native fields (`nativeCol`, `nativeRow`, and the offsets, which
 * default to 0), or a fractional 0-based `col` and `row` (`{ col: 1.5, row: 0 }`: the middle of column B, row 1), e.g.
 * another Anchor.
 */
export type AnchorPosition = Partial<AnchorModel> &
    ({ readonly col?: undefined; readonly row?: undefined } | { readonly col: number; readonly row: number });

// The sizes of the cells are approximated in EMU (the unit of the offsets): 10,000 per unit of column width / row height.
/** Width of a column without a custom width, in EMU. */
const DEFAULT_COL_WIDTH = 640000;
/** Height of a row without a height, in EMU. */
const DEFAULT_ROW_HEIGHT = 180000;

/**
 * The corner of an image (`tl` or `br` of an image range). `nativeCol` / `nativeRow` are the 0-based column and row of
 * the anchor cell, `nativeColOff` / `nativeRowOff` the offset inside it, in EMU. `col` / `row` are the same position as a
 * fractional 0-based column / row, computed from the size of the anchor cell.
 */
class Anchor {
    /** The worksheet that gives the sizes of the cells; without one, the default sizes apply. */
    worksheet: AnchorWorksheet | undefined;
    /** 0-based column of the anchor cell. */
    nativeCol!: number;
    /** Offset in the anchor cell, from its left edge. */
    nativeColOff!: number;
    /** 0-based row of the anchor cell. */
    nativeRow!: number;
    /** Offset in the anchor cell, from its top edge. */
    nativeRowOff!: number;

    /**
     * @param worksheet - gives the sizes of the cells (for `col` / `row`)
     * @param address - a cell address (`"B2"`: the anchor is the top-left corner of the cell, moved by `offset` columns
     *   and rows), or a position; nothing (or an object with neither `nativeCol` nor `col`) is the top-left corner of A1
     * @param offset - added to the column and row of an address or a `col` / `row` position (not to native fields)
     */
    constructor(worksheet?: AnchorWorksheet, address?: string | AnchorPosition | null, offset = 0) {
        this.worksheet = worksheet;

        if (!address) {
            this.nativeCol = 0;
            this.nativeColOff = 0;
            this.nativeRow = 0;
            this.nativeRowOff = 0;
        } else if (typeof address === "string") {
            const decoded = colCache.decodeAddress(address);
            this.nativeCol = decoded.col + offset;
            this.nativeColOff = 0;
            this.nativeRow = decoded.row + offset;
            this.nativeRowOff = 0;
        } else if (address.nativeCol !== undefined) {
            this.nativeCol = address.nativeCol || 0;
            this.nativeColOff = address.nativeColOff || 0;
            this.nativeRow = address.nativeRow || 0;
            this.nativeRowOff = address.nativeRowOff || 0;
        } else if (address.col !== undefined) {
            this.col = address.col + offset;
            this.row = address.row + offset;
        } else {
            this.nativeCol = 0;
            this.nativeColOff = 0;
            this.nativeRow = 0;
            this.nativeRowOff = 0;
        }
    }

    /**
     * The model as an Anchor: an Anchor, `null` and `undefined` are returned as they are.
     *
     * NB: any other value is passed as the worksheet of a new Anchor (as the JavaScript version did), which gives an
     * anchor at A1: the method is not used by the library.
     */
    static asInstance(model: unknown): Anchor | null | undefined {
        return model instanceof Anchor || model == null ? model : new Anchor(model as AnchorWorksheet);
    }

    /** The fractional 0-based column: `nativeCol` plus the offset relative to the width of the column. */
    get col(): number {
        return this.nativeCol + Math.min(this.colWidth - 1, this.nativeColOff) / this.colWidth;
    }

    set col(v: number) {
        this.nativeCol = Math.floor(v);
        this.nativeColOff = Math.floor((v - this.nativeCol) * this.colWidth);
    }

    /** The fractional 0-based row: `nativeRow` plus the offset relative to the height of the row. */
    get row(): number {
        return this.nativeRow + Math.min(this.rowHeight - 1, this.nativeRowOff) / this.rowHeight;
    }

    set row(v: number) {
        this.nativeRow = Math.floor(v);
        this.nativeRowOff = Math.floor((v - this.nativeRow) * this.rowHeight);
    }

    // the size of the anchor cell: reading it must not create the row or column (and, for a column, the
    // ones before it), as the position comes from the file
    /** The width of the anchor column, in EMU (the default for a column without a custom width). */
    get colWidth(): number {
        const column = this.worksheet && this.worksheet._findColumn(this.nativeCol + 1);
        return column && column.isCustomWidth && column.width !== undefined ? Math.floor(column.width * 10000) : DEFAULT_COL_WIDTH;
    }

    /** The height of the anchor row, in EMU (the default for a row without a height). */
    get rowHeight(): number {
        const row = this.worksheet && this.worksheet.findRow(this.nativeRow + 1);
        return row && row.height ? Math.floor(row.height * 10000) : DEFAULT_ROW_HEIGHT;
    }

    /** The native fields, as a new object. */
    get model(): AnchorModel {
        return {
            nativeCol: this.nativeCol,
            nativeColOff: this.nativeColOff,
            nativeRow: this.nativeRow,
            nativeRowOff: this.nativeRowOff,
        };
    }

    set model(value: AnchorModel) {
        this.nativeCol = value.nativeCol;
        this.nativeColOff = value.nativeColOff;
        this.nativeRow = value.nativeRow;
        this.nativeRowOff = value.nativeRowOff;
    }
}

export default Anchor;
