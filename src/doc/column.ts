// A column of a worksheet: header, key, width, style, outline. The worksheet condenses runs of equivalent columns into
// `<col>` ranges when it is written (Column.toModel).

import type { ColModel, StyleModel } from "../types";
import { isEqual } from "../utils/under-dash";
import colCache from "../utils/col-cache";
import { ValueType } from "./enums";
import type Cell from "./cell";
import type { CellValue, StyleName } from "./cell";

/** The width of a column without a width (in characters). */
const DEFAULT_COLUMN_WIDTH = 9;

/**
 * The definition of a column (`worksheet.columns = [...]`, `column.defn`). A {@link ColModel} is a definition too (its
 * extra fields are ignored).
 */
export interface ColumnDefn {
    /** The header: one row (a string) or several (an array), written in the first rows of the column. */
    header?: string | string[] | undefined;
    /** The key of the column, for the rows set from objects (`row.values = { key: value }`). */
    key?: string | undefined;
    /** The width in characters (default {@link DEFAULT_COLUMN_WIDTH}). */
    width?: number | undefined;
    /** `null` (a parsed `<col>` without style) is the same as no style. */
    style?: StyleModel | null | undefined;
    hidden?: boolean | undefined;
    outlineLevel?: number | undefined;
}

/**
 * What a {@link Column} reads of its worksheet (a Worksheet, or the streaming WorksheetWriter / WorksheetReader): the
 * cells of the column, the column keys, and the properties.
 */
export interface ColumnWorksheet {
    getCell(row: number, col: number): Cell;
    getColumnKey(key: string): Column | undefined;
    setColumnKey(key: string, column: Column): void;
    deleteColumnKey(key: string): void;
    eachRow(
        options: { readonly includeEmpty?: boolean } | null | undefined,
        iteratee: (row: { getCell(col: number): Cell }, rowNumber: number) => void
    ): void;
    readonly properties: { readonly outlineLevelCol?: number | undefined; readonly defaultColWidth?: number | undefined };
}

/** A range of columns of a loaded `<col>` (see {@link Column.rangesFromModel}): `min`..`max`, 1-based. */
export interface ColumnRange {
    min: number;
    max: number;
    defn: ColModel;
    /** A Column standing for the columns of the range (created by the worksheet when it writes the model). */
    column?: Column;
}

/** The columns that {@link Column.fromModel} would create, without creating them. */
export interface ColumnRanges {
    /** Sorted and disjoint. */
    ranges: ColumnRange[];
    /** The number of columns: those up to `length` that are in no range are plain columns. */
    length: number;
}

/** The fields of a column that {@link Column.equivalentTo} compares. */
export interface ColumnLook {
    readonly width?: number | undefined;
    readonly hidden?: boolean | undefined;
    readonly outlineLevel?: number | undefined;
    readonly style?: StyleModel | null | undefined;
}

/** A run of columns `min`..`max` (1-based) that `column` stands for (see {@link Column.segmentsToModel}). */
export interface ColumnSegment {
    min: number;
    max: number;
    column: Column;
}

/**
 * A column of a worksheet: its header, key, width, style, outline level. A column is created on demand
 * (`worksheet.getColumn`); its style is the default of the cells created after it is set (and setting a style property
 * also sets it on the existing cells).
 */
class Column {
    /** The worksheet. */
    _worksheet: ColumnWorksheet;
    /** 1-based column number. */
    _number: number;
    // the fields below are set by the defn setter (not at all when the constructor gets defn false)
    declare _header?: string | string[] | undefined;
    declare _key?: string | undefined;
    declare _width?: number | undefined;
    /** The width was not given: the default width is not written when the sheet default is another one. */
    declare _implicitWidth?: boolean;
    declare _hidden?: boolean;
    declare _outlineLevel?: number | undefined;
    /** The style of the column: the default style of the cells created in it. Shared: copy on write. */
    declare style: StyleModel;

    /**
     * @param number - 1-based column number
     * @param defn - the definition; `false` sets nothing (the caller sets `defn` next); nothing gives the defaults
     */
    constructor(worksheet: ColumnWorksheet, number: number, defn?: ColumnDefn | false | null) {
        this._worksheet = worksheet;
        this._number = number;
        if (defn !== false) {
            // sometimes defn will follow
            this.defn = defn;
        }
    }

    /** The 1-based column number. */
    get number(): number {
        return this._number;
    }

    /** The worksheet of the column. */
    get worksheet(): ColumnWorksheet {
        return this._worksheet;
    }

    /** The column letters (`"A"`). */
    get letter(): string {
        return colCache.n2l(this._number);
    }

    /** The width in characters, if any. */
    get width(): number | undefined {
        return this._width;
    }

    set width(value: number | undefined) {
        this._width = value;
        // a width assigned explicitly is written even when it equals DEFAULT_COLUMN_WIDTH
        this._implicitWidth = false;
    }

    /** Whether the width is set to something else than the default width. */
    get isCustomWidth(): boolean {
        return this.width !== undefined && this.width !== DEFAULT_COLUMN_WIDTH;
    }

    /** The definition of the column (a new object). */
    get defn(): ColumnDefn {
        return {
            header: this._header,
            key: this.key,
            width: this.width,
            style: this.style,
            hidden: this.hidden,
            outlineLevel: this.outlineLevel,
        };
    }

    /**
     * Sets the definition: key, width (default {@link DEFAULT_COLUMN_WIDTH}), outline level, style, then the header
     * (written in the first rows of the column), hidden. Nothing resets the column (no header, key, width, style...).
     */
    set defn(value: ColumnDefn | null | undefined) {
        if (value) {
            this.key = value.key;
            this.width = value.width !== undefined ? value.width : DEFAULT_COLUMN_WIDTH;
            this._implicitWidth = value.width === undefined;
            this.outlineLevel = value.outlineLevel;
            if (value.style) {
                this.style = value.style;
            } else {
                this.style = {};
            }

            // headers must be set after style
            this.header = value.header;
            this._hidden = !!value.hidden;
        } else {
            delete this._header;
            delete this._key;
            this.width = undefined;
            this.style = {};
            this.outlineLevel = 0;
        }
    }

    /** The header rows: `[header]` for a single header (`[undefined]` without one). */
    get headers(): readonly (string | undefined)[] {
        return this._header && this._header instanceof Array ? this._header : [this._header];
    }

    /** The header: one row (a string) or several (an array). */
    get header(): string | string[] | undefined {
        return this._header;
    }

    /** Sets the header, and writes it in the first rows of the column. */
    set header(value: string | string[] | undefined) {
        if (value !== undefined) {
            this._header = value;
            this.headers.forEach((text, index) => {
                this._worksheet.getCell(index + 1, this.number).value = text;
            });
        } else {
            this._header = undefined;
        }
    }

    /** The key of the column (rows set from objects use it). */
    get key(): string | undefined {
        return this._key;
    }

    /** Sets the key, and registers it in the worksheet (replacing the previous key of this column). */
    set key(value: string | undefined) {
        const column = this._key && this._worksheet.getColumnKey(this._key);
        if (column === this) {
            this._worksheet.deleteColumnKey(this._key!);
        }

        this._key = value;
        if (value) {
            this._worksheet.setColumnKey(value, this);
        }
    }

    /** Whether the column is hidden. */
    get hidden(): boolean {
        return !!this._hidden;
    }

    set hidden(value: boolean) {
        this._hidden = value;
    }

    /** The outline level (0 when not set). */
    get outlineLevel(): number {
        return this._outlineLevel || 0;
    }

    set outlineLevel(value: number | undefined) {
        this._outlineLevel = value;
    }

    /** Whether the column is collapsed: its outline level is at least the worksheet's `outlineLevelCol`. */
    get collapsed(): boolean {
        const level = this._worksheet.properties.outlineLevelCol;
        // no level in the worksheet properties: never collapsed (`n >= undefined` is false)
        return !!(this._outlineLevel && level !== undefined && this._outlineLevel >= level);
    }

    /** JSON of the key, width and headers. */
    toString(): string {
        return JSON.stringify({
            key: this.key,
            width: this.width,
            headers: this.headers.length ? this.headers : undefined,
        });
    }

    /** Whether the column is written as the other one (same width, hidden, outline level and style). */
    equivalentTo(other: ColumnLook): boolean {
        return (
            this.width === other.width &&
            this.hidden === other.hidden &&
            this.outlineLevel === other.outlineLevel &&
            isEqual(this.style, other.style)
        );
    }

    /** Whether the column needs no `<col>` (default width, visible, no outline, no style). */
    get isDefault(): boolean {
        if (this.isCustomWidth) {
            return false;
        }
        // an explicit width of DEFAULT_COLUMN_WIDTH must be written when the sheet default differs
        const properties = this._worksheet && this._worksheet.properties;
        const sheetDefault = properties && properties.defaultColWidth;
        if (this.width !== undefined && !this._implicitWidth && sheetDefault !== undefined && this.width !== sheetDefault) {
            return false;
        }
        if (this.hidden) {
            return false;
        }
        if (this.outlineLevel) {
            return false;
        }
        const s = this.style;
        if (s && (s.font || s.numFmt || s.alignment || s.border || s.fill || s.protection)) {
            return false;
        }
        return true;
    }

    /** The number of header rows. */
    get headerCount(): number {
        return this.headers.length;
    }

    /**
     * Calls `iteratee` with the cell of this column in each row that has values (with `{ includeEmpty: true }`: in
     * every row up to the last one, creating the missing rows and cells).
     */
    eachCell(iteratee: (cell: Cell, rowNumber: number) => void): void;
    eachCell(options: { readonly includeEmpty?: boolean } | null | undefined, iteratee: (cell: Cell, rowNumber: number) => void): void;
    eachCell(
        options: { readonly includeEmpty?: boolean } | null | undefined | ((cell: Cell, rowNumber: number) => void),
        iteratee?: (cell: Cell, rowNumber: number) => void
    ): void {
        const colNumber = this.number;
        // eachCell(iteratee) or eachCell(options, iteratee)
        const callback = iteratee || (options as (cell: Cell, rowNumber: number) => void);
        this._worksheet.eachRow(iteratee ? (options as { readonly includeEmpty?: boolean } | null | undefined) : null, (row, rowNumber) => {
            callback(row.getCell(colNumber), rowNumber);
        });
    }

    /** The values of the column, as a sparse array indexed by row number (index 0 unused). */
    get values(): CellValue[] {
        const v: CellValue[] = [];
        this.eachCell((cell, rowNumber) => {
            if (cell && cell.type !== ValueType.Null) {
                v[rowNumber] = cell.value;
            }
        });
        return v;
    }

    /**
     * Sets the values of the column: a sparse array indexed by row number, or a contiguous array (index 0 set) for rows
     * 1, 2... Nothing does nothing.
     */
    set values(v: readonly CellValue[] | null | undefined) {
        if (!v) {
            return;
        }
        const colNumber = this.number;
        let offset = 0;
        if (Object.hasOwn(v, "0")) {
            // assume contiguous array, start at row 1
            offset = 1;
        }
        v.forEach((value, index) => {
            this._worksheet.getCell(index + offset, colNumber).value = value;
        });
    }

    // =========================================================================
    // styles
    /** Sets a style property of the column and of its existing cells. */
    _applyStyle<K extends StyleName>(name: K, value: StyleModel[K]): StyleModel[K] {
        // copy on write: the style object may be shared with other rows, columns or cells
        this.style = { ...this.style, [name]: value };
        this.eachCell((cell) => {
            cell._applyStyle(name, value);
        });
        return value;
    }

    /** The number format of the column. Setting a style property sets it on the column and on its existing cells (copy on write). */
    get numFmt(): string | undefined {
        return this.style.numFmt;
    }

    set numFmt(value: StyleModel["numFmt"]) {
        this._applyStyle("numFmt", value);
    }

    /** The font of the column. Setting it also sets it on the existing cells (see {@link numFmt}). */
    get font(): StyleModel["font"] {
        return this.style.font;
    }

    set font(value: StyleModel["font"]) {
        this._applyStyle("font", value);
    }

    /** The alignment of the column. Setting it also sets it on the existing cells (see {@link numFmt}). */
    get alignment(): StyleModel["alignment"] {
        return this.style.alignment;
    }

    set alignment(value: StyleModel["alignment"]) {
        this._applyStyle("alignment", value);
    }

    /** The protection of the column. Setting it also sets it on the existing cells (see {@link numFmt}). */
    get protection(): StyleModel["protection"] {
        return this.style.protection;
    }

    set protection(value: StyleModel["protection"]) {
        this._applyStyle("protection", value);
    }

    /** The borders of the column. Setting them also sets them on the existing cells (see {@link numFmt}). */
    get border(): StyleModel["border"] {
        return this.style.border;
    }

    set border(value: StyleModel["border"]) {
        this._applyStyle("border", value);
    }

    /** The fill of the column. Setting it also sets it on the existing cells (see {@link numFmt}). */
    get fill(): StyleModel["fill"] {
        return this.style.fill;
    }

    set fill(value: StyleModel["fill"]) {
        this._applyStyle("fill", value);
    }

    // =============================================================================
    // static functions

    /** The `<col>` ranges of an array of columns (index `n - 1` holds column `n`): see {@link Column.segmentsToModel}. */
    static toModel(columns: readonly (Column | undefined)[] | null | undefined): ColModel[] | undefined {
        // Convert array of Column into compressed list cols
        const segments: ColumnSegment[] = [];
        if (columns) {
            columns.forEach((column, index) => {
                segments.push({ min: index + 1, max: index + 1, column: column! });
            });
        }
        return Column.segmentsToModel(segments);
    }

    /**
     * Compresses runs of equivalent columns into a list of cols (default columns are left out).
     * @param segments - `[{min, max, column}]` in column order, where column stands for each of the columns min..max
     * @returns the cols, or `undefined` when there are none
     */
    static segmentsToModel(segments: readonly ColumnSegment[]): ColModel[] | undefined {
        const cols: ColModel[] = [];
        let col: ColModel | null = null;
        for (const { min, max, column } of segments) {
            if (column.isDefault) {
                if (col) {
                    col = null;
                }
            } else if (!col || !column.equivalentTo(col)) {
                col = {
                    min,
                    max,
                    width: column.width !== undefined ? column.width : DEFAULT_COLUMN_WIDTH,
                    style: column.style,
                    isCustomWidth: column.isCustomWidth,
                    hidden: column.hidden,
                    outlineLevel: column.outlineLevel,
                    collapsed: column.collapsed,
                };
                cols.push(col);
            } else {
                col.max = max;
            }
        }
        return cols.length ? cols : undefined;
    }

    /**
     * The columns that {@link Column.fromModel} would create, without creating them. Ranges are sorted and disjoint (a
     * column belongs to the first col that covers it); the columns up to `length` that are in no range are plain columns.
     */
    static rangesFromModel(cols: readonly ColModel[] | null | undefined): ColumnRanges {
        const ranges: ColumnRange[] = [];
        let count = 1;
        (cols || [])
            .slice()
            .sort((pre, next) => pre.min - next.min)
            .forEach((col) => {
                const min = Math.ceil(col.min);
                if (min > count) {
                    count = min;
                }
                const max = Math.floor(col.max);
                if (count <= max) {
                    ranges.push({ min: count, max, defn: col });
                    count = max + 1;
                }
            });
        return { ranges, length: count - 1 };
    }

    /** The definition of column `n` in `colRanges` (see {@link Column.rangesFromModel}), if any (binary search). */
    static defnInRanges(colRanges: ColumnRanges | null | undefined, n: number): ColModel | undefined {
        const ranges = colRanges ? colRanges.ranges : [];
        let lo = 0;
        let hi = ranges.length - 1;
        while (lo <= hi) {
            const mid = (lo + hi) >> 1;
            const range = ranges[mid]!;
            if (n < range.min) {
                hi = mid - 1;
            } else if (n > range.max) {
                lo = mid + 1;
            } else {
                return range.defn;
            }
        }
        return undefined;
    }

    /**
     * Creates the columns of a list of cols, up to the last one (the columns between the cols are plain columns).
     * Sorts `cols` in place.
     * @returns the columns, or `null` when there are none
     */
    static fromModel(worksheet: ColumnWorksheet, cols: ColModel[] | null | undefined): Column[] | null {
        cols = cols || [];
        const columns: Column[] = [];
        let count = 1;
        let index = 0;
        /**
         * sort cols by min
         * If it is not sorted, the subsequent column configuration will be overwritten
         * */
        cols = cols.sort(function (pre, next) {
            return pre.min - next.min;
        });
        while (index < cols.length) {
            const col = cols[index++]!;
            while (count < col.min) {
                columns.push(new Column(worksheet, count++));
            }
            while (count <= col.max) {
                columns.push(new Column(worksheet, count++, col));
            }
        }
        return columns.length ? columns : null;
    }
}

export default Column;
