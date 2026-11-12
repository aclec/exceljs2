// Conversions between A1-style addresses and row / column numbers, with caches for the column letters and the
// addresses of the top-left 100x100 cells.

import { createDictionary } from "./utils";

const addressRegex = /^[A-Z]+\d+$/;
/** `[sheetName!]reference`: the quoted sheet name (`''` for a quote), the unquoted one, the reference. */
const sheetReferenceRegex = /(?:(?:(?:'((?:[^']|'')*)')|([^'^ !]*))!)?(.*)/;

// Grid size of an xlsx worksheet (ECMA-376 / Excel limits)
const MAX_ROWS = 1048576;
const MAX_COLS = 16384;

/**
 * A decoded cell address, e.g. `B3` → `{ address: "B3", col: 2, row: 3, $col$row: "$B$3" }`.
 *
 * The objects returned for the top-left 100x100 cells are cached and shared between callers: do not mutate them (copy
 * them, e.g. `{ ...address, sheetName }`).
 *
 * `col` and `row` are 1-based. For a string without column letters (`"12"`) or without row digits (`"B"`), the missing
 * number is `undefined` at runtime, although typed `number`: validate untrusted text first (`validateAddress`).
 */
export interface CellAddress {
    /** The address without `$` signs, e.g. `"B3"`. */
    readonly address: string;
    /** 1-based column number (1 to {@link MAX_COLS}). */
    readonly col: number;
    /** 1-based row number (1 to {@link MAX_ROWS}). */
    readonly row: number;
    /** The absolute form of the address, e.g. `"$B$3"`. */
    readonly $col$row: string;
}

/** The bounds of a rectangle of cells: 1-based row and column numbers, inclusive. */
export interface RangeBounds {
    top: number;
    left: number;
    bottom: number;
    right: number;
}

/** A range decoded by {@link colCache.decode}, e.g. `A1:B2`, normalised so that `tl` is the top-left corner. */
export interface DecodedRange extends RangeBounds {
    /** The top-left address, e.g. `"A1"`. */
    tl: string;
    /** The bottom-right address, e.g. `"B2"`. */
    br: string;
    /** `"tl:br"`, e.g. `"A1:B2"`. */
    dimensions: string;
}

/** A cell address that may belong to a worksheet, as returned by {@link colCache.decodeEx}. */
export type SheetCellAddress = CellAddress & { readonly sheetName?: string };

/** A corner of a {@link DecodedExRange}: a new object, whose `sheetName` key is always present. */
export interface RangeCorner {
    address: string;
    col: number;
    row: number;
    $col$row: string;
    sheetName: string | undefined;
}

/** A range decoded by {@link colCache.decodeEx}, e.g. `Sheet1!$A$1:$B$2`. */
export interface DecodedExRange extends RangeBounds {
    /** The sheet name, unquoted; `undefined` when the reference has none. */
    sheetName: string | undefined;
    tl: RangeCorner;
    br: RangeCorner;
    /** `"A1:B2"`, without the sheet name nor `$` signs. */
    dimensions: string;
}

/** An error reference decoded by {@link colCache.decodeEx}, e.g. `Sheet1!#REF!`. */
export interface DecodedExError {
    sheetName?: string;
    /** The reference, starting with `#`, e.g. `"#REF!"`. */
    error: string;
}

/** What {@link colCache.decodeEx} returns: a range, an error reference, or a cell address. */
export type DecodedEx = DecodedExRange | DecodedExError | SheetCellAddress;

/** The letter of a 26-letter "digit" of a column, from 0 (`A`) to 25 (`Z`). */
const letter = (i: number): string => String.fromCharCode(65 + i);

/**
 * Decodes the address of a cell given by its row and column numbers, or by its address.
 * @returns A {@link CellAddress}, shared with other callers for the top-left 100x100 cells.
 * @throws Error when the column or the row is out of bounds.
 */
function getAddress(row: number, col: number): CellAddress;
function getAddress(address: string): CellAddress;
function getAddress(r: number | string, c?: number): CellAddress {
    if (typeof r === "string") {
        return colCache.decodeAddress(r);
    }
    // the (row, col) form: n2l throws for a column out of bounds (0 included)
    return colCache.decodeAddress(colCache.n2l(c ?? 0) + r);
}

/**
 * The address of a cell (`encode(row, col)` → `"B3"`) or of a range (`encode(top, left, bottom, right)` → `"A1:B2"`).
 * @throws Error with another number of arguments, or when a column is out of bounds.
 */
function encode(row: number, col: number): string;
function encode(top: number, left: number, bottom: number, right: number): string;
function encode(...args: [number, number] | [number, number, number, number]): string {
    switch (args.length) {
        case 2:
            return colCache.encodeAddress(args[0], args[1]);
        case 4:
            return `${colCache.encodeAddress(args[0], args[1])}:${colCache.encodeAddress(args[2], args[3])}`;
        default:
            throw new Error("Can only encode with 2 or 4 arguments");
    }
}

// =========================================================================
// Column Letter to Number conversion
/**
 * Conversions between column letters and numbers, and between A1-style addresses and row / column numbers.
 *
 * Column numbers are 1-based, from 1 (`A`) to {@link MAX_COLS} (`XFD`); row numbers from 1 to {@link MAX_ROWS}.
 */
const colCache = {
    /** The number of rows of a worksheet (1,048,576). */
    MAX_ROWS,
    /** The number of columns of a worksheet (16,384, column `XFD`). */
    MAX_COLS,
    /** The 26 column letters. */
    _dictionary: Array.from({ length: 26 }, (_, i) => letter(i)) as readonly string[],
    /** The number of letters (1 to 3) of the columns filled in `_l2n` / `_n2l` so far. */
    _l2nFill: 0,
    /** Column letters → number. No prototype: l2n('constructor') must throw. */
    _l2n: createDictionary<number>(),
    /** Column number → letters (index 0 unused). */
    _n2l: [] as string[],
    /** The number of letters of column n. */
    _level(n: number): 1 | 2 | 3 {
        if (n <= 26) {
            return 1;
        }
        if (n <= 26 * 26) {
            return 2;
        }
        return 3;
    },
    /**
     * Fills the caches with all the columns of up to `level` letters.
     * @throws Error when level is 4 or more (no such column in Excel).
     */
    _fill(level: number): void {
        let c: string;
        let v: number;
        let n = 1;
        if (level >= 4) {
            throw new Error("Out of bounds. Excel supports columns from 1 to 16384");
        }
        if (this._l2nFill < 1 && level >= 1) {
            while (n <= 26) {
                c = letter(n - 1);
                this._n2l[n] = c;
                this._l2n[c] = n;
                n++;
            }
            this._l2nFill = 1;
        }
        if (this._l2nFill < 2 && level >= 2) {
            n = 27;
            while (n <= 26 + 26 * 26) {
                v = n - (26 + 1);
                c = letter(Math.floor(v / 26)) + letter(v % 26);
                this._n2l[n] = c;
                this._l2n[c] = n;
                n++;
            }
            this._l2nFill = 2;
        }
        if (this._l2nFill < 3 && level >= 3) {
            n = 26 + 26 * 26 + 1;
            while (n <= 16384) {
                v = n - (26 * 26 + 26 + 1);
                c = letter(Math.floor(v / (26 * 26))) + letter(Math.floor(v / 26) % 26) + letter(v % 26);
                this._n2l[n] = c;
                this._l2n[c] = n;
                n++;
            }
            this._l2nFill = 3;
        }
    },
    /**
     * The number of a column, e.g. `"AB"` → 28.
     * @param l - Column letters, upper case, without `$`.
     * @throws Error when l is not a column of Excel (`A` to `XFD`).
     */
    l2n(l: string): number {
        if (!this._l2n[l]) {
            this._fill(l.length);
        }
        const n = this._l2n[l];
        if (!n) {
            throw new Error(`Out of bounds. Invalid column letter: ${l}`);
        }
        return n;
    },
    /**
     * The letters of a column, e.g. 28 → `"AB"`.
     * @param n - An integer from 1 to {@link MAX_COLS}.
     * @throws Error when n is out of these bounds.
     */
    n2l(n: number): string {
        if (n < 1 || n > MAX_COLS) {
            throw new Error(`${n} is out of bounds. Excel supports columns from 1 to ${MAX_COLS}`);
        }
        if (!this._n2l[n]) {
            this._fill(this._level(n));
        }
        // filled above for an integer n in bounds. A fractional or NaN n reads undefined, as it always did: a whole-row
        // or whole-column reference such as `Sheet1!$1:$1` (print titles) decodes to NaN bounds, and must not throw
        // (the defined-name parser keeps the references that decodeEx accepts)
        // eslint-disable-next-line @typescript-eslint/no-unsafe-type-assertion -- documented above: string for every integer in bounds
        return this._n2l[n] as string;
    },

    // =========================================================================
    // Address processing
    /** The decoded addresses of the top-left 100x100 cells, by address and absolute address (`$A$1`). */
    _hash: createDictionary<CellAddress>(),

    /**
     * Checks that value looks like a cell address: upper-case column letters then row digits, e.g. `"B3"` (no `$`).
     * @throws Error when it does not.
     */
    validateAddress(value: string): true {
        if (!addressRegex.test(value)) {
            throw new Error(`Invalid Address: ${value}`);
        }
        return true;
    },

    /**
     * Decodes a cell address such as `"B3"` or `"$B$3"`: the column letters (upper case) then the row digits, `$` signs
     * ignored; parsing stops at the first other character after both.
     * @returns A {@link CellAddress}, shared with other callers for the top-left 100x100 cells.
     * @throws Error when the column is beyond `XFD` or the row beyond {@link MAX_ROWS}.
     */
    decodeAddress(value: string): CellAddress {
        const addr = value.length < 5 && this._hash[value];
        if (addr) {
            return addr;
        }
        let hasCol = false;
        let col = "";
        let colNumber: number | undefined = 0;
        let hasRow = false;
        let row = "";
        let rowNumber: number | undefined = 0;
        for (let i = 0, char; i < value.length; i++) {
            char = value.charCodeAt(i);
            // col should before row
            if (!hasRow && char >= 65 && char <= 90) {
                // 65 = 'A'.charCodeAt(0)
                // 90 = 'Z'.charCodeAt(0)
                hasCol = true;
                col += value[i];
                // colNumber starts from 1
                colNumber = colNumber * 26 + char - 64;
            } else if (char >= 48 && char <= 57) {
                // 48 = '0'.charCodeAt(0)
                // 57 = '9'.charCodeAt(0)
                hasRow = true;
                row += value[i];
                // rowNumber starts from 0
                rowNumber = rowNumber * 10 + char - 48;
            } else if (hasRow && hasCol && char !== 36) {
                // 36 = '$'.charCodeAt(0)
                break;
            }
        }
        if (!hasCol) {
            colNumber = undefined;
        } else if (colNumber > MAX_COLS) {
            throw new Error(`Out of bounds. Invalid column letter: ${col}`);
        }
        if (!hasRow) {
            rowNumber = undefined;
        } else if (rowNumber > MAX_ROWS) {
            throw new Error(`Out of bounds. Excel supports rows from 1 to ${MAX_ROWS}: ${value}`);
        }

        // in case $row$col
        value = col + row;

        const address = {
            address: value,
            // undefined without column letters / row digits (see CellAddress: a whole-row or whole-column reference)
            // eslint-disable-next-line @typescript-eslint/no-unsafe-type-assertion -- the documented exception of CellAddress.col
            col: colNumber as number,
            // eslint-disable-next-line @typescript-eslint/no-unsafe-type-assertion -- the documented exception of CellAddress.row
            row: rowNumber as number,
            $col$row: `$${col}$${row}`,
        };

        // mem fix - cache only the tl 100x100 square
        if (address.col <= 100 && address.row <= 100) {
            this._hash[value] = address;
            this._hash[address.$col$row] = address;
        }

        return address;
    },

    getAddress,

    /**
     * Decodes a cell address (`"B3"`) or a range (`"A1:B2"`, normalised so that tl is the top-left corner).
     * @throws Error when a column or a row is out of bounds.
     */
    decode(value: string): CellAddress | DecodedRange {
        // exactly one colon: a range
        const colon = value.indexOf(":");
        if (colon !== -1 && value.indexOf(":", colon + 1) === -1) {
            const tl = this.decodeAddress(value.slice(0, colon));
            const br = this.decodeAddress(value.slice(colon + 1));
            const top = Math.min(tl.row, br.row);
            const left = Math.min(tl.col, br.col);
            const bottom = Math.max(tl.row, br.row);
            const right = Math.max(tl.col, br.col);
            // reconstruct tl, br and dimensions
            const tlAddress = this.n2l(left) + top;
            const brAddress = this.n2l(right) + bottom;
            return { top, left, bottom, right, tl: tlAddress, br: brAddress, dimensions: `${tlAddress}:${brAddress}` };
        }
        return this.decodeAddress(value);
    },

    /**
     * Decodes a reference `[sheetName!][$]col[$]row[:[$]col[$]row]`: a cell address or a range, optionally preceded by a
     * sheet name (quoted with `'` when it contains spaces, `''` for a quote), or an error such as `Sheet1!#REF!`.
     * @throws Error when a column or a row is out of bounds.
     */
    decodeEx(value: string): DecodedEx {
        // matches every string (all of its groups are optional): without a match, the whole value is the reference
        const groups = sheetReferenceRegex.exec(value);

        const sheetName = groups ? groups[1] || groups[2] : undefined; // Qouted and unqouted groups
        const reference = groups ? (groups[3] ?? "") : value; // Remaining address

        // a range: the text up to the first colon, and from there to the next one (or the end)
        const colon = reference.indexOf(":");
        if (colon !== -1) {
            const end = reference.indexOf(":", colon + 1);
            const tl = this.decodeAddress(reference.slice(0, colon));
            const br = this.decodeAddress(reference.slice(colon + 1, end === -1 ? undefined : end));
            const top = Math.min(tl.row, br.row);
            const left = Math.min(tl.col, br.col);
            const bottom = Math.max(tl.row, br.row);
            const right = Math.max(tl.col, br.col);

            const tlAddress = this.n2l(left) + top;
            const brAddress = this.n2l(right) + bottom;

            return {
                top,
                left,
                bottom,
                right,
                sheetName,
                tl: { address: tlAddress, col: left, row: top, $col$row: `$${this.n2l(left)}$${top}`, sheetName },
                br: {
                    address: brAddress,
                    col: right,
                    row: bottom,
                    $col$row: `$${this.n2l(right)}$${bottom}`,
                    sheetName,
                },
                dimensions: `${tlAddress}:${brAddress}`,
            };
        }
        if (reference.startsWith("#")) {
            return sheetName ? { sheetName, error: reference } : { error: reference };
        }

        const address = this.decodeAddress(reference);
        return sheetName ? { sheetName, ...address } : address;
    },

    /**
     * The address of a cell, e.g. (3, 2) → `"B3"`.
     * @throws Error when col is out of bounds.
     */
    encodeAddress(row: number, col: number): string {
        return colCache.n2l(col) + row;
    },

    encode,

    /**
     * Whether a cell is in a rectangle.
     * @param range - `[left, top, (ignored), right, bottom]`.
     * @param address - `[col, row]`.
     */
    inRange(range: readonly [number, number, unknown, number, number], address: readonly [number, number]): boolean {
        const [left, top, , right, bottom] = range;
        const [col, row] = address;
        return col >= left && col <= right && row >= top && row <= bottom;
    },
};

/** The type of the {@link colCache} object. */
export type ColCache = typeof colCache;

export default colCache;
