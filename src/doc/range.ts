// A rectangle of cells (`A1:B2`, optionally on a named sheet): the dimensions of a worksheet, merges, defined names...

import colCache from "../utils/col-cache";

/**
 * The bounds of a {@link Range}: 1-based row and column numbers, inclusive. The keys are always present; a bound is
 * `undefined` for a range decoded from an error reference (`Sheet1!#REF!`), and 0 for an empty range (`new Range()`):
 * the getters of the Range then give 1.
 */
export interface RangeModel {
    top: number | undefined;
    left: number | undefined;
    bottom: number | undefined;
    right: number | undefined;
    /** Absent from the model of an empty range. */
    sheetName?: string | undefined;
}

/** A rectangle given by its bounds (all truthy: 1-based). */
export interface RangeLocation {
    readonly top: number;
    readonly left: number;
    readonly bottom: number;
    readonly right: number;
    readonly sheetName?: string | undefined;
}

/**
 * The arguments that define a range (Range constructor, `decode`, `worksheet.mergeCells`...):
 * - nothing: an empty range (all bounds 0, i.e. `A1:A1`);
 * - a Range (copied), a {@link RangeLocation}, a reference (`"A1:B2"`, `"Sheet1!$A$1"`), or an array of arguments;
 * - two addresses (`"A1", "B2"`) and a sheet name, which is ignored (as the JavaScript version always did);
 * - four bounds and a sheet name (`1, 1, 2, 2, "Sheet1"`).
 *
 * The corners may be given in any order: the bounds are normalised (top <= bottom, left <= right).
 */
export type RangeArgs =
    | []
    | [range: Range | RangeLocation | string | RangeArgs]
    | [tl: string, br: string, sheetName?: string]
    | [top: number, left: number, bottom: number, right: number, sheetName?: string];

/** The sheet names written without quotes in a reference. Anchored: linear. */
const plainSheetNameRegex = /^[a-zA-Z0-9]*$/;

/**
 * A rectangle of cells. Used for the dimensions of a worksheet, merges, defined names, shared formulas...
 *
 * The getters `top`, `left`, `bottom`, `right` give 1 for a bound that is 0 or undefined in the model.
 */
class Range {
    /** The bounds (see {@link RangeModel}). */
    declare model: RangeModel;

    /** @throws Error with more than 5 arguments */
    constructor(...args: RangeArgs) {
        this.decode(args);
    }

    /**
     * Sets the bounds from two addresses (`setTLBR("B2", "A1")`: the third argument is ignored), or from four numbers
     * and a sheet name. The bounds are normalised.
     */
    setTLBR(tl: string, br: string, sheetName?: string): void;
    setTLBR(top: number, left: number, bottom: number, right: number, sheetName?: string): void;
    setTLBR(...args: [string, string, string?] | [number, number, number, number, string?]): void {
        if (args.length < 4) {
            // setTLBR(tl, br, s)
            const tl = colCache.decodeAddress(args[0] as string);
            const br = colCache.decodeAddress(args[1] as string);
            this.model = {
                top: Math.min(tl.row, br.row),
                left: Math.min(tl.col, br.col),
                bottom: Math.max(tl.row, br.row),
                right: Math.max(tl.col, br.col),
                sheetName: args[2] as string | undefined,
            };

            // NB: replaces the model above, without the sheet name
            this.setTLBR(tl.row, tl.col, br.row, br.col, undefined);
        } else {
            // setTLBR(t, l, b, r, s)
            const [t, l, b, r, s] = args as [number, number, number, number, string?];
            this.model = {
                top: Math.min(t, b),
                left: Math.min(l, r),
                bottom: Math.max(t, b),
                right: Math.max(l, r),
                sheetName: s,
            };
        }
    }

    /**
     * Sets the range from a list of arguments (see {@link RangeArgs}): `decode(["A1:B2"])`, `decode([1, 1, 2, 2])`.
     * @throws Error with more than 5 arguments
     */
    decode(argv: Readonly<RangeArgs>): void {
        switch (argv.length) {
            case 5: // [t,l,b,r,s]
                this.setTLBR(argv[0], argv[1], argv[2], argv[3], argv[4]);
                break;
            case 4: // [t,l,b,r]
                this.setTLBR(argv[0], argv[1], argv[2], argv[3]);
                break;

            case 3: // [tl,br,s]
                this.setTLBR(argv[0], argv[1], argv[2]);
                break;
            case 2: // [tl,br]
                this.setTLBR(argv[0], argv[1]);
                break;

            case 1: {
                const value = argv[0];
                if (value instanceof Range) {
                    // copy constructor
                    this.model = {
                        top: value.model.top,
                        left: value.model.left,
                        bottom: value.model.bottom,
                        right: value.model.right,
                        sheetName: value.sheetName,
                    };
                } else if (value instanceof Array) {
                    // an arguments array
                    this.decode(value);
                } else if (typeof value === "object" && value.top && value.left && value.bottom && value.right) {
                    // a model
                    this.model = {
                        top: value.top,
                        left: value.left,
                        bottom: value.bottom,
                        right: value.right,
                        sheetName: value.sheetName,
                    };
                } else {
                    // [sheetName!]tl:br
                    const tlbr = colCache.decodeEx(value as string);
                    if ("top" in tlbr && tlbr.top) {
                        this.model = {
                            top: tlbr.top,
                            left: tlbr.left,
                            bottom: tlbr.bottom,
                            right: tlbr.right,
                            sheetName: tlbr.sheetName,
                        };
                    } else {
                        // a cell; an error reference (#REF!) has no row and column
                        const cell = "row" in tlbr ? tlbr : undefined;
                        this.model = {
                            top: cell?.row,
                            left: cell?.col,
                            bottom: cell?.row,
                            right: cell?.col,
                            sheetName: tlbr.sheetName,
                        };
                    }
                }
                break;
            }

            case 0:
                this.model = {
                    top: 0,
                    left: 0,
                    bottom: 0,
                    right: 0,
                };
                break;

            default:
                throw new Error(`Invalid number of arguments to _getDimensions() - ${(argv as readonly unknown[]).length}`);
        }
    }

    /** The first row (1-based). */
    get top(): number {
        return this.model.top || 1;
    }

    set top(value: number) {
        this.model.top = value;
    }

    /** The first column (1-based). */
    get left(): number {
        return this.model.left || 1;
    }

    set left(value: number) {
        this.model.left = value;
    }

    /** The last row (1-based, inclusive). */
    get bottom(): number {
        return this.model.bottom || 1;
    }

    set bottom(value: number) {
        this.model.bottom = value;
    }

    /** The last column (1-based, inclusive). */
    get right(): number {
        return this.model.right || 1;
    }

    set right(value: number) {
        this.model.right = value;
    }

    /** The sheet of the range, if any. */
    get sheetName(): string | undefined {
        return this.model.sheetName;
    }

    set sheetName(value: string | undefined) {
        this.model.sheetName = value;
    }

    /** The sheet name as it prefixes a reference (`Sheet1!`, `'My sheet'!`), or "" when there is none. */
    get _serialisedSheetName(): string {
        const { sheetName } = this.model;
        if (sheetName) {
            if (plainSheetNameRegex.test(sheetName)) {
                return `${sheetName}!`;
            }
            return `'${sheetName}'!`;
        }
        return "";
    }

    /** Grows the range to include the rectangle (an empty bound of the range takes the given one). */
    expand(top: number, left: number, bottom: number, right: number): void {
        if (!this.model.top || top < this.top) this.top = top;
        if (!this.model.left || left < this.left) this.left = left;
        if (!this.model.bottom || bottom > this.bottom) this.bottom = bottom;
        if (!this.model.right || right > this.right) this.right = right;
    }

    /** Grows the range to include the non-empty cells of a row (nothing for an empty row). */
    expandRow(row: { readonly number: number; readonly dimensions: { min: number; max: number } | null } | null | undefined): void {
        if (row) {
            const { dimensions, number } = row;
            if (dimensions) {
                this.expand(number, dimensions.min, number, dimensions.max);
            }
        }
    }

    /** Grows the range to include a cell (`"B2"`, `"Sheet1!B2"`). */
    expandToAddress(addressStr: string): void {
        const address = colCache.decodeEx(addressStr) as { row: number; col: number };
        this.expand(address.row, address.col, address.row, address.col);
    }

    /** The top-left address (`"A1"`). */
    get tl(): string {
        return colCache.n2l(this.left) + this.top;
    }

    /** The absolute top-left address (`"$A$1"`). */
    get $t$l(): string {
        return `$${colCache.n2l(this.left)}$${this.top}`;
    }

    /** The bottom-right address (`"B2"`). */
    get br(): string {
        return colCache.n2l(this.right) + this.bottom;
    }

    /** The absolute bottom-right address (`"$B$2"`). */
    get $b$r(): string {
        return `$${colCache.n2l(this.right)}$${this.bottom}`;
    }

    /** The reference, with the sheet name if any (`"Sheet1!A1:B2"`; `"A1:A1"` for a single cell). */
    get range(): string {
        return `${this._serialisedSheetName + this.tl}:${this.br}`;
    }

    /** The absolute reference (`"Sheet1!$A$1:$B$2"`). */
    get $range(): string {
        return `${this._serialisedSheetName + this.$t$l}:${this.$b$r}`;
    }

    /** The reference, as a single address for a single cell (`"A1"`). */
    get shortRange(): string {
        return this.count > 1 ? this.range : this._serialisedSheetName + this.tl;
    }

    /** The absolute reference, as a single address for a single cell (`"$A$1"`). */
    get $shortRange(): string {
        return this.count > 1 ? this.$range : this._serialisedSheetName + this.$t$l;
    }

    /** The number of cells. */
    get count(): number {
        return (1 + this.bottom - this.top) * (1 + this.right - this.left);
    }

    /** The reference ({@link Range.range}). */
    toString(): string {
        return this.range;
    }

    /** Whether the ranges share a cell (ranges of two different named sheets never do). */
    intersects(other: Range | RangeLocation): boolean {
        if (other.sheetName && this.sheetName && other.sheetName !== this.sheetName) return false;
        if (other.bottom < this.top) return false;
        if (other.top > this.bottom) return false;
        if (other.right < this.left) return false;
        if (other.left > this.right) return false;
        return true;
    }

    /** Whether the range contains a cell (`"B2"`, `"Sheet1!B2"`). */
    contains(addressStr: string): boolean {
        const address = colCache.decodeEx(addressStr) as { row: number; col: number; sheetName?: string | undefined };
        return this.containsEx(address);
    }

    /** Whether the range contains a cell (a cell of another named sheet is not contained). */
    containsEx(address: { readonly row: number; readonly col: number; readonly sheetName?: string | undefined }): boolean {
        if (address.sheetName && this.sheetName && address.sheetName !== this.sheetName) return false;
        return address.row >= this.top && address.row <= this.bottom && address.col >= this.left && address.col <= this.right;
    }

    /** Calls `cb` for each cell, column by column (top to bottom in each column). */
    forEachAddress(cb: (address: string, row: number, col: number) => void): void {
        for (let col = this.left; col <= this.right; col++) {
            for (let row = this.top; row <= this.bottom; row++) {
                cb(colCache.encodeAddress(row, col), row, col);
            }
        }
    }
}

export default Range;
