// The data validations of a worksheet (`worksheet.dataValidations`), keyed by cell address or by range.

import colCache from "../utils/col-cache";
import RangeIndex from "../utils/range-index";
import type { DataValidationModel, DataValidationsModel } from "../types";

const isRangeKey = (key: string): boolean => key.includes(":");

// anchored, bounded repetitions: linear on untrusted text
const wholeColumnsRegex = /^\$?([A-Z]{1,3}):\$?([A-Z]{1,3})$/;
const wholeRowsRegex = /^\$?(\d{1,7}):\$?(\d{1,7})$/;

const isColumn = (letters: string): boolean => letters.length < 3 || letters <= "XFD";
const isRow = (number: number): boolean => number >= 1 && number <= colCache.MAX_ROWS;

/**
 * Whole columns (`'A:C'`) and whole rows (`'1:2'`) as the rectangle they cover (`'A1:C1048576'`, `'A1:XFD2'`), the form
 * Excel writes in sqref.
 * @returns the rectangle; null for whole columns / rows outside the grid (`'A:XFE'`, `'0:0'`); any other address
 *   unchanged
 */
export function wholeRangeAddress(address: string): string | null {
    const columns = wholeColumnsRegex.exec(address);
    // both groups are not optional: a match has them
    const [, first, last] = columns || [];
    if (first !== undefined && last !== undefined) {
        if (!isColumn(first) || !isColumn(last)) {
            return null;
        }
        const [a, b] = [colCache.l2n(first), colCache.l2n(last)];
        return colCache.encode(1, Math.min(a, b), colCache.MAX_ROWS, Math.max(a, b));
    }
    const rows = wholeRowsRegex.exec(address);
    const [, from, to] = rows || [];
    if (from !== undefined && to !== undefined) {
        const [a, b] = [parseInt(from, 10), parseInt(to, 10)];
        const top = Math.min(a, b);
        const bottom = Math.max(a, b);
        if (!isRow(top) || !isRow(bottom)) {
            return null;
        }
        return colCache.encode(top, 1, bottom, colCache.MAX_COLS);
    }
    return address;
}

/** @throws Error for whole columns / rows outside the grid */
function normaliseAddress(address: string): string {
    const normalised = wholeRangeAddress(address);
    if (normalised === null) {
        throw new Error(`Out of bounds. Excel supports columns from A to XFD and rows from 1 to ${colCache.MAX_ROWS}: ${address}`);
    }
    return normalised;
}

/** A range key of the model in the range index. */
interface IndexedValidationRange {
    key: string;
    top: number;
    left: number;
    bottom: number;
    right: number;
    order?: number;
    sizeClass?: readonly [number, number];
}

/** A piece of a range split around a cell: a single cell (key only) or a rectangle. */
type RangePiece = { key: string; top?: undefined } | IndexedValidationRange;

/** The index of the range keys of the model. */
interface ValidationIndex {
    /** key => entry of the range index */
    ranges: Map<string, IndexedValidationRange>;
    rangeIndex: RangeIndex<IndexedValidationRange>;
    /** key => position of the key in the model (the first one applies between overlapping ranges) */
    order: Map<string, number>;
    nextOrder: number;
}

/**
 * The parts of `range` around the cell at `row`, `col`: up to 4 rectangles, the part below first (filling a range cell
 * by cell, downwards, then finds the rest of the range first).
 */
function splitAround(range: IndexedValidationRange, row: number, col: number): RangePiece[] {
    const { top, left, bottom, right } = range;
    const parts: [number, number, number, number][] = [
        [row + 1, left, bottom, right],
        [row, col + 1, row, right],
        [row, left, row, col - 1],
        [top, left, row - 1, right],
    ];
    return parts
        .filter(([t, l, b, r]) => t <= b && l <= r)
        .map(([t, l, b, r]) =>
            t === b && l === r
                ? { key: colCache.encodeAddress(t, l) }
                : { key: colCache.encode(t, l, b, r), top: t, left: l, bottom: b, right: r }
        );
}

/**
 * The data validations of a worksheet: by cell address (`'B2'`) or by range (`'B1:B1048576'`).
 *
 * Validations may be keyed by range: large ranges read from a file are not expanded per cell, and apps can add ranges
 * directly to the model. find() looks them up in an index (utils/range-index), built from the model when first needed
 * and kept up to date by add() and remove(). The model returned by the getter may be changed directly, so reading it
 * drops the index (rebuilt by the next lookup): write to the model through a fresh `dataValidations.model`.
 */
class DataValidations {
    /** {@link wholeRangeAddress} (used by the data validations xform). */
    static wholeRangeAddress = wholeRangeAddress;

    declare _model: DataValidationsModel;
    declare _index: ValidationIndex | undefined;

    /** @param model - the validations (by reference); a new `{}` when there is none */
    constructor(model?: DataValidationsModel) {
        this.model = model || {};
    }

    /** The validations by key (by reference; reading it drops the index, see the class). */
    get model(): DataValidationsModel {
        this._index = undefined;
        return this._model;
    }

    set model(value: DataValidationsModel) {
        this._model = value;
        this._index = undefined;
    }

    /**
     * Sets the validation of a cell, a range (`'A1:B10'`) or whole columns or rows (`'A:A'`, `'1:3'`). A cell inside a
     * range key splits the range around it.
     * @returns the validation
     * @throws Error for whole columns / rows outside the grid
     */
    add(address: string, validation: DataValidationModel): DataValidationModel {
        address = normaliseAddress(address);
        this._update(address, validation);
        return (this._model[address] = validation);
    }

    /**
     * The validation of a cell, from its own key or from a range key that contains it.
     * @throws Error for whole columns / rows outside the grid
     */
    find(address: string): DataValidationModel | undefined {
        address = normaliseAddress(address);
        if (Object.hasOwn(this._model, address)) {
            return this._model[address];
        }
        const match = this._findRange(address);
        return match && this._model[match.key];
    }

    /**
     * Removes the validation of a cell (splitting a range key around it) or a range key: the key is kept, `undefined`.
     * @throws Error for whole columns / rows outside the grid
     */
    remove(address: string): void {
        address = normaliseAddress(address);
        this._update(address, undefined);
        this._model[address] = undefined;
    }

    /** Before the key at `address` changes to `validation`: keeps the index up to date, splits a range around a cell. */
    _update(address: string, validation: DataValidationModel | undefined): void {
        if (isRangeKey(address)) {
            if (this._index) {
                this._setRange(address, validation);
            }
            return;
        }
        // A cell inside a range key would overlap it in the written file, where the range would apply again on
        // reload: replace the range containing the cell (the one find() uses) by the rectangles around it, so
        // that the cell can be overridden or removed. Overlapping range keys (which Excel does not write) are
        // not split further.
        const range = this._findRange(address);
        if (range) {
            const rangeValidation = this._model[range.key];
            delete this._model[range.key];
            this._setRange(range.key, undefined);
            this._getIndex().order.delete(range.key);
            const { row, col } = colCache.decodeAddress(address);
            // a key that already exists (from the app or the file) keeps its own validation
            const pieces = splitAround(range, row, col).filter((piece) => !Object.hasOwn(this._model, piece.key));
            for (const piece of pieces) {
                this._model[piece.key] = rangeValidation;
                if (piece.top !== undefined) {
                    this._setRange(piece.key, rangeValidation);
                }
            }
        }
    }

    /**
     * Sets the index entry of a range key for the validation it gets: none when undefined. Between overlapping range
     * keys, the first one in the model applies.
     */
    _setRange(key: string, validation: DataValidationModel | undefined): void {
        const index = this._getIndex();
        const { ranges, rangeIndex, order } = index;
        let orderNo = order.get(key);
        if (orderNo === undefined) {
            orderNo = index.nextOrder++;
            order.set(key, orderNo);
        }
        const existing = ranges.get(key);
        if (existing) {
            rangeIndex.remove(existing);
            ranges.delete(key);
        }
        if (validation) {
            // a range key (it has a colon); one of a single cell (`A1:A1`) is the rectangle of that cell
            const decoded = colCache.decode(key);
            const range =
                "top" in decoded
                    ? { key, top: decoded.top, left: decoded.left, bottom: decoded.bottom, right: decoded.right }
                    : { key, top: decoded.row, left: decoded.col, bottom: decoded.row, right: decoded.col };
            rangeIndex.add(range, orderNo);
            ranges.set(key, range);
        }
    }

    /** The range key that contains a cell (the first one in the model), if any. */
    _findRange(address: string): IndexedValidationRange | undefined {
        const { row, col } = colCache.decodeAddress(address);
        return this._getIndex().rangeIndex.find(row, col);
    }

    /** The index of the range keys, built from the model when needed. */
    _getIndex(): ValidationIndex {
        if (!this._index) {
            this._index = { ranges: new Map(), rangeIndex: new RangeIndex(), order: new Map(), nextOrder: 0 };
            for (const key of Object.keys(this._model)) {
                if (isRangeKey(key)) {
                    this._setRange(key, this._model[key]);
                }
            }
        }
        return this._index;
    }
}

export default DataValidations;
