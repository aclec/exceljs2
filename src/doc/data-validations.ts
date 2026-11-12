/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import colCache from "../utils/col-cache";
import RangeIndex from "../utils/range-index";

const { hasOwnProperty } = Object.prototype;

const isRangeKey = (key) => key.includes(":");

const wholeColumnsRegex = /^\$?([A-Z]{1,3}):\$?([A-Z]{1,3})$/;
const wholeRowsRegex = /^\$?(\d{1,7}):\$?(\d{1,7})$/;

const isColumn = (letters) => letters.length < 3 || letters <= "XFD";
const isRow = (number) => number >= 1 && number <= colCache.MAX_ROWS;

// 'A:C' (whole columns) and '1:2' (whole rows) as the rectangle they cover ('A1:C1048576', 'A1:XFD2'), the form Excel
// writes in sqref; null for whole columns / rows outside the grid ('A:XFE', '0:0'); other addresses are returned unchanged
function wholeRangeAddress(address) {
    let match = wholeColumnsRegex.exec(address);
    if (match) {
        if (!isColumn(match[1]) || !isColumn(match[2])) {
            return null;
        }
        const [left, right] = [colCache.l2n(match[1]), colCache.l2n(match[2])].sort((a, b) => a - b);
        return colCache.encode(1, left, colCache.MAX_ROWS, right);
    }
    match = wholeRowsRegex.exec(address);
    if (match) {
        const [top, bottom] = [parseInt(match[1], 10), parseInt(match[2], 10)].sort((a, b) => a - b);
        if (!isRow(top) || !isRow(bottom)) {
            return null;
        }
        return colCache.encode(top, 1, bottom, colCache.MAX_COLS);
    }
    return address;
}

function normaliseAddress(address) {
    const normalised = wholeRangeAddress(address);
    if (normalised === null) {
        throw new Error(`Out of bounds. Excel supports columns from A to XFD and rows from 1 to ${colCache.MAX_ROWS}: ${address}`);
    }
    return normalised;
}

// the parts of range (top/left/bottom/right) around the cell at row, col: up to 4 rectangles, the part
// below first (filling a range cell by cell, downwards, then finds the rest of the range first)
function splitAround(range, row, col) {
    const { top, left, bottom, right } = range;
    return [
        [row + 1, left, bottom, right],
        [row, col + 1, row, right],
        [row, left, row, col - 1],
        [top, left, row - 1, right],
    ]
        .filter(([t, l, b, r]) => t <= b && l <= r)
        .map(([t, l, b, r]) =>
            t === b && l === r
                ? { key: colCache.encodeAddress(t, l) }
                : { key: colCache.encode(t, l, b, r), top: t, left: l, bottom: b, right: r }
        );
}

class DataValidations {
    constructor(model) {
        this.model = model || {};
    }

    // Validations may be keyed by range ('A1:B10'): large ranges read from a file are not expanded per cell, and apps can
    // add ranges directly to the model. find() looks them up in an index (utils/range-index), built from the model when
    // first needed and kept up to date by add() and remove(). The model returned here may be changed directly, so reading
    // it drops the index (rebuilt by the next lookup): write to the model through a fresh `dataValidations.model`.
    get model() {
        this._index = undefined;
        return this._model;
    }

    set model(value) {
        this._model = value;
        this._index = undefined;
    }

    add(address, validation) {
        address = normaliseAddress(address);
        this._update(address, validation);
        return (this._model[address] = validation);
    }

    find(address) {
        address = normaliseAddress(address);
        if (hasOwnProperty.call(this._model, address)) {
            return this._model[address];
        }
        const match = this._findRange(address);
        return match && this._model[match.key];
    }

    remove(address) {
        address = normaliseAddress(address);
        this._update(address, undefined);
        this._model[address] = undefined;
    }

    // before the key at address changes to validation
    _update(address, validation) {
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
            this._index.order.delete(range.key);
            const { row, col } = colCache.decodeAddress(address);
            // a key that already exists (from the app or the file) keeps its own validation
            const pieces = splitAround(range, row, col).filter((piece) => !hasOwnProperty.call(this._model, piece.key));
            pieces.forEach((piece) => {
                this._model[piece.key] = rangeValidation;
                if (piece.top !== undefined) {
                    this._setRange(piece.key, rangeValidation);
                }
            });
        }
    }

    // the index entry of the range key, for the validation it gets: none when undefined. Between overlapping range keys,
    // the first one in the model applies.
    _setRange(key, validation) {
        const { ranges, rangeIndex, order } = this._index;
        if (!order.has(key)) {
            order.set(key, this._index.nextOrder++);
        }
        const existing = ranges.get(key);
        if (existing) {
            rangeIndex.remove(existing);
            ranges.delete(key);
        }
        if (validation) {
            const { top, left, bottom, right } = colCache.decode(key);
            const range = { key, top, left, bottom, right };
            rangeIndex.add(range, order.get(key));
            ranges.set(key, range);
        }
    }

    _findRange(address) {
        const { row, col } = colCache.decodeAddress(address);
        return this._getIndex().rangeIndex.find(row, col);
    }

    _getIndex() {
        if (!this._index) {
            // ranges: key => entry of the range index; order: key => position of the key in the model
            this._index = { ranges: new Map(), rangeIndex: new RangeIndex(), order: new Map(), nextOrder: 0 };
            Object.keys(this._model).forEach((key) => {
                if (isRangeKey(key)) {
                    this._setRange(key, this._model[key]);
                }
            });
        }
        return this._index;
    }
}

DataValidations.wholeRangeAddress = wholeRangeAddress;

export default DataValidations;
