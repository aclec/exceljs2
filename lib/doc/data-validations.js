const colCache = require("../utils/col-cache");

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

    get model() {
        return this._model;
    }

    set model(value) {
        this._model = value;
        this._ranges = undefined;
    }

    add(address, validation) {
        address = normaliseAddress(address);
        this._update(address);
        return (this._model[address] = validation);
    }

    find(address) {
        address = normaliseAddress(address);
        if (hasOwnProperty.call(this._model, address)) {
            return this._model[address];
        }
        // Validations may also be keyed by range ('A1:B10'): large ranges read from a
        // file are not expanded per cell, and apps can add ranges directly.
        const match = this._findRange(address);
        return match && this._model[match.key];
    }

    remove(address) {
        address = normaliseAddress(address);
        this._update(address);
        this._model[address] = undefined;
    }

    // before the key at address changes
    _update(address) {
        if (isRangeKey(address)) {
            // only range keys change the range index: rebuilding it after every add was quadratic
            this._ranges = undefined;
            return;
        }
        // A cell inside a range key would overlap it in the written file, where the range would apply again on
        // reload: replace the range containing the cell (the one find() uses) by the rectangles around it, so
        // that the cell can be overridden or removed. Overlapping range keys (which Excel does not write) are
        // not split further. The range index is updated in place: rebuilding it would make filling the cells
        // of a whole-column range one by one quadratic.
        const ranges = this._getRanges();
        const index = ranges.length ? this._findRangeIndex(address) : -1;
        if (index !== -1) {
            const range = ranges[index];
            const validation = this._model[range.key];
            delete this._model[range.key];
            const { row, col } = colCache.decodeAddress(address);
            // a key that already exists (from the app or the file) keeps its own validation
            const pieces = splitAround(range, row, col).filter((piece) => !hasOwnProperty.call(this._model, piece.key));
            pieces.forEach((piece) => {
                this._model[piece.key] = validation;
            });
            ranges.splice(index, 1, ...pieces.filter((piece) => piece.top !== undefined));
        }
    }

    _findRangeIndex(address) {
        const { row, col } = colCache.decodeAddress(address);
        return this._getRanges().findIndex((r) => row >= r.top && row <= r.bottom && col >= r.left && col <= r.right);
    }

    _findRange(address) {
        const ranges = this._getRanges();
        return ranges.length ? ranges[this._findRangeIndex(address)] : undefined;
    }

    _getRanges() {
        if (!this._ranges) {
            this._ranges = Object.keys(this._model)
                .filter((key) => isRangeKey(key) && this._model[key])
                .map((key) => ({ key, ...colCache.decode(key) }));
        }
        return this._ranges;
    }
}

DataValidations.wholeRangeAddress = wholeRangeAddress;

module.exports = DataValidations;
