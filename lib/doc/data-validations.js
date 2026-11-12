const colCache = require("../utils/col-cache");

const { hasOwnProperty } = Object.prototype;

const isRangeKey = (key) => key.includes(":");

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
        this._update(address);
        return (this._model[address] = validation);
    }

    find(address) {
        if (hasOwnProperty.call(this._model, address)) {
            return this._model[address];
        }
        // Validations may also be keyed by range ('A1:B10'): large ranges read from a
        // file are not expanded per cell, and apps can add ranges directly.
        const match = this._findRange(address);
        return match && this._model[match.key];
    }

    remove(address) {
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

module.exports = DataValidations;
