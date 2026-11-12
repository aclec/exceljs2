const colCache = require("../utils/col-cache");

const { hasOwnProperty } = Object.prototype;

const isRangeKey = (key) => key.includes(":");

// the parts of range (top/left/bottom/right) around the cell at row, col: up to 4 rectangles
function splitAround(range, row, col) {
    const { top, left, bottom, right } = range;
    return [
        [top, left, row - 1, right],
        [row, left, row, col - 1],
        [row, col + 1, row, right],
        [row + 1, left, bottom, right],
    ]
        .filter(([t, l, b, r]) => t <= b && l <= r)
        .map(([t, l, b, r]) => (t === b && l === r ? colCache.encodeAddress(t, l) : colCache.encode(t, l, b, r)));
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
        // reload: replace the range by the rectangles around the cell, so that the cell can be overridden or removed.
        const range = this._findRange(address);
        if (range) {
            const validation = this._model[range.key];
            delete this._model[range.key];
            const { row, col } = colCache.decodeAddress(address);
            splitAround(range, row, col).forEach((key) => {
                this._model[key] = validation;
            });
            this._ranges = undefined;
        }
    }

    _findRange(address) {
        const ranges = this._getRanges();
        if (!ranges.length) {
            return undefined;
        }
        const { row, col } = colCache.decodeAddress(address);
        return ranges.find((r) => row >= r.top && row <= r.bottom && col >= r.left && col <= r.right);
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
