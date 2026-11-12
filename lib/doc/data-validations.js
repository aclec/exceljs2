const colCache = require("../utils/col-cache");

const { hasOwnProperty } = Object.prototype;

const isRangeKey = (key) => key.includes(":");

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
        // only range keys change the range index: rebuilding it after every add was quadratic
        if (isRangeKey(address)) {
            this._ranges = undefined;
        }
        return (this._model[address] = validation);
    }

    find(address) {
        if (hasOwnProperty.call(this._model, address)) {
            return this._model[address];
        }
        // Validations may also be keyed by range ('A1:B10'): large ranges read from a
        // file are not expanded per cell, and apps can add ranges directly.
        const ranges = this._getRanges();
        if (!ranges.length) {
            return undefined;
        }
        const { row, col } = colCache.decodeAddress(address);
        const match = ranges.find((r) => row >= r.top && row <= r.bottom && col >= r.left && col <= r.right);
        return match && this._model[match.key];
    }

    remove(address) {
        if (isRangeKey(address)) {
            this._ranges = undefined;
        }
        this._model[address] = undefined;
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
