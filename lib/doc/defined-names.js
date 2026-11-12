"use strict";

const _ = require("../utils/under-dash");
const colCache = require("../utils/col-cache");
const CellMatrix = require("../utils/cell-matrix");
const Range = require("./range");

const rangeRegexp = /[$](\w+)[$](\d+)(:[$](\w+)[$](\d+))?/;

// Ranges are stored one entry per cell (so that names follow row/column splices).
// Ranges larger than this are kept as a single rectangle instead, so that a tiny
// definedName such as Sheet!$A$1:$XFD$1048576 cannot exhaust memory.
const MAX_EXPANDED_CELLS = 100000;

// shift a [start, end] interval after `numDelete` lines were removed and `numInsert`
// inserted at `start`. Returns undefined when the whole interval was deleted.
function spliceInterval(lo, hi, start, numDelete, numInsert) {
    const delta = numInsert - numDelete;
    const deletedEnd = start + numDelete - 1;
    const shift = (n) => (n < start ? n : n + delta);
    if (numDelete > 0) {
        if (lo >= start && hi <= deletedEnd) {
            return undefined;
        }
        const newLo = lo >= start && lo <= deletedEnd ? start + numInsert : shift(lo);
        const newHi = hi >= start && hi <= deletedEnd ? start - 1 : shift(hi);
        return [newLo, newHi];
    }
    return [shift(lo), shift(hi)];
}

class DefinedNames {
    constructor() {
        this.matrixMap = Object.create(null);
        // name -> [{sheetName, top, left, bottom, right}] for ranges too large to expand
        this.rangeMap = Object.create(null);
    }

    getMatrix(name) {
        const matrix = this.matrixMap[name] || (this.matrixMap[name] = new CellMatrix());
        return matrix;
    }

    // add a name to a cell. locStr in the form SheetName!$col$row or SheetName!$c1$r1:$c2:$r2
    add(locStr, name) {
        const location = colCache.decodeEx(locStr);
        this.addEx(location, name);
    }

    addEx(location, name) {
        const matrix = this.getMatrix(name);
        if (location.top && (1 + location.bottom - location.top) * (1 + location.right - location.left) > MAX_EXPANDED_CELLS) {
            const { sheetName, top, left, bottom, right } = location;
            (this.rangeMap[name] || (this.rangeMap[name] = [])).push({ sheetName, top, left, bottom, right });
            return;
        }
        if (location.top) {
            for (let col = location.left; col <= location.right; col++) {
                for (let row = location.top; row <= location.bottom; row++) {
                    const address = {
                        sheetName: location.sheetName,
                        address: colCache.n2l(col) + row,
                        row,
                        col,
                    };

                    matrix.addCellEx(address);
                }
            }
        } else {
            matrix.addCellEx(location);
        }
    }

    remove(locStr, name) {
        const location = colCache.decodeEx(locStr);
        this.removeEx(location, name);
    }

    removeEx(location, name) {
        const matrix = this.getMatrix(name);
        matrix.removeCellEx(location);
        this._removeFromRanges(location, name);
    }

    removeAllNames(location) {
        _.each(this.matrixMap, (matrix) => {
            matrix.removeCellEx(location);
        });
        Object.keys(this.rangeMap).forEach((name) => this._removeFromRanges(location, name));
    }

    // large ranges are atomic: removing any cell of one removes the whole range
    _removeFromRanges(location, name) {
        const ranges = this.rangeMap[name];
        if (ranges) {
            this.rangeMap[name] = ranges.filter((range) => !this._rangeContains(range, location));
        }
    }

    _rangeContains(range, address) {
        return (
            range.sheetName === address.sheetName &&
            address.row >= range.top &&
            address.row <= range.bottom &&
            address.col >= range.left &&
            address.col <= range.right
        );
    }

    // note: ranges kept in rangeMap (too large to expand) are not enumerated cell by cell
    forEach(callback) {
        _.each(this.matrixMap, (matrix, name) => {
            matrix.forEach((cell) => {
                callback(name, cell);
            });
        });
    }

    // get all the names of a cell
    getNames(addressStr) {
        return this.getNamesEx(colCache.decodeEx(addressStr));
    }

    getNamesEx(address) {
        const names = _.map(this.matrixMap, (matrix, name) => matrix.findCellEx(address) && name).filter(Boolean);
        _.each(this.rangeMap, (ranges, name) => {
            if (!names.includes(name) && ranges.some((range) => this._rangeContains(range, address))) {
                names.push(name);
            }
        });
        return names;
    }

    _explore(matrix, cell) {
        cell.mark = false;
        const { sheetName } = cell;

        const range = new Range(cell.row, cell.col, cell.row, cell.col, sheetName);
        let x;
        let y;

        // grow vertical - only one col to worry about
        function vGrow(yy, edge) {
            const c = matrix.findCellAt(sheetName, yy, cell.col);
            if (!c || !c.mark) {
                return false;
            }
            range[edge] = yy;
            c.mark = false;
            return true;
        }
        for (y = cell.row - 1; vGrow(y, "top"); y--);
        for (y = cell.row + 1; vGrow(y, "bottom"); y++);

        // grow horizontal - ensure all rows can grow
        function hGrow(xx, edge) {
            const cells = [];
            for (y = range.top; y <= range.bottom; y++) {
                const c = matrix.findCellAt(sheetName, y, xx);
                if (c && c.mark) {
                    cells.push(c);
                } else {
                    return false;
                }
            }
            range[edge] = xx;
            for (let i = 0; i < cells.length; i++) {
                cells[i].mark = false;
            }
            return true;
        }
        for (x = cell.col - 1; hGrow(x, "left"); x--);
        for (x = cell.col + 1; hGrow(x, "right"); x++);

        return range;
    }

    getRanges(name, matrix) {
        matrix = matrix || this.matrixMap[name];
        const largeRanges = (this.rangeMap[name] || []).map(
            ({ sheetName, top, left, bottom, right }) => new Range(top, left, bottom, right, sheetName).$shortRange
        );

        if (!matrix) {
            return { name, ranges: largeRanges };
        }

        // mark and sweep!
        matrix.forEach((cell) => {
            cell.mark = true;
        });
        const ranges = matrix
            .map((cell) => cell.mark && this._explore(matrix, cell))
            .filter(Boolean)
            .map((range) => range.$shortRange);

        return {
            name,
            ranges: ranges.concat(largeRanges),
        };
    }

    normaliseMatrix(matrix, sheetName) {
        // some of the cells might have shifted on specified sheet
        // need to reassign rows, cols
        matrix.forEachInSheet(sheetName, (cell, row, col) => {
            if (cell) {
                if (cell.row !== row || cell.col !== col) {
                    cell.row = row;
                    cell.col = col;
                    cell.address = colCache.n2l(col) + row;
                }
            }
        });
    }

    spliceRows(sheetName, start, numDelete, numInsert) {
        _.each(this.matrixMap, (matrix) => {
            matrix.spliceRows(sheetName, start, numDelete, numInsert);
            this.normaliseMatrix(matrix, sheetName);
        });
        this._spliceRanges(sheetName, start, numDelete, numInsert, "top", "bottom");
    }

    spliceColumns(sheetName, start, numDelete, numInsert) {
        _.each(this.matrixMap, (matrix) => {
            matrix.spliceColumns(sheetName, start, numDelete, numInsert);
            this.normaliseMatrix(matrix, sheetName);
        });
        this._spliceRanges(sheetName, start, numDelete, numInsert, "left", "right");
    }

    _spliceRanges(sheetName, start, numDelete, numInsert, lo, hi) {
        _.each(this.rangeMap, (ranges, name) => {
            this.rangeMap[name] = ranges
                .map((range) => {
                    if (range.sheetName !== sheetName) {
                        return range;
                    }
                    const interval = spliceInterval(range[lo], range[hi], start, numDelete, numInsert);
                    return interval && { ...range, [lo]: interval[0], [hi]: interval[1] };
                })
                .filter(Boolean);
        });
    }

    get model() {
        // To get names per cell - just iterate over all names finding cells if they exist
        const names = Object.keys(this.matrixMap);
        Object.keys(this.rangeMap).forEach((name) => {
            if (!names.includes(name)) {
                names.push(name);
            }
        });
        return names.map((name) => this.getRanges(name)).filter((definedName) => definedName.ranges.length);
    }

    set model(value) {
        // value is [ { name, ranges }, ... ]
        this.matrixMap = Object.create(null);
        this.rangeMap = Object.create(null);
        value.forEach((definedName) => {
            this.getMatrix(definedName.name);
            definedName.ranges.forEach((rangeStr) => {
                if (rangeRegexp.test(rangeStr.split("!").pop() || "")) {
                    this.add(rangeStr, definedName.name);
                }
            });
        });
    }
}

module.exports = DefinedNames;
