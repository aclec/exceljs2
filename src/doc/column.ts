/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import _ from "../utils/under-dash";

import Enums from "./enums";
import colCache from "../utils/col-cache";

const DEFAULT_COLUMN_WIDTH = 9;

// Column defines the column properties for 1 column.
// This includes header rows, widths, key, (style), etc.
// Worksheet will condense the columns as appropriate during serialization
class Column {
    constructor(worksheet, number, defn) {
        this._worksheet = worksheet;
        this._number = number;
        if (defn !== false) {
            // sometimes defn will follow
            this.defn = defn;
        }
    }

    get number() {
        return this._number;
    }

    get worksheet() {
        return this._worksheet;
    }

    get letter() {
        return colCache.n2l(this._number);
    }

    get width() {
        return this._width;
    }

    set width(value) {
        this._width = value;
        // a width assigned explicitly is written even when it equals DEFAULT_COLUMN_WIDTH
        this._implicitWidth = false;
    }

    get isCustomWidth() {
        return this.width !== undefined && this.width !== DEFAULT_COLUMN_WIDTH;
    }

    get defn() {
        return {
            header: this._header,
            key: this.key,
            width: this.width,
            style: this.style,
            hidden: this.hidden,
            outlineLevel: this.outlineLevel,
        };
    }

    set defn(value) {
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

    get headers() {
        return this._header && this._header instanceof Array ? this._header : [this._header];
    }

    get header() {
        return this._header;
    }

    set header(value) {
        if (value !== undefined) {
            this._header = value;
            this.headers.forEach((text, index) => {
                this._worksheet.getCell(index + 1, this.number).value = text;
            });
        } else {
            this._header = undefined;
        }
    }

    get key() {
        return this._key;
    }

    set key(value) {
        const column = this._key && this._worksheet.getColumnKey(this._key);
        if (column === this) {
            this._worksheet.deleteColumnKey(this._key);
        }

        this._key = value;
        if (value) {
            this._worksheet.setColumnKey(this._key, this);
        }
    }

    get hidden() {
        return !!this._hidden;
    }

    set hidden(value) {
        this._hidden = value;
    }

    get outlineLevel() {
        return this._outlineLevel || 0;
    }

    set outlineLevel(value) {
        this._outlineLevel = value;
    }

    get collapsed() {
        return !!(this._outlineLevel && this._outlineLevel >= this._worksheet.properties.outlineLevelCol);
    }

    toString() {
        return JSON.stringify({
            key: this.key,
            width: this.width,
            headers: this.headers.length ? this.headers : undefined,
        });
    }

    equivalentTo(other) {
        return (
            this.width === other.width &&
            this.hidden === other.hidden &&
            this.outlineLevel === other.outlineLevel &&
            _.isEqual(this.style, other.style)
        );
    }

    get isDefault() {
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

    get headerCount() {
        return this.headers.length;
    }

    eachCell(options, iteratee) {
        const colNumber = this.number;
        if (!iteratee) {
            iteratee = options;
            options = null;
        }
        this._worksheet.eachRow(options, (row, rowNumber) => {
            iteratee(row.getCell(colNumber), rowNumber);
        });
    }

    get values() {
        const v = [];
        this.eachCell((cell, rowNumber) => {
            if (cell && cell.type !== Enums.ValueType.Null) {
                v[rowNumber] = cell.value;
            }
        });
        return v;
    }

    set values(v) {
        if (!v) {
            return;
        }
        const colNumber = this.number;
        let offset = 0;
        if (v.hasOwnProperty("0")) {
            // assume contiguous array, start at row 1
            offset = 1;
        }
        v.forEach((value, index) => {
            this._worksheet.getCell(index + offset, colNumber).value = value;
        });
    }

    // =========================================================================
    // styles
    _applyStyle(name, value) {
        // copy on write: the style object may be shared with other rows, columns or cells
        this.style = { ...this.style, [name]: value };
        this.eachCell((cell) => {
            cell[name] = value;
        });
        return value;
    }

    get numFmt() {
        return this.style.numFmt;
    }

    set numFmt(value) {
        this._applyStyle("numFmt", value);
    }

    get font() {
        return this.style.font;
    }

    set font(value) {
        this._applyStyle("font", value);
    }

    get alignment() {
        return this.style.alignment;
    }

    set alignment(value) {
        this._applyStyle("alignment", value);
    }

    get protection() {
        return this.style.protection;
    }

    set protection(value) {
        this._applyStyle("protection", value);
    }

    get border() {
        return this.style.border;
    }

    set border(value) {
        this._applyStyle("border", value);
    }

    get fill() {
        return this.style.fill;
    }

    set fill(value) {
        this._applyStyle("fill", value);
    }

    // =============================================================================
    // static functions

    static toModel(columns) {
        // Convert array of Column into compressed list cols
        const segments = [];
        if (columns) {
            columns.forEach((column, index) => {
                segments.push({ min: index + 1, max: index + 1, column });
            });
        }
        return Column.segmentsToModel(segments);
    }

    // Compresses runs of equivalent columns into a list of cols. segments: [{min, max, column}] in column
    // order, where column stands for each of the columns min..max
    static segmentsToModel(segments) {
        const cols = [];
        let col = null;
        segments.forEach(({ min, max, column }) => {
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
        });
        return cols.length ? cols : undefined;
    }

    // The columns that fromModel would create, without creating them: {ranges: [{min, max, defn}], length}.
    // Ranges are sorted and disjoint (a column belongs to the first col that covers it); the columns up
    // to length that are in no range are plain columns.
    static rangesFromModel(cols) {
        const ranges = [];
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

    // the definition of column n in colRanges (see rangesFromModel), if any
    static defnInRanges(colRanges, n) {
        const ranges = colRanges ? colRanges.ranges : [];
        let lo = 0;
        let hi = ranges.length - 1;
        while (lo <= hi) {
            const mid = (lo + hi) >> 1;
            if (n < ranges[mid].min) {
                hi = mid - 1;
            } else if (n > ranges[mid].max) {
                lo = mid + 1;
            } else {
                return ranges[mid].defn;
            }
        }
        return undefined;
    }

    static fromModel(worksheet, cols) {
        cols = cols || [];
        const columns = [];
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
            const col = cols[index++];
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
