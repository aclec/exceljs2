const _ = require("../utils/under-dash");

const colCache = require("../utils/col-cache");
const Range = require("./range");
const Row = require("./row");
const Column = require("./column");
const Enums = require("./enums");
const Image = require("./image");
const Table = require("./table");
const DataValidations = require("./data-validations");
const { makePivotTable } = require("./pivot-table");
const Encryptor = require("../utils/encryptor");
const { copyStyle } = require("../utils/copy-style");
const { checkMergedCells } = require("../utils/merge-limits");
const { eachEntry, entryIndices, lastEntryIndex, shiftSteps } = require("../utils/sparse-array");

// worksheet.tables is a plain object (by table name). Names come from loaded files, so it is only
// read and written through own properties: "__proto__" or "constructor" are valid table names.
function setTable(tables, name, table) {
    Object.defineProperty(tables, name, { value: table, enumerable: true, writable: true, configurable: true });
}

// The span [from, to] of a merge's rows (or columns) after a splice removed count of them at start and inserted nInserts
// there: the moved span of its remaining rows, which also covers rows inserted between two remaining ones (as Excel widens
// a merge when rows are inserted inside it). undefined when none of its rows remains.
function spliceSpan(from, to, start, count, nInserts) {
    const end = start + count; // first row after the removed ones
    const first = from >= start && from < end ? end : from;
    const last = to >= start && to < end ? start - 1 : to;
    if (first > last) {
        return undefined;
    }
    const move = (n) => (n < start ? n : n + nInserts - count);
    return [move(first), move(last)];
}

// Excel widens a merge over the rows or columns inserted inside it (spliceSpan), and refuses to change part of a merged cell:
// only the master of a merge holds a value, so a value inserted in another cell of the widened merge would be lost. Splices
// refuse such values, before anything changes.
function checkInsertedValue(value, row, col, range) {
    if (value !== undefined && value !== null && value !== "") {
        throw new Error(
            `Cannot insert a value in ${colCache.encodeAddress(row, col)}: it would be inside the merged cells ${range.range}, unmerge them first`
        );
    }
}

// the value that row.values = values puts in column col (see Row#values)
function insertedRowValue(worksheet, values, col) {
    if (!values) {
        return undefined;
    }
    if (Array.isArray(values)) {
        return values[Object.prototype.hasOwnProperty.call(values, "0") ? col - 1 : col];
    }
    const column = worksheet._columns && worksheet._columns[col - 1];
    const key = column && column.key;
    return key !== undefined && Object.prototype.hasOwnProperty.call(values, key) ? values[key] : undefined;
}

// Worksheet requirements
//  Operate as sheet inside workbook or standalone
//  Load and Save from file and stream
//  Access/Add/Delete individual cells
//  Manage column widths and row heights

class Worksheet {
    constructor(options) {
        options = options || {};
        this._workbook = options.workbook;

        // in a workbook, each sheet will have a number
        this.id = options.id;
        this.orderNo = options.orderNo;

        // and a name
        this.name = options.name;

        // add a state
        this.state = options.state || "visible";

        // rows allows access organised by row. Sparse array of arrays indexed by row-1, col
        // Note: _rows is zero based. Must subtract 1 to go from cell.row to index
        this._rows = [];

        // column definitions: Column objects indexed by number - 1, created on demand (see getColumn), so the
        // array may have holes; ranges of loaded cols ({ranges, length}, see Column.rangesFromModel) whose
        // columns are not created yet. A single cell in column XFD or a <col min="1" max="16384"> on each
        // worksheet used to create 16384 Column objects per worksheet.
        this._columns = null;
        this._colRanges = null;
        // true when this._columns may have holes or this._colRanges columns not yet created
        this._columnsSparse = false;

        // column keys (addRow convenience): key ==> this._collumns index
        // no prototype: keys come from data (e.g. CSV headers) and may be '__proto__' or 'constructor'
        this._keys = Object.create(null);

        // keep record of all merges
        this._merges = {};
        this._mergedCellCount = 0;

        // record of all row and column pageBreaks
        this.rowBreaks = [];

        // for tabColor, default row height, outline levels, etc
        this.properties = Object.assign(
            {},
            {
                defaultRowHeight: 15,
                dyDescent: 55,
                outlineLevelCol: 0,
                outlineLevelRow: 0,
            },
            options.properties
        );

        // for all things printing
        this.pageSetup = Object.assign(
            {},
            {
                margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
                orientation: "portrait",
                horizontalDpi: 4294967295,
                verticalDpi: 4294967295,
                fitToPage: !!(
                    options.pageSetup &&
                    (options.pageSetup.fitToWidth || options.pageSetup.fitToHeight) &&
                    !options.pageSetup.scale
                ),
                pageOrder: "downThenOver",
                blackAndWhite: false,
                draft: false,
                cellComments: "None",
                errors: "displayed",
                scale: 100,
                fitToWidth: 1,
                fitToHeight: 1,
                paperSize: undefined,
                showRowColHeaders: false,
                showGridLines: false,
                firstPageNumber: undefined,
                horizontalCentered: false,
                verticalCentered: false,
                rowBreaks: null,
                colBreaks: null,
            },
            options.pageSetup
        );

        this.headerFooter = Object.assign(
            {},
            {
                differentFirst: false,
                differentOddEven: false,
                oddHeader: null,
                oddFooter: null,
                evenHeader: null,
                evenFooter: null,
                firstHeader: null,
                firstFooter: null,
            },
            options.headerFooter
        );

        this.dataValidations = new DataValidations();

        // for freezepanes, split, zoom, gridlines, etc
        this.views = options.views || [];

        this.autoFilter = options.autoFilter || null;

        // for images, etc
        this._media = [];

        // worksheet protection
        this.sheetProtection = null;

        // for tables
        this.tables = {}; // by table name, see setTable

        this.pivotTables = [];

        this.conditionalFormattings = [];
    }

    get name() {
        return this._name;
    }

    set name(name) {
        if (name === undefined) {
            name = `sheet${this.id}`;
        }

        if (this._name === name) return;

        if (typeof name !== "string") {
            throw new Error("The name has to be a string.");
        }

        if (name === "") {
            throw new Error("The name can't be empty.");
        }

        if (name === "History") {
            throw new Error('The name "History" is protected. Please use a different name.');
        }

        // Illegal character in worksheet name: asterisk (*), question mark (?),
        // colon (:), forward slash (/ \), or bracket ([])
        if (/[*?:/\\[\]]/.test(name)) {
            throw new Error(`Worksheet name ${name} cannot include any of the following characters: * ? : \\ / [ ]`);
        }

        if (/(^')|('$)/.test(name)) {
            throw new Error(`The first or last character of worksheet name cannot be a single quotation mark: ${name}`);
        }

        if (name && name.length > 31) {
            console.warn(`Worksheet name ${name} exceeds 31 chars. This will be truncated`);
            name = name.substring(0, 31);
        }

        const lowerName = name.toLowerCase();
        if (this._workbook._worksheets.find((ws) => ws && ws !== this && ws.name.toLowerCase() === lowerName)) {
            throw new Error(`Worksheet name already exists: ${name}`);
        }

        this._name = name;
    }

    get workbook() {
        return this._workbook;
    }

    // when you're done with this worksheet, call this to remove from workbook
    destroy() {
        this._workbook.removeWorksheetEx(this);
    }

    // Get the bounding range of the cells in this worksheet
    get dimensions() {
        const dimensions = new Range();
        eachEntry(this._rows, (row) => {
            const rowDims = row.dimensions;
            if (rowDims) {
                dimensions.expand(row.number, rowDims.min, row.number, rowDims.max);
            }
        });
        return dimensions;
    }

    // =========================================================================
    // Columns

    // get the current columns array.
    get columns() {
        if (this._columnsSparse) {
            // fills the holes in place: an array returned earlier stays the columns array
            const length = this._columnsLength();
            if (length) {
                const columns = this._columns || (this._columns = []);
                for (let n = 1; n <= length; n++) {
                    if (!columns[n - 1]) {
                        columns[n - 1] = new Column(this, n, this._columnDefn(n));
                    }
                }
            }
            this._colRanges = null;
            this._columnsSparse = false;
        }
        return this._columns;
    }

    _columnsLength() {
        return Math.max(this._columns ? this._columns.length : 0, this._colRanges ? this._colRanges.length : 0);
    }

    // the definition of column n in the loaded ranges, if any
    _columnDefn(n) {
        const ranges = this._colRanges ? this._colRanges.ranges : [];
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

    // the cols of the model, in time proportional to the created columns and loaded ranges
    _colsModel() {
        if (!this._columnsSparse) {
            return Column.toModel(this._columns);
        }
        const columns = this._columns || [];
        const created = Object.keys(columns).map((key) => Number(key) + 1);
        const ranges = this._colRanges ? this._colRanges.ranges : [];
        const length = this._columnsLength();
        const segments = [];
        let plainColumn;
        let i = 0;
        let r = 0;
        let n = 1;
        while (n <= length) {
            if (created[i] === n) {
                segments.push({ min: n, max: n, column: columns[n - 1] });
                i++;
                n++;
            } else {
                // columns n.. up to the next created column stand for the same range, or are plain columns
                const nextCreated = i < created.length ? created[i] : length + 1;
                while (r < ranges.length && ranges[r].max < n) {
                    r++;
                }
                const range = ranges[r];
                let max;
                let column;
                if (range && range.min <= n) {
                    max = Math.min(range.max, nextCreated - 1);
                    column = range.column || (range.column = new Column(this, range.min, range.defn));
                } else {
                    max = Math.min(range ? range.min - 1 : length, nextCreated - 1);
                    column = plainColumn || (plainColumn = new Column(this, n));
                }
                segments.push({ min: n, max, column });
                n = max + 1;
            }
        }
        return Column.segmentsToModel(segments);
    }

    // set the columns from an array of column definitions.
    // Note: any headers defined will overwrite existing values.
    set columns(value) {
        // calculate max header row count
        this._headerRowCount = value.reduce((pv, cv) => {
            const headerCount = (cv.header && 1) || (cv.headers && cv.headers.length) || 0;
            return Math.max(pv, headerCount);
        }, 0);

        // construct Column objects
        let count = 1;
        this._colRanges = null;
        this._columnsSparse = false;
        const columns = (this._columns = []);
        value.forEach((defn) => {
            const column = new Column(this, count++, false);
            columns.push(column);
            column.defn = defn;
        });
    }

    getColumnKey(key) {
        return this._keys[key];
    }

    setColumnKey(key, value) {
        this._keys[key] = value;
    }

    deleteColumnKey(key) {
        delete this._keys[key];
    }

    eachColumnKey(f) {
        _.each(this._keys, f);
    }

    // get a single column by col number. If it doesn't exist, create it and any gaps before it
    getColumn(c) {
        if (typeof c === "string") {
            // if it matches a key'd column, return that
            const col = this._keys[c];
            if (col) return col;

            // otherwise, assume letter
            c = colCache.l2n(c);
        }
        if (!this._columns) {
            this._columns = [];
        }
        if (!Number.isInteger(c) || c < 1) {
            return undefined;
        }
        let column = this._columns[c - 1];
        if (!column) {
            // only this column is created: the ones before it are created when they are used, or by
            // the columns getter
            if (c - 1 !== this._columns.length) {
                this._columnsSparse = true;
            }
            column = this._columns[c - 1] = new Column(this, c, this._columnDefn(c));
        }
        return column;
    }

    // column c if it exists, without creating it (and the ones before it) like getColumn: a column of a
    // loaded range that was not created yet is returned as a detached Column
    _findColumn(c) {
        const column = this._columns && this._columns[c - 1];
        if (column) {
            return column;
        }
        const defn = Number.isInteger(c) && c >= 1 ? this._columnDefn(c) : undefined;
        return defn ? new Column(this, c, defn) : undefined;
    }

    spliceColumns(start, count, ...inserts) {
        const merges = this._planMerges(
            (merge) => merge.right >= start,
            (merge) => {
                const span = spliceSpan(merge.left, merge.right, start, count, inserts.length);
                return span && [merge.top, span[0], merge.bottom, span[1]];
            }
        );
        merges.moved.forEach((range) => {
            for (let col = Math.max(start, range.left + 1); col <= Math.min(start + inserts.length - 1, range.right); col++) {
                const values = inserts[col - start];
                for (let row = range.top; row <= range.bottom; row++) {
                    checkInsertedValue(values && values[row - 1], row, col, range);
                }
            }
        });
        this._detachMerges(merges.touched);
        const rows = this._rows;
        const nRows = rows.length;
        if (inserts.length > 0) {
            // the rows that exist and the ones an insert gives a value: a missing row would only get empty cells
            const indices = new Set(entryIndices(rows));
            inserts.forEach((insert) => {
                eachEntry(insert, (value, i) => {
                    if (value && i < nRows) {
                        indices.add(i);
                    }
                });
            });
            Array.from(indices)
                .sort((a, b) => a - b)
                .forEach((i) => {
                    const rowArguments = [start, count];

                    inserts.forEach((insert) => {
                        rowArguments.push(insert[i] || null);
                    });
                    const row = this.getRow(i + 1);

                    row.splice.apply(row, rowArguments);
                });
        } else {
            // nothing to insert, so just splice all rows
            eachEntry(rows, (r) => {
                r.splice(start, count);
            });
        }

        // splice column definitions
        const nExpand = inserts.length - count;
        const nKeep = start + count;
        const nEnd = this.columns.length;
        if (nExpand < 0) {
            for (let i = start + inserts.length; i <= nEnd; i++) {
                this.getColumn(i).defn = this.getColumn(i - nExpand).defn;
            }
        } else if (nExpand > 0) {
            for (let i = nEnd; i >= nKeep; i--) {
                this.getColumn(i + nExpand).defn = this.getColumn(i).defn;
            }
        }
        for (let i = start; i < start + inserts.length; i++) {
            this.getColumn(i).defn = null;
        }

        this._attachMerges(merges.moved);

        // account for defined names
        this.workbook.definedNames.spliceColumns(this.name, start, count, inserts.length);
    }

    get lastColumn() {
        return this.getColumn(this.columnCount);
    }

    get columnCount() {
        let maxCount = 0;
        this.eachRow((row) => {
            maxCount = Math.max(maxCount, row.cellCount);
        });
        return maxCount;
    }

    get actualColumnCount() {
        // performance nightmare - for each row, counts all the columns used
        const counts = [];
        let count = 0;
        this.eachRow((row) => {
            row.eachCell(({ col }) => {
                if (!counts[col]) {
                    counts[col] = true;
                    count++;
                }
            });
        });
        return count;
    }

    // =========================================================================
    // Rows

    _commitRow() {
        // nop - allows streaming reader to fill a document
    }

    get _lastRowNumber() {
        // need to cope with results of splice
        return lastEntryIndex(this._rows) + 1;
    }

    get _nextRow() {
        return this._lastRowNumber + 1;
    }

    get lastRow() {
        if (this._rows.length) {
            return this._rows[this._rows.length - 1];
        }
        return undefined;
    }

    // find a row (if exists) by row number
    findRow(r) {
        return this._rows[r - 1];
    }

    // find multiple rows (if exists) by row number
    findRows(start, length) {
        return this._rows.slice(start - 1, start - 1 + length);
    }

    get rowCount() {
        return this._lastRowNumber;
    }

    get actualRowCount() {
        // counts actual rows that have actual data
        let count = 0;
        this.eachRow(() => {
            count++;
        });
        return count;
    }

    // get a row by row number.
    getRow(r) {
        let row = this._rows[r - 1];
        if (!row) {
            row = this._rows[r - 1] = new Row(this, r);
        }
        return row;
    }

    // get multiple rows by row number.
    getRows(start, length) {
        if (length < 1) return undefined;
        const rows = [];
        for (let i = start; i < start + length; i++) {
            rows.push(this.getRow(i));
        }
        return rows;
    }

    addRow(value, style = "n") {
        const rowNo = this._nextRow;
        const row = this.getRow(rowNo);
        row.values = value;
        this._setStyleOption(rowNo, style[0] === "i" ? style : "n");
        return row;
    }

    addRows(value, style = "n") {
        const rows = [];
        value.forEach((row) => {
            rows.push(this.addRow(row, style));
        });
        return rows;
    }

    insertRow(pos, value, style = "n") {
        this.spliceRows(pos, 0, value);
        this._setStyleOption(pos, style);
        return this.getRow(pos);
    }

    insertRows(pos, values, style = "n") {
        this.spliceRows(pos, 0, ...values);
        if (style !== "n") {
            // copy over the styles
            for (let i = 0; i < values.length; i++) {
                if (style[0] === "o" && this.findRow(values.length + pos + i) !== undefined) {
                    this._copyStyle(values.length + pos + i, pos + i, style[1] === "+");
                } else if (style[0] === "i" && this.findRow(pos - 1) !== undefined) {
                    this._copyStyle(pos - 1, pos + i, style[1] === "+");
                }
            }
        }
        return this.getRows(pos, values.length);
    }

    // set row at position to same style as of either pervious row (option 'i') or next row (option 'o')
    _setStyleOption(pos, style = "n") {
        if (style[0] === "o" && this.findRow(pos + 1) !== undefined) {
            this._copyStyle(pos + 1, pos, style[1] === "+");
        } else if (style[0] === "i" && this.findRow(pos - 1) !== undefined) {
            this._copyStyle(pos - 1, pos, style[1] === "+");
        }
    }

    _copyStyle(src, dest, styleEmpty = false) {
        const rSrc = this.getRow(src);
        const rDst = this.getRow(dest);
        rDst.style = copyStyle(rSrc.style);

        rSrc.eachCell({ includeEmpty: styleEmpty }, (cell, colNumber) => {
            rDst.getCell(colNumber).style = copyStyle(cell.style);
        });
        rDst.height = rSrc.height;
    }

    duplicateRow(rowNum, count, insert = false) {
        // create count duplicates of rowNum
        // either inserting new or overwriting existing rows

        // a row that does not exist (spliceColumns no longer creates the empty rows) is duplicated as an empty row
        const rSrc = this.getRow(rowNum);
        const inserts = new Array(count).fill(rSrc.values);
        // the single-row merges of the row are copied too
        const rowMerges = Object.values(this._merges).filter((merge) => merge.top === rowNum && merge.bottom === rowNum);
        const copiedCells = rowMerges.reduce((cells, merge) => cells + merge.count, 0) * count;
        // the values of the row in a merge are the master's: the copies lose nothing when the merge widens over them
        this._spliceRows(rowNum + 1, insert ? 0 : count, inserts, copiedCells, true);

        // now copy styles...
        for (let i = 0; i < count; i++) {
            const rDst = this._rows[rowNum + i];
            rDst.style = rSrc.style;
            rDst.height = rSrc.height;

            rSrc.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                rDst.getCell(colNumber).style = cell.style;
            });
        }

        // and the merges of the row
        rowMerges.forEach((merge) => {
            for (let i = 1; i <= count; i++) {
                this._mergeCellsInternal(new Range(rowNum + i, merge.left, rowNum + i, merge.right), true);
            }
        });
    }

    spliceRows(start, count, ...inserts) {
        this._spliceRows(start, count, inserts);
    }

    // extraMergedCells: merged cells the caller adds after the splice (checked against the cap with the moved merges)
    // copiesOfMergedRow: the inserts are copies of a row, whose values inside a merge are the master's
    _spliceRows(start, count, inserts, extraMergedCells = 0, copiesOfMergedRow = false) {
        // same problem as row.splice, except worse.
        const merges = this._planMerges(
            (merge) => merge.bottom >= start,
            (merge) => {
                const span = spliceSpan(merge.top, merge.bottom, start, count, inserts.length);
                return span && [span[0], merge.left, span[1], merge.right];
            },
            extraMergedCells
        );
        if (!copiesOfMergedRow) {
            merges.moved.forEach((range) => {
                for (let row = Math.max(start, range.top + 1); row <= Math.min(start + inserts.length - 1, range.bottom); row++) {
                    for (let col = range.left; col <= range.right; col++) {
                        checkInsertedValue(insertedRowValue(this, inserts[row - start], col), row, col, range);
                    }
                }
            });
        }
        this._detachMerges(merges.touched);
        const nKeep = start + count;
        const nInserts = inserts.length;
        const nExpand = nInserts - count;
        const nEnd = this._rows.length;
        // the loops below from nKeep to nEnd only visit the rows i whose source (i) or destination (i + nExpand)
        // exists: the other steps do nothing, and most of them are empty when the last row is far
        const existing = entryIndices(this._rows).map((index) => index + 1);
        let i;
        let rSrc;
        if (nExpand < 0) {
            // remove rows
            for (i of shiftSteps(existing, nKeep, nEnd, nExpand)) {
                rSrc = this._rows[i - 1];
                if (rSrc) {
                    const rDst = this.getRow(i + nExpand);
                    rDst.values = rSrc.values;
                    rDst.style = rSrc.style;
                    rDst.height = rSrc.height;

                    rSrc.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                        const cDst = rDst.getCell(colNumber);
                        cDst.style = cell.style;
                        cDst._comment = cell._comment;
                    });
                    this._rows[i - 1] = undefined;
                } else {
                    this._rows[i + nExpand - 1] = undefined;
                }
            }
            // clear removed rows that nothing was moved or inserted into (deleting near the end);
            // bounded by nEnd so that deleting past the end does not grow _rows
            const nClearStart = Math.max(start + nInserts, nEnd + nExpand + 1);
            const nClearEnd = Math.min(nKeep - 1, nEnd);
            existing.forEach((r) => {
                if (r >= nClearStart && r <= nClearEnd) {
                    this._rows[r - 1] = undefined;
                }
            });
        } else if (nExpand > 0) {
            // insert new cells
            for (i of shiftSteps(existing, nKeep, nEnd, nExpand).reverse()) {
                rSrc = this._rows[i - 1];
                if (rSrc) {
                    const rDst = this.getRow(i + nExpand);
                    rDst.values = rSrc.values;
                    rDst.style = rSrc.style;
                    rDst.height = rSrc.height;

                    rSrc.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                        const cDst = rDst.getCell(colNumber);
                        cDst.style = cell.style;
                        cDst._comment = cell._comment;
                    });
                } else {
                    this._rows[i + nExpand - 1] = undefined;
                }
            }
        }

        // now copy over the new values
        for (i = 0; i < nInserts; i++) {
            const rDst = this.getRow(start + i);
            rDst.style = {};
            rDst.values = inserts[i];
        }

        // drop the empty slots that removed rows left at the end (lastRow, eachRow includeEmpty)
        this._rows.length = this._lastRowNumber;

        this._attachMerges(merges.moved);

        // account for defined names
        this.workbook.definedNames.spliceRows(this.name, start, count, nInserts);
    }

    // iterate over every row in the worksheet, including maybe empty rows
    eachRow(options, iteratee) {
        if (!iteratee) {
            iteratee = options;
            options = undefined;
        }
        if (options && options.includeEmpty) {
            const n = this._rows.length;
            for (let i = 1; i <= n; i++) {
                iteratee(this.getRow(i), i);
            }
        } else {
            eachEntry(this._rows, (row) => {
                if (row.hasValues) {
                    iteratee(row, row.number);
                }
            });
        }
    }

    // return all rows as sparse array
    getSheetValues() {
        const rows = [];
        eachEntry(this._rows, (row) => {
            rows[row.number] = row.values;
        });
        return rows;
    }

    // =========================================================================
    // Cells

    // returns the cell at [r,c] or address given by r. If not found, return undefined
    findCell(r, c) {
        const address = colCache.getAddress(r, c);
        const row = this._rows[address.row - 1];
        return row ? row.findCell(address.col) : undefined;
    }

    // return the cell at [r,c] or address given by r. If not found, create a new one.
    getCell(r, c) {
        const address = colCache.getAddress(r, c);
        const row = this.getRow(address.row);
        return row.getCellEx(address);
    }

    // =========================================================================
    // Merge

    // convert the range defined by ['tl:br'], [tl,br] or [t,l,b,r] into a single 'merged' cell
    mergeCells(...cells) {
        const dimensions = new Range(cells);
        this._mergeCellsInternal(dimensions);
    }

    mergeCellsWithoutStyle(...cells) {
        const dimensions = new Range(cells);
        this._mergeCellsInternal(dimensions, true);
    }

    _mergeCellsInternal(dimensions, ignoreStyle) {
        // Every cell of a merge is materialised, so the total merged area is capped (utils/merge-limits)
        const mergedCellCount = this._mergedCellCount + dimensions.count;
        this._checkMergedCells(`Cannot merge ${dimensions.range}`, mergedCellCount);

        // check cells aren't already merged - O(area) instead of O(number of merges)
        for (let i = dimensions.top; i <= dimensions.bottom; i++) {
            for (let j = dimensions.left; j <= dimensions.right; j++) {
                const cell = this.findCell(i, j);
                if (cell && (cell.isMerged || this._merges[cell.address])) {
                    throw new Error("Cannot merge already merged cells");
                }
            }
        }

        // apply merge
        const master = this.getCell(dimensions.top, dimensions.left);
        for (let i = dimensions.top; i <= dimensions.bottom; i++) {
            for (let j = dimensions.left; j <= dimensions.right; j++) {
                // merge all but the master cell
                if (i > dimensions.top || j > dimensions.left) {
                    this.getCell(i, j).merge(master, ignoreStyle);
                }
            }
        }

        // index merge
        this._merges[master.address] = dimensions;
        this._mergedCellCount = mergedCellCount;
    }

    // Splices move merges as rectangles: the merges that isTouched, and where move(merge) puts each of them
    // ([top, left, bottom, right], or nothing when the splice removes it). A merge reduced to a single cell is not a merge
    // any more. Throws before anything changes when the moved merges (and extraCells more merged cells) would exceed the cap.
    _planMerges(isTouched, move, extraCells = 0) {
        const touched = Object.values(this._merges).filter(isTouched);
        const moved = touched
            .map(move)
            .filter((span) => span && (span[0] !== span[2] || span[1] !== span[3]))
            .map((span) => new Range(...span));
        const count = [
            this._mergedCellCount + extraCells,
            ...touched.map((merge) => -merge.count),
            ...moved.map((range) => range.count),
        ].reduce((sum, n) => sum + n);
        this._checkMergedCells("Cannot splice", count);
        return { touched, moved };
    }

    // throws when the worksheet would hold count merged cells: more than a worksheet or its workbook may hold
    _checkMergedCells(what, count) {
        const workbook = this._workbook;
        if (workbook) {
            const others = workbook._worksheets.reduce((sum, ws) => (ws && ws !== this ? sum + ws._mergedCellCount : sum), 0);
            checkMergedCells(what, count, others + count, workbook.maxMergedCells);
        } else {
            checkMergedCells(what, count, count, Infinity);
        }
    }

    // The merges are released before the cells move, each cell keeping its own value (empty for all but the master)
    // and style, so that rows and columns move as plain cells...
    _detachMerges(merges) {
        merges.forEach((merge) => {
            for (let i = merge.top; i <= merge.bottom; i++) {
                for (let j = merge.left; j <= merge.right; j++) {
                    const cell = this.findCell(i, j);
                    if (cell) {
                        cell.unmerge(true);
                    }
                }
            }
            delete this._merges[merge.tl];
            this._mergedCellCount -= merge.count;
        });
    }

    // ...then merged again at their new place
    _attachMerges(ranges) {
        ranges.forEach((range) => this._mergeCellsInternal(range, true));
    }

    _unMergeMaster(master) {
        // master is always top left of a rectangle
        const merge = this._merges[master.address];
        if (merge) {
            for (let i = merge.top; i <= merge.bottom; i++) {
                for (let j = merge.left; j <= merge.right; j++) {
                    this.getCell(i, j).unmerge();
                }
            }
            delete this._merges[master.address];
            this._mergedCellCount -= merge.count;
        }
    }

    get hasMerges() {
        // return true if this._merges has a merge object
        return _.some(this._merges, Boolean);
    }

    // scan the range defined by ['tl:br'], [tl,br] or [t,l,b,r] and if any cell is part of a merge,
    // un-merge the group. Note this function can affect multiple merges and merge-blocks are
    // atomic - either they're all merged or all un-merged.
    unMergeCells(...cells) {
        const dimensions = new Range(cells);

        // find any cells in that range and unmerge them
        for (let i = dimensions.top; i <= dimensions.bottom; i++) {
            for (let j = dimensions.left; j <= dimensions.right; j++) {
                const cell = this.findCell(i, j);
                if (cell) {
                    if (cell.type === Enums.ValueType.Merge) {
                        // this cell merges to another master
                        this._unMergeMaster(cell.master);
                    } else if (this._merges[cell.address]) {
                        // this cell is a master
                        this._unMergeMaster(cell);
                    }
                }
            }
        }
    }

    // ===========================================================================
    // Shared/Array Formula
    fillFormula(range, formula, results, shareType = "shared") {
        // Define formula for top-left cell and share to rest
        const decoded = colCache.decode(range);
        const { top, left, bottom, right } = decoded;
        const width = right - left + 1;
        const masterAddress = colCache.encodeAddress(top, left);
        const isShared = shareType === "shared";

        // work out result accessor
        let getResult;
        if (typeof results === "function") {
            getResult = results;
        } else if (Array.isArray(results)) {
            if (Array.isArray(results[0])) {
                getResult = (row, col) => results[row - top][col - left];
            } else {
                getResult = (row, col) => results[(row - top) * width + (col - left)];
            }
        } else {
            getResult = () => undefined;
        }
        let first = true;
        for (let r = top; r <= bottom; r++) {
            for (let c = left; c <= right; c++) {
                if (first) {
                    this.getCell(r, c).value = {
                        shareType,
                        formula,
                        ref: range,
                        result: getResult(r, c),
                    };
                    first = false;
                } else {
                    this.getCell(r, c).value = isShared
                        ? {
                              sharedFormula: masterAddress,
                              result: getResult(r, c),
                          }
                        : getResult(r, c);
                }
            }
        }
    }

    // =========================================================================
    // Images
    addImage(imageId, range) {
        const model = {
            type: "image",
            imageId,
            range,
        };
        this._media.push(new Image(this, model));
    }

    getImages() {
        return this._media.filter((m) => m.type === "image");
    }

    addBackgroundImage(imageId) {
        const model = {
            type: "background",
            imageId,
        };
        this._media.push(new Image(this, model));
    }

    getBackgroundImageId() {
        const image = this._media.find((m) => m.type === "background");
        return image && image.imageId;
    }

    // =========================================================================
    // Worksheet Protection
    protect(password, options) {
        // TODO: make this function truly async
        // perhaps marshal to worker thread or something
        return new Promise((resolve) => {
            this.sheetProtection = {
                sheet: true,
            };
            if (options && "spinCount" in options) {
                // force spinCount to be integer >= 0
                options.spinCount = Number.isFinite(options.spinCount) ? Math.round(Math.max(0, options.spinCount)) : 100000;
            }
            if (password) {
                this.sheetProtection.algorithmName = "SHA-512";
                this.sheetProtection.saltValue = Encryptor.randomBytes(16).toString("base64");
                this.sheetProtection.spinCount = options && "spinCount" in options ? options.spinCount : 100000; // allow user specified spinCount
                this.sheetProtection.hashValue = Encryptor.convertPasswordToHash(
                    password,
                    "SHA512",
                    this.sheetProtection.saltValue,
                    this.sheetProtection.spinCount
                );
            }
            if (options) {
                this.sheetProtection = Object.assign(this.sheetProtection, options);
                if (!password && "spinCount" in options) {
                    delete this.sheetProtection.spinCount;
                }
            }
            resolve();
        });
    }

    unprotect() {
        this.sheetProtection = null;
    }

    // =========================================================================
    // Tables
    addTable(model) {
        const table = new Table(this, model);
        setTable(this.tables, model.name, table);
        return table;
    }

    getTable(name) {
        return Object.prototype.hasOwnProperty.call(this.tables, name) ? this.tables[name] : undefined;
    }

    removeTable(name) {
        delete this.tables[name];
    }

    getTables() {
        return Object.values(this.tables);
    }

    // =========================================================================
    // Pivot Tables
    addPivotTable(model) {
        console.warn(
            `Warning: Pivot Table support is experimental. 
Please leave feedback at https://github.com/exceljs/exceljs/discussions/2575`
        );

        const pivotTable = makePivotTable(this, model);

        this.pivotTables.push(pivotTable);
        this.workbook.pivotTables.push(pivotTable);

        return pivotTable;
    }

    // ===========================================================================
    // Conditional Formatting
    addConditionalFormatting(cf) {
        this.conditionalFormattings.push(cf);
    }

    removeConditionalFormatting(filter) {
        if (typeof filter === "number") {
            this.conditionalFormattings.splice(filter, 1);
        } else if (filter instanceof Function) {
            this.conditionalFormattings = this.conditionalFormattings.filter(filter);
        } else {
            this.conditionalFormattings = [];
        }
    }

    // ===========================================================================
    // Deprecated
    get tabColor() {
        // eslint-disable-next-line no-console
        console.trace("worksheet.tabColor property is now deprecated. Please use worksheet.properties.tabColor");
        return this.properties.tabColor;
    }

    set tabColor(value) {
        // eslint-disable-next-line no-console
        console.trace("worksheet.tabColor property is now deprecated. Please use worksheet.properties.tabColor");
        this.properties.tabColor = value;
    }

    // ===========================================================================
    // Model

    get model() {
        const model = {
            id: this.id,
            name: this.name,
            dataValidations: this.dataValidations.model,
            properties: this.properties,
            state: this.state,
            pageSetup: this.pageSetup,
            headerFooter: this.headerFooter,
            rowBreaks: this.rowBreaks,
            views: this.views,
            autoFilter: this.autoFilter,
            media: this._media.map((medium) => medium.model),
            sheetProtection: this.sheetProtection,
            tables: Object.values(this.tables).map((table) => table.model),
            pivotTables: this.pivotTables,
            conditionalFormattings: this.conditionalFormattings,
        };

        // =================================================
        // columns
        model.cols = this._colsModel();

        // ==========================================================
        // Rows
        const rows = (model.rows = []);
        const dimensions = (model.dimensions = new Range());
        eachEntry(this._rows, (row) => {
            const rowModel = row.model;
            if (rowModel) {
                dimensions.expand(rowModel.number, rowModel.min, rowModel.number, rowModel.max);
                rows.push(rowModel);
            }
        });

        // ==========================================================
        // Merges
        model.merges = [];
        _.each(this._merges, (merge) => {
            model.merges.push(merge.range);
        });

        return model;
    }

    _parseRows(model) {
        this._rows = [];
        model.rows.forEach((rowModel) => {
            const row = new Row(this, rowModel.number);
            this._rows[row.number - 1] = row;
            row.model = rowModel;
        });
    }

    _parseMergeCells(model) {
        _.each(model.mergeCells, (merge) => {
            // Do not merge styles when importing an Excel file
            // since each cell may have different styles intentionally.
            this.mergeCellsWithoutStyle(merge);
        });
    }

    set model(value) {
        this.name = value.name;
        if ((value.cols || []).some((col) => col.key || col.header)) {
            // keys and headers take effect when their columns are created
            this._columns = Column.fromModel(this, value.cols);
            this._colRanges = null;
            this._columnsSparse = false;
        } else {
            // the columns of the ranges are created on demand
            this._columns = null;
            this._colRanges = Column.rangesFromModel(value.cols);
            this._columnsSparse = this._colRanges.length > 0;
        }
        this._parseRows(value);

        this._parseMergeCells(value);
        this.dataValidations = new DataValidations(value.dataValidations);
        this.properties = value.properties;
        this.pageSetup = value.pageSetup;
        this.headerFooter = value.headerFooter;
        this.views = value.views;
        this.autoFilter = value.autoFilter;
        this._media = value.media.map((medium) => new Image(this, medium));
        this.sheetProtection = value.sheetProtection;
        this.tables = value.tables.reduce((tables, table) => {
            const t = new Table();
            t.model = table;
            setTable(tables, table.name, t);
            return tables;
        }, {});
        this.pivotTables = value.pivotTables;
        this.conditionalFormattings = value.conditionalFormattings;
    }
}

module.exports = Worksheet;
