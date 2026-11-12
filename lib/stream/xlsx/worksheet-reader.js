const { EventEmitter } = require("events");
const parseSax = require("../../utils/parse-sax");

const _ = require("../../utils/under-dash");
const utils = require("../../utils/utils");
const colCache = require("../../utils/col-cache");
const isoDate = require("../../utils/iso-date");
const Dimensions = require("../../doc/range");

const Row = require("../../doc/row");
const Column = require("../../doc/column");
const ColXform = require("../../xlsx/xform/sheet/col-xform");
const RowXform = require("../../xlsx/xform/sheet/row-xform");
const SheetFormatPropertiesXform = require("../../xlsx/xform/sheet/sheet-format-properties-xform");

// value of a t="d" cell: a Date, or the text when it is not a valid ISO 8601 date (see iso-date)
const parseIsoDate = (text) => isoDate.parseIsoDate(text) || text;

class WorksheetReader extends EventEmitter {
    constructor({ workbook, id, iterator, options }) {
        super();

        this.workbook = workbook;
        this.id = id;
        this.iterator = iterator;
        this.options = options || {};

        // and a name
        this.name = `Sheet${this.id}`;

        // column definitions: Column objects indexed by number - 1, created on demand (see getColumn), so the
        // array may have holes; and the ranges of the <col> read ({ranges, length}, see Column.rangesFromModel)
        // whose columns are not created yet. A <col min="1" max="16384"> or a cell in column XFD used to create
        // 16,384 Column objects per worksheet.
        this._columns = null;
        this._colRanges = null;
        // true when this._columns may have holes or this._colRanges columns not yet created
        this._columnsSparse = false;
        // no prototype: keys come from data (e.g. CSV headers) and may be '__proto__' or 'constructor'
        this._keys = Object.create(null);

        // keep a record of dimensions
        this._dimensions = new Dimensions();

        // sheet format properties (outline levels, default sizes), read from <sheetFormatPr>
        this.properties = { outlineLevelCol: 0, outlineLevelRow: 0 };
    }

    // destroy - not a valid operation for a streaming writer
    // even though some streamers might be able to, it's a bad idea.
    destroy() {
        throw new Error("Invalid Operation: destroy");
    }

    // return the current dimensions of the writer
    get dimensions() {
        return this._dimensions;
    }

    // =========================================================================
    // Columns

    // get the current columns array: every column up to the last one read or used, created when it is called
    get columns() {
        if (this._columnsSparse) {
            const length = Math.max(this._columns ? this._columns.length : 0, this._colRanges ? this._colRanges.length : 0);
            if (length) {
                const columns = this._columns || (this._columns = []);
                for (let n = 1; n <= length; n++) {
                    if (!columns[n - 1]) {
                        columns[n - 1] = new Column(this, n, Column.defnInRanges(this._colRanges, n));
                    }
                }
            }
            this._colRanges = null;
            this._columnsSparse = false;
        }
        return this._columns;
    }

    // get a single column by col number. If it doesn't exist, it is created (the ones before it are created by
    // the columns getter).
    getColumn(c) {
        if (typeof c === "string") {
            // if it matches a key'd column, return that
            const col = this._keys[c];
            if (col) {
                return col;
            }

            // otherise, assume letter
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
            if (c - 1 !== this._columns.length) {
                this._columnsSparse = true;
            }
            column = this._columns[c - 1] = new Column(this, c, Column.defnInRanges(this._colRanges, c));
        }
        return column;
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

    async read() {
        try {
            for await (const events of this.parse()) {
                for (const { eventType, value } of events) {
                    this.emit(eventType, value);
                }
            }
            this.emit("finished");
        } catch (error) {
            this.emit("error", error);
        }
    }

    async *[Symbol.asyncIterator]() {
        for await (const events of this.parse()) {
            for (const { eventType, value } of events) {
                if (eventType === "row") {
                    yield value;
                }
            }
        }
    }

    async *parse() {
        const { iterator, options } = this;
        let emitSheet = false;
        let emitHyperlinks = false;
        let hyperlinks = null;
        switch (options.worksheets) {
            case "emit":
                emitSheet = true;
                break;
            case "prep":
                break;
            default:
                break;
        }
        switch (options.hyperlinks) {
            case "emit":
                emitHyperlinks = true;
                break;
            case "cache":
                // keyed by cell references from the file
                this.hyperlinks = hyperlinks = Object.create(null);
                break;
            default:
                break;
        }
        if (!emitSheet && !emitHyperlinks && !hyperlinks) {
            return;
        }

        // references
        const { sharedStrings, styles, properties } = this.workbook;

        // xml position
        let inCols = false;
        let inRows = false;
        let inHyperlinks = false;

        // parse state
        const colXform = new ColXform();
        const sheetFormatXform = new SheetFormatPropertiesXform();
        let cols = null;
        let row = null;
        let c = null;
        let current = null;
        let rowNumber = 0;
        let colNumber = 0;
        for await (const events of parseSax(iterator)) {
            const worksheetEvents = [];
            for (const { eventType, value } of events) {
                if (eventType === "opentag") {
                    const node = value;
                    if (emitSheet) {
                        switch (node.name) {
                            case "sheetFormatPr":
                                sheetFormatXform.parseOpen(node);
                                Object.assign(this.properties, sheetFormatXform.model);
                                break;
                            case "cols":
                                inCols = true;
                                cols = [];
                                break;
                            case "sheetData":
                                inRows = true;
                                break;

                            case "col":
                                if (inCols) {
                                    // same parsing (and grid clamping) as the document reader
                                    colXform.parseOpen(node);
                                    colXform.reconcile(colXform.model, { styles });
                                    cols.push(colXform.model);
                                }
                                break;

                            case "row":
                                if (inRows) {
                                    // r is optional: a row without it follows the previous one
                                    const r = RowXform.parseRowNumber(node.attributes.r);
                                    rowNumber = r > 0 ? r : rowNumber + 1;
                                    colNumber = 0;
                                    row = new Row(this, rowNumber);
                                    if (node.attributes.ht) {
                                        row.height = parseFloat(node.attributes.ht);
                                    }
                                    if (utils.parseBoolean(node.attributes.hidden)) {
                                        row.hidden = true;
                                    }
                                    if (node.attributes.outlineLevel) {
                                        row.outlineLevel = parseInt(node.attributes.outlineLevel, 10);
                                    }
                                    if (node.attributes.s) {
                                        const styleId = parseInt(node.attributes.s, 10);
                                        const style = styles.getStyleModel(styleId);
                                        if (style) {
                                            row.style = style;
                                        }
                                    }
                                }
                                break;
                            case "c":
                                if (row) {
                                    // r is optional: a cell without it follows the previous one
                                    colNumber = node.attributes.r ? colCache.decodeAddress(node.attributes.r).col : colNumber + 1;
                                    c = {
                                        ref: node.attributes.r || colCache.encodeAddress(row.number, colNumber),
                                        col: colNumber,
                                        s: parseInt(node.attributes.s, 10),
                                        t: node.attributes.t,
                                    };
                                }
                                break;
                            case "f":
                                if (c) {
                                    current = c.f = { text: "" };
                                }
                                break;
                            case "v":
                                if (c) {
                                    current = c.v = { text: "" };
                                }
                                break;
                            case "is":
                            case "t":
                                if (c) {
                                    current = c.v = { text: "" };
                                }
                                break;
                            case "mergeCell":
                                break;
                            default:
                                break;
                        }
                    }

                    // =================================================================
                    //
                    if (emitHyperlinks || hyperlinks) {
                        switch (node.name) {
                            case "hyperlinks":
                                inHyperlinks = true;
                                break;
                            case "hyperlink":
                                if (inHyperlinks) {
                                    // the attributes of CT_Hyperlink: an internal link has a location and no r:id
                                    const hyperlink = { ref: node.attributes.ref };
                                    [
                                        ["rId", "r:id"],
                                        ["location", "location"],
                                        ["tooltip", "tooltip"],
                                        ["display", "display"],
                                    ].forEach(([key, name]) => {
                                        if (node.attributes[name] !== undefined) {
                                            hyperlink[key] = node.attributes[name];
                                        }
                                    });
                                    if (emitHyperlinks) {
                                        worksheetEvents.push({ eventType: "hyperlink", value: hyperlink });
                                    } else {
                                        hyperlinks[hyperlink.ref] = hyperlink;
                                    }
                                }
                                break;
                            default:
                                break;
                        }
                    }
                } else if (eventType === "text") {
                    // only text data is for sheet values
                    if (emitSheet) {
                        if (current) {
                            current.text += value;
                        }
                    }
                } else if (eventType === "closetag") {
                    const node = value;
                    if (emitSheet) {
                        switch (node.name) {
                            case "cols":
                                inCols = false;
                                // the columns of the ranges are created when they are used
                                this._columns = null;
                                this._colRanges = Column.rangesFromModel(cols);
                                this._columnsSparse = this._colRanges.length > 0;
                                break;
                            case "sheetData":
                                inRows = false;
                                break;

                            case "row":
                                this._dimensions.expandRow(row);
                                worksheetEvents.push({ eventType: "row", value: row });
                                row = null;
                                break;

                            case "c":
                                if (row && c) {
                                    const cell = row.getCell(c.col);
                                    if (c.s) {
                                        const style = styles.getStyleModel(c.s);
                                        if (style) {
                                            cell.style = style;
                                        }
                                    }

                                    if (c.f) {
                                        const cellValue = {
                                            formula: c.f.text,
                                        };
                                        if (c.v) {
                                            // as the document reader (CellXform)
                                            if (c.t === "str") {
                                                cellValue.result = utils.xmlDecode(c.v.text);
                                            } else if (c.t === "b") {
                                                cellValue.result = parseInt(c.v.text, 10) !== 0;
                                            } else if (c.t === "e") {
                                                cellValue.result = { error: c.v.text };
                                            } else if (c.t === "d") {
                                                cellValue.result = parseIsoDate(c.v.text);
                                            } else {
                                                cellValue.result = parseFloat(c.v.text);
                                            }
                                        }
                                        cell.value = cellValue;
                                    } else if (c.v) {
                                        switch (c.t) {
                                            case "s": {
                                                const index = parseInt(c.v.text, 10);
                                                if (sharedStrings) {
                                                    cell.value = sharedStrings[index];
                                                } else {
                                                    cell.value = {
                                                        sharedString: index,
                                                    };
                                                }
                                                break;
                                            }

                                            case "inlineStr":
                                            case "str":
                                                cell.value = utils.xmlDecode(c.v.text);
                                                break;

                                            case "e":
                                                cell.value = { error: c.v.text };
                                                break;

                                            case "b":
                                                cell.value = parseInt(c.v.text, 10) !== 0;
                                                break;

                                            case "d":
                                                cell.value = parseIsoDate(c.v.text);
                                                break;

                                            default:
                                                if (utils.isDateFmt(cell.numFmt)) {
                                                    cell.value = utils.excelToDate(
                                                        parseFloat(c.v.text),
                                                        properties.model && properties.model.date1904
                                                    );
                                                } else {
                                                    cell.value = parseFloat(c.v.text);
                                                }
                                                break;
                                        }
                                    }
                                    if (hyperlinks) {
                                        const hyperlink = hyperlinks[c.ref];
                                        if (hyperlink) {
                                            cell.text = cell.value;
                                            cell.value = undefined;
                                            cell.hyperlink = hyperlink;
                                        }
                                    }
                                    c = null;
                                }
                                break;
                            default:
                                break;
                        }
                    }
                    if (emitHyperlinks || hyperlinks) {
                        switch (node.name) {
                            case "hyperlinks":
                                inHyperlinks = false;
                                break;
                            default:
                                break;
                        }
                    }
                }
            }
            if (worksheetEvents.length > 0) {
                yield worksheetEvents;
            }
        }
    }
}

module.exports = WorksheetReader;
