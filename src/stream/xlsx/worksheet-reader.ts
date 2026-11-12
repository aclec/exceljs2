import { EventEmitter } from "events";
import parseSax, { type SaxInput } from "../../utils/parse-sax";

import { createDictionary, excelToDate, isDateFmt, parseBoolean, xmlDecode } from "../../utils/utils";
import colCache from "../../utils/col-cache";
import { slideFormula } from "../../utils/shared-formula";
import { parseIsoDate as parseIsoDateOrUndefined } from "../../utils/iso-date";
import Range from "../../doc/range";

import Row, { type RowWorksheet } from "../../doc/row";
import Column, { type ColumnRanges, type ColumnWorksheet } from "../../doc/column";
import type { CellValue } from "../../doc/cell";
import type DataValidations from "../../doc/data-validations";
import ColXform from "../../xlsx/xform/sheet/col-xform";
import RowXform from "../../xlsx/xform/sheet/row-xform";
import SheetFormatPropertiesXform from "../../xlsx/xform/sheet/sheet-format-properties-xform";
import type { CellFormulaValue, WorksheetProperties, WorksheetState } from "../../../index";
import type { ColModel, RowBreakModel } from "../../types";
import type WorkbookReader from "./workbook-reader";
import type { WorkbookReaderOptions } from "./workbook-reader";

const { parseRowNumber } = RowXform;

/**
 * The column number of a cell reference (the `r` of a `<c>`), as `colCache.decodeAddress(r).col`. A canonical
 * reference (`AB12`: letters then digits, in the grid) is scanned without building its decoded address (this runs for
 * every cell); anything else goes through decodeAddress, which validates it (bounds, `$` signs...).
 */
export function columnOfReference(r: string): number {
    let col = 0;
    let row = 0;
    let i = 0;
    let char = r.charCodeAt(0);
    // letters: A (65) to Z (90)
    while (char >= 65 && char <= 90) {
        col = col * 26 + char - 64;
        char = r.charCodeAt(++i);
    }
    const letters = i;
    // digits: 0 (48) to 9 (57)
    while (char >= 48 && char <= 57) {
        row = row * 10 + char - 48;
        char = r.charCodeAt(++i);
    }
    if (letters > 0 && i > letters && i === r.length && col <= colCache.MAX_COLS && row <= colCache.MAX_ROWS) {
        return col;
    }
    return colCache.decodeAddress(r).col;
}

// value of a t="d" cell: a Date, or the text when it is not a valid ISO 8601 date (see iso-date)
const parseIsoDate = (text: string): Date | string => parseIsoDateOrUndefined(text) || text;

/**
 * A `<hyperlink>` of the worksheet, as read (`hyperlinks: "cache"` or `"emit"`): the attributes of CT_Hyperlink present,
 * and with `hyperlinks: "cache"` the target of its relationship.
 */
export interface WorksheetHyperlink {
    /** the cell reference (`A1`) */
    ref: string | undefined;
    /** the relationship of an external link (`r:id`) */
    rId?: string;
    /** an internal link: a location in the workbook */
    location?: string;
    tooltip?: string;
    display?: string;
    /**
     * `hyperlinks: "cache"`: the link (`Target` of the relationship `rId` of the worksheet's relationships part), set
     * once that part is read: when the worksheet is read if the part comes before it in the file, else when the reader
     * reaches it (at the latest when the workbook has been read). Absent when the relationship is missing.
     */
    target?: string;
    /** `hyperlinks: "cache"`: the `TargetMode` of the relationship (`"External"` for a link outside the package). */
    targetMode?: string;
}

/** The arguments of the {@link WorksheetReader} constructor. */
export interface WorksheetReaderOptions {
    workbook: WorkbookReader;
    /** The number of the worksheet part (`sheet3.xml` → "3"), replaced by the sheetId of the workbook when known. */
    id: string | number;
    /** The XML of the worksheet part. */
    iterator: SaxInput;
    /** The options of the workbook reader (`worksheets`, `hyperlinks`). */
    options?: Partial<Pick<WorkbookReaderOptions, "worksheets" | "hyperlinks">>;
}

/** An event of the worksheet: a row, or a hyperlink (`hyperlinks: "emit"`). */
export type WorksheetReaderEvent =
    { readonly eventType: "row"; readonly value: Row } | { readonly eventType: "hyperlink"; readonly value: WorksheetHyperlink };

/** The events of a {@link WorksheetReader}. */
export interface WorksheetReaderEvents {
    row: [row: Row];
    hyperlink: [hyperlink: WorksheetHyperlink];
    finished: [];
    error: [error: unknown];
}

/** The sheet format properties read from `<sheetFormatPr>` (SheetFormatPropertiesXform). */
type SheetFormatProperties = Partial<WorksheetProperties> & { outlineLevelCol: number; outlineLevelRow: number };

// the <f> of a <c> being parsed
interface FormulaBeingParsed {
    text: string;
    /** `t="shared"`: the `si` of the shared formula (`"undefined"` when absent, as the document reader keys it) */
    si?: string;
    /** the `ref` of a shared formula master */
    ref?: string;
}

// a <c> being parsed
interface CellBeingParsed {
    ref: string;
    col: number;
    /** the style id (`NaN` when absent) */
    s: number;
    t: string | undefined;
    f?: FormulaBeingParsed;
    v?: { text: string };
}

/** The master of a shared formula: its formula and the address of its cell. */
interface SharedFormulaMaster {
    formula: string;
    address: string;
}

/**
 * The shared formula masters kept per worksheet (by `si`), for their clones: past this number, the oldest master is
 * dropped (its later clones get no formula). Clones follow their master closely, so this only bounds the memory of a
 * hostile file.
 */
const MAX_SHARED_FORMULA_MASTERS = 65536;

/** The error raised by the members of the worksheet that a read row or column cannot use. */
function readOnly(member: string): Error {
    return new Error(`Invalid Operation: ${member} (the rows and columns of a streaming reader are read-only)`);
}

// the attributes of CT_Hyperlink that a WorksheetHyperlink keeps: [key, attribute name]
const HYPERLINK_ATTRIBUTES = [
    ["rId", "r:id"],
    ["location", "location"],
    ["tooltip", "tooltip"],
    ["display", "display"],
] as const;

/**
 * Reads a worksheet part for the streaming {@link WorkbookReader}: the rows, one at a time (async iteration or `row`
 * events), with their values and, when the styles are cached, their styles; the column definitions; the hyperlinks
 * (`hyperlinks: "emit"` or `"cache"`). The rows are not kept.
 */
class WorksheetReader extends EventEmitter<WorksheetReaderEvents> {
    /** The workbook reader that the worksheet belongs to. */
    readonly workbook: WorkbookReader;
    /** The sheetId from the workbook, or the number of the part (a string) if it could not be matched to the workbook. */
    id: string | number;
    /** The XML of the worksheet part. */
    readonly iterator: SaxInput;
    /** The `worksheets` and `hyperlinks` options of the workbook reader. */
    readonly options: Partial<Pick<WorkbookReaderOptions, "worksheets" | "hyperlinks">>;
    /** The worksheet name (`Sheet<id>` if the worksheet could not be matched to the workbook). */
    name: string;
    /** The visibility of the worksheet (from the workbook), if known. */
    state?: WorksheetState;

    /** `hyperlinks: "cache"`: the hyperlinks read, keyed by cell reference (no prototype: the references come from the file). */
    hyperlinks?: Record<string, WorksheetHyperlink>;

    /** The sheet format properties (outline levels, default sizes), read from `<sheetFormatPr>`. */
    properties: SheetFormatProperties;

    // column definitions: Column objects indexed by number - 1, created on demand (see getColumn), so the array may have
    // holes; and the ranges of the <col> read (see Column.rangesFromModel) whose columns are not created yet. A
    // <col min="1" max="16384"> or a cell in column XFD used to create 16,384 Column objects per worksheet.
    private _columns: Column[] | null;
    private _colRanges: ColumnRanges | null;
    // true when this._columns may have holes or this._colRanges columns not yet created
    private _columnsSparse: boolean;
    // no prototype: keys come from data (e.g. CSV headers) and may be '__proto__' or 'constructor'
    private _keys: Record<string, Column>;
    private _dimensions: Range;
    // the number of the worksheet part (the `id` given), for its relationships part
    private readonly _partNumber: string;

    constructor({ workbook, id, iterator, options }: WorksheetReaderOptions) {
        super();

        this.workbook = workbook;
        this.id = id;
        this._partNumber = String(id);
        this.iterator = iterator;
        this.options = options || {};

        // and a name
        this.name = `Sheet${this.id}`;

        this._columns = null;
        this._colRanges = null;
        this._columnsSparse = false;
        this._keys = createDictionary();

        // keep a record of dimensions
        this._dimensions = new Range();

        // sheet format properties (outline levels, default sizes), read from <sheetFormatPr>
        this.properties = { outlineLevelCol: 0, outlineLevelRow: 0 };
    }

    /**
     * Not a valid operation for a streaming reader.
     *
     * @throws Error always
     */
    destroy(): never {
        throw new Error("Invalid Operation: destroy");
    }

    // =========================================================================
    // The members of a document worksheet that its rows, columns and cells use: a read row or column is read-only, so
    // the Row / Column / Cell methods that reach them (to change or commit a row, to look up another cell, the data
    // validations or the defined names) throw.

    /**
     * Not a valid operation for a read row or column.
     *
     * @throws Error always
     */
    getCell(address: string): never;
    getCell(row: number, col: number): never;
    getCell(): never {
        throw readOnly("getCell");
    }

    /**
     * Not a valid operation for a read row or column.
     *
     * @throws Error always
     */
    findCell(): never {
        throw readOnly("findCell");
    }

    /**
     * Not a valid operation for a read row or column.
     *
     * @throws Error always
     */
    eachRow(): never {
        throw readOnly("eachRow");
    }

    /**
     * Not a valid operation for a read row.
     *
     * @throws Error always
     */
    _commitRow(): never {
        throw readOnly("commit");
    }

    /**
     * Not available for a read row or cell.
     *
     * @throws Error always
     */
    get dataValidations(): DataValidations {
        throw readOnly("dataValidations");
    }

    /**
     * Not available for a read row.
     *
     * @throws Error always
     */
    get rowBreaks(): RowBreakModel[] {
        throw readOnly("rowBreaks");
    }

    /** This reader, as the worksheet of its rows and columns (see above). */
    private get _host(): RowWorksheet & ColumnWorksheet {
        return this;
    }

    /** The range covering the rows read so far. */
    get dimensions(): Range {
        return this._dimensions;
    }

    // =========================================================================
    // Columns

    /** Every column up to the last one read or used (created when it is called), or `null` when there is none. */
    get columns(): Column[] | null {
        if (this._columnsSparse) {
            const length = Math.max(this._columns ? this._columns.length : 0, this._colRanges ? this._colRanges.length : 0);
            if (length) {
                const columns = this._columns || (this._columns = []);
                for (let n = 1; n <= length; n++) {
                    if (!columns[n - 1]) {
                        columns[n - 1] = new Column(this._host, n, Column.defnInRanges(this._colRanges, n));
                    }
                }
            }
            this._colRanges = null;
            this._columnsSparse = false;
        }
        return this._columns;
    }

    /**
     * A column by key, letter or number (1-based), created if needed (the ones before it are created by {@link columns}).
     *
     * @returns `undefined` for a number that is not a positive integer
     * @throws Error for letters that are not a column
     */
    getColumn(c: number | string): Column | undefined {
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
            column = this._columns[c - 1] = new Column(this._host, c, Column.defnInRanges(this._colRanges, c));
        }
        return column;
    }

    /** The column with this key (Column.key), if any. */
    getColumnKey(key: string): Column | undefined {
        return this._keys[key];
    }

    /** Registers the column of a key (Column.key setter). */
    setColumnKey(key: string, value: Column): void {
        this._keys[key] = value;
    }

    /** Unregisters a key (Column.key setter). */
    deleteColumnKey(key: string): void {
        delete this._keys[key];
    }

    /** Calls `f` for each keyed column. */
    eachColumnKey(f: (column: Column, key: string) => void): void {
        for (const [key, column] of Object.entries(this._keys)) {
            f(column, key);
        }
    }

    /**
     * Reads the whole worksheet, emitting its events (`row`, `hyperlink`), then `finished`.
     *
     * @returns resolves once the part is read; errors are emitted (`error`), not thrown
     */
    async read(): Promise<void> {
        try {
            for await (const events of this.parse()) {
                for (const event of events) {
                    if (event.eventType === "row") {
                        this.emit("row", event.value);
                    } else {
                        this.emit("hyperlink", event.value);
                    }
                }
            }
            this.emit("finished");
        } catch (error) {
            this.emit("error", error);
        }
    }

    /** The rows of the worksheet. */
    async *[Symbol.asyncIterator](): AsyncGenerator<Row, void, undefined> {
        for await (const events of this.parse()) {
            for (const event of events) {
                if (event.eventType === "row") {
                    yield event.value;
                }
            }
        }
    }

    /**
     * Parses the worksheet part.
     *
     * @returns the events of the worksheet, in batches (one per chunk of the part that has some)
     * @throws Error on malformed XML or a row number outside the grid
     */
    async *parse(): AsyncGenerator<WorksheetReaderEvent[], void, undefined> {
        const { iterator, options } = this;
        const emitSheet = options.worksheets === "emit";
        const emitHyperlinks = options.hyperlinks === "emit";
        let hyperlinks: Record<string, WorksheetHyperlink> | null = null;
        if (options.hyperlinks === "cache") {
            // keyed by cell references from the file
            this.hyperlinks = hyperlinks = createDictionary();
        }
        if (!emitSheet && !emitHyperlinks && !hyperlinks) {
            return;
        }

        // references
        const { sharedStrings, styles, properties } = this.workbook;
        const date1904 = Boolean(properties && properties.model && properties.model.date1904);

        // xml position
        let inCols = false;
        let inRows = false;
        let inHyperlinks = false;

        // parse state
        const colXform = new ColXform();
        const sheetFormatXform = new SheetFormatPropertiesXform();
        let cols: ColModel[] = [];
        let row: Row | null = null;
        let c: CellBeingParsed | null = null;
        let current: { text: string } | null = null;
        let rowNumber = 0;
        let colNumber = 0;
        // the masters of the shared formulas seen, by si (no prototype: the keys come from the file)
        const sharedFormulas = new Map<string, SharedFormulaMaster>();
        for await (const events of parseSax(iterator)) {
            const worksheetEvents: WorksheetReaderEvent[] = [];
            for (const event of events) {
                if (event.eventType === "opentag") {
                    const node = event.value;
                    const { attributes } = node;
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
                                    // set by parseOpen of a <col>
                                    const col = colXform.model;
                                    if (col) {
                                        colXform.reconcile(col, { styles });
                                        cols.push(col);
                                    }
                                }
                                break;

                            case "row":
                                if (inRows) {
                                    // r is optional: a row without it follows the previous one
                                    const r = parseRowNumber(attributes.r);
                                    rowNumber = r > 0 ? r : rowNumber + 1;
                                    colNumber = 0;
                                    row = new Row(this._host, rowNumber);
                                    if (attributes.ht) {
                                        row.height = parseFloat(attributes.ht);
                                    }
                                    if (parseBoolean(attributes.hidden)) {
                                        row.hidden = true;
                                    }
                                    if (attributes.outlineLevel) {
                                        row.outlineLevel = parseInt(attributes.outlineLevel, 10);
                                    }
                                    if (attributes.s) {
                                        const style = styles.getStyleModel(parseInt(attributes.s, 10));
                                        if (style) {
                                            row.style = style;
                                        }
                                    }
                                }
                                break;
                            case "c":
                                if (row) {
                                    // r is optional: a cell without it follows the previous one
                                    colNumber = attributes.r ? columnOfReference(attributes.r) : colNumber + 1;
                                    c = {
                                        ref: attributes.r || colCache.encodeAddress(row.number, colNumber),
                                        col: colNumber,
                                        s: attributes.s === undefined ? NaN : parseInt(attributes.s, 10),
                                        t: attributes.t,
                                    };
                                }
                                break;
                            case "f":
                                if (c) {
                                    current = c.f = { text: "" };
                                    if (attributes.t === "shared") {
                                        // as the document reader: a missing si is the key "undefined"
                                        c.f.si = String(attributes.si);
                                        c.f.ref = attributes.ref;
                                    }
                                }
                                break;
                            case "v":
                            case "is":
                            case "t":
                                if (c) {
                                    current = c.v = { text: "" };
                                }
                                break;
                            default:
                                break;
                        }
                    }

                    if ((emitHyperlinks || hyperlinks) && node.name === "hyperlinks") {
                        inHyperlinks = true;
                    } else if ((emitHyperlinks || hyperlinks) && node.name === "hyperlink" && inHyperlinks) {
                        // the attributes of CT_Hyperlink: an internal link has a location and no r:id
                        const hyperlink: WorksheetHyperlink = { ref: attributes.ref };
                        for (const [key, name] of HYPERLINK_ATTRIBUTES) {
                            const value = attributes[name];
                            if (value !== undefined) {
                                hyperlink[key] = value;
                            }
                        }
                        if (emitHyperlinks) {
                            worksheetEvents.push({ eventType: "hyperlink", value: hyperlink });
                        } else if (hyperlinks) {
                            // String(): a hyperlink without ref is keyed "undefined", as before
                            hyperlinks[String(hyperlink.ref)] = hyperlink;
                            if (hyperlink.rId !== undefined) {
                                // the target, from the relationships part of the worksheet (now, or once it is read)
                                this.workbook._cacheHyperlinkTarget(this._partNumber, hyperlink);
                            }
                        }
                    }
                } else if (event.eventType === "text") {
                    // only text data is for sheet values
                    if (emitSheet && current) {
                        current.text += event.value;
                    }
                } else if (emitSheet) {
                    switch (event.value.name) {
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
                            // a </row> without its <row> (one outside <sheetData>) has no row to report
                            if (row) {
                                this._dimensions.expandRow(row);
                                worksheetEvents.push({ eventType: "row", value: row });
                                row = null;
                            }
                            break;

                        case "c":
                            if (row && c) {
                                this._reconcileCell(row, c, sharedStrings, styles, date1904, sharedFormulas);
                                c = null;
                            }
                            break;
                        case "hyperlinks":
                            inHyperlinks = false;
                            break;
                        default:
                            break;
                    }
                } else if (event.value.name === "hyperlinks") {
                    inHyperlinks = false;
                }
            }
            if (worksheetEvents.length > 0) {
                yield worksheetEvents;
            }
        }
    }

    // sets the value and style of the cell of a <c> read
    private _reconcileCell(
        row: Row,
        c: CellBeingParsed,
        sharedStrings: WorkbookReader["sharedStrings"],
        styles: WorkbookReader["styles"],
        date1904: boolean,
        sharedFormulas: Map<string, SharedFormulaMaster>
    ): void {
        const cell = row.getCell(c.col);
        if (c.s) {
            const style = styles.getStyleModel(c.s);
            if (style) {
                cell.style = style;
            }
        }

        let value: CellValue;
        if (c.f) {
            const cellValue: CellFormulaValue = {
                formula: sharedFormula(c.ref, c.f, sharedFormulas),
            };
            if (c.v) {
                // as the document reader (CellXform)
                if (c.t === "str") {
                    cellValue.result = xmlDecode(c.v.text);
                } else if (c.t === "b") {
                    cellValue.result = parseInt(c.v.text, 10) !== 0;
                } else if (c.t === "e") {
                    // the error code of the file, kept as is (as CellXform does): Excel writes codes that the
                    // CellErrorValue union of index.d.ts lacks (#SPILL!, #CALC!, #GETTING_DATA...)
                    // eslint-disable-next-line @typescript-eslint/no-unsafe-type-assertion -- see above: the union is to be widened in index.d.ts, for both readers at once
                    cellValue.result = { error: c.v.text } as CellFormulaValue["result"];
                } else if (c.t === "d") {
                    cellValue.result = parseIsoDate(c.v.text);
                } else {
                    cellValue.result = parseFloat(c.v.text);
                }
            }
            value = cellValue;
        } else if (c.v) {
            switch (c.t) {
                case "s": {
                    const index = parseInt(c.v.text, 10);
                    // an index outside the table (malformed file) gives an empty cell
                    value = sharedStrings ? (sharedStrings[index] ?? null) : { sharedString: index };
                    break;
                }

                case "inlineStr":
                case "str":
                    value = xmlDecode(c.v.text);
                    break;

                case "e":
                    value = { error: c.v.text };
                    break;

                case "b":
                    value = parseInt(c.v.text, 10) !== 0;
                    break;

                case "d":
                    value = parseIsoDate(c.v.text);
                    break;

                case undefined:
                default:
                    value = isDateFmt(cell.numFmt) ? excelToDate(parseFloat(c.v.text), date1904) : parseFloat(c.v.text);
                    break;
            }
        } else {
            return;
        }
        cell.value = value;
    }
}

/**
 * The formula of a `<c>`: its own text, or for a clone of a shared formula the formula of its master translated to the
 * cell (as `cell.formula` gives it after `workbook.xlsx.load`). A master (an `<f t="shared">` with a `ref`) is kept by
 * its si for its clones, which follow it in the worksheet; a clone whose master was not seen (or was dropped, see
 * MAX_SHARED_FORMULA_MASTERS) keeps its own text: none.
 * @throws Error when a translated reference is out of the grid, or the formula of the master is longer than the Excel
 * limit (as `cell.formula` of the clone after a load)
 */
function sharedFormula(address: string, f: FormulaBeingParsed, masters: Map<string, SharedFormulaMaster>): string {
    if (f.si === undefined) {
        return f.text;
    }
    if (f.ref) {
        if (!masters.has(f.si) && masters.size >= MAX_SHARED_FORMULA_MASTERS) {
            // the oldest one (Map keeps the insertion order)
            for (const si of masters.keys()) {
                masters.delete(si);
                break;
            }
        }
        masters.set(f.si, { formula: f.text, address });
        return f.text;
    }
    const master = masters.get(f.si);
    return master ? slideFormula(master.formula, master.address, address) : f.text;
}

export default WorksheetReader;
