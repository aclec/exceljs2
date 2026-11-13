// A worksheet of a Workbook document: rows and cells, columns, merges, images, tables, data validations, conditional
// formatting, printing and protection settings.
//
// Worksheet requirements
//  Operate as sheet inside workbook or standalone
//  Load and Save from file and stream
//  Access/Add/Delete individual cells
//  Manage column widths and row heights

import type { WorksheetProtection, WorksheetState } from "../../index";
import colCache from "../utils/col-cache";
import Range from "./range";
import { findCellAddress, isColumnNumber, isRowNumber, resolveCellAddress, resolveRange, type SheetRangeArgs } from "./sheet-reference";
import Row, { type RowValues } from "./row";
import Column, { type ColumnDefn, type ColumnRanges, type ColumnSegment } from "./column";
import { ValueType } from "./enums";
import Image, { type ImageRangeInput } from "./image";
import Table from "./table";
import DataValidations from "./data-validations";
import { makePivotTable, type PivotTable, type PivotTableInput } from "./pivot-table";
import Encryptor from "../utils/encryptor";
import { copyStyle } from "../utils/copy-style";
import { checkMergedCells } from "../utils/merge-limits";
import { defaultHeaderFooter } from "../xlsx/xform/sheet/header-footer-xform";
import { eachEntry, entryIndices, lastEntryIndex, shiftSteps } from "../utils/sparse-array";
import type Cell from "./cell";
import { isFormulaResult, type CellValue, type FormulaValueObject } from "./cell";
import type Workbook from "./workbook";
import type {
    AutoFilterModel,
    ColModel,
    ConditionalFormattingModel,
    FormulaResult,
    HeaderFooterModel,
    PageSetupModel,
    RowBreakModel,
    RowModel,
    SheetProtectionModel,
    SheetViewModel,
    TableModel,
    WorksheetModel,
    WorksheetPropertiesModel,
} from "../types";

/** The options of a new worksheet (`workbook.addWorksheet` adds `id`, `name`, `orderNo` and `workbook`). */
export interface WorksheetOptions {
    workbook: Workbook;
    /** The id of the worksheet in the workbook (`sheetId`). */
    id: number;
    /** The position of the worksheet among the others (undefined or -1 for a worksheet loaded without `<sheet>`). */
    orderNo?: number | undefined;
    /** The name (default `sheet<id>`). */
    name?: string | undefined;
    state?: WorksheetState | undefined;
    properties?: WorksheetPropertiesModel | undefined;
    pageSetup?: PageSetupModel | undefined;
    headerFooter?: HeaderFooterModel | undefined;
    views?: SheetViewModel[] | undefined;
    autoFilter?: AutoFilterModel | null | undefined;
}

/**
 * {@link Worksheet.model} as the getter produces it: a {@link WorksheetModel} with the fields only the getter sets
 * (`state`, `rowBreaks`, the used range `dimensions` and the merged ranges `merges`).
 */
export type DocumentWorksheetModel = Omit<WorksheetModel, "state" | "rowBreaks" | "dimensions" | "merges"> & {
    state: WorksheetState;
    rowBreaks: RowBreakModel[];
    dimensions: Range;
    merges: string[];
};

/** Options of {@link Worksheet.eachRow}. */
export interface EachRowOptions {
    /** Also visit the empty rows, up to the last one (creating the missing ones). */
    readonly includeEmpty?: boolean;
}

/** An iteratee of {@link Worksheet.eachRow}: the row and its 1-based number. */
export type RowIteratee = (row: Row, rowNumber: number) => void;

/** The results of {@link Worksheet.fillFormula}: by cell, as a 2D array (rows of the range), or as a flat array. */
export type FillFormulaResults = ((row: number, col: number) => FormulaResult | undefined) | FormulaResultRows | FormulaResultCells;

/** The results of {@link Worksheet.fillFormula} by row of the range, then by column. */
type FormulaResultRows = readonly (readonly (FormulaResult | undefined)[])[];
/** The results of {@link Worksheet.fillFormula} in a flat array, row by row. */
type FormulaResultCells = readonly (FormulaResult | undefined)[];

/** Whether fillFormula results are given by row: a formula result is never an array. */
function isResultRows(results: FormulaResultRows | FormulaResultCells): results is FormulaResultRows {
    return Array.isArray(results[0]);
}

/**
 * The style option of `addRow` / `insertRow`: `"i"` copies the style of the row above (`"i+"`: also its empty cells),
 * `"o"` the style of the row below (`"o+"`), `"n"` (the default) none.
 */
export type RowStyleOption = string;

// worksheet.tables is a plain object (by table name). Names come from loaded files, so it is only
// read and written through own properties: "__proto__" or "constructor" are valid table names.
function setTable(tables: Record<string, Table>, name: string, table: Table): void {
    Object.defineProperty(tables, name, { value: table, enumerable: true, writable: true, configurable: true });
}

// The span [from, to] of a merge's rows (or columns) after a splice removed count of them at start and inserted nInserts
// there: the moved span of its remaining rows, which also covers rows inserted between two remaining ones (as Excel widens
// a merge when rows are inserted inside it). undefined when none of its rows remains.
function spliceSpan(from: number, to: number, start: number, count: number, nInserts: number): [number, number] | undefined {
    const end = start + count; // first row after the removed ones
    const first = from >= start && from < end ? end : from;
    const last = to >= start && to < end ? start - 1 : to;
    if (first > last) {
        return undefined;
    }
    const move = (n: number): number => (n < start ? n : n + nInserts - count);
    return [move(first), move(last)];
}

// Excel widens a merge over the rows or columns inserted inside it (spliceSpan), and refuses to change part of a merged cell:
// only the master of a merge holds a value, so a value inserted in another cell of the widened merge would be lost. Splices
// refuse such values, before anything changes.
function checkInsertedValue(value: CellValue | undefined, row: number, col: number, range: Range): void {
    if (value !== undefined && value !== null && value !== "") {
        throw new Error(
            `Cannot insert a value in ${colCache.encodeAddress(row, col)}: it would be inside the merged cells ${range.range}, unmerge them first`
        );
    }
}

/** Whether row values are an array of values (else values by column key). */
function isValueArray(values: RowValues): values is readonly (CellValue | undefined)[] {
    return Array.isArray(values);
}

/** Whether the columns array has no hole (once the columns getter filled them). */
function isDense(columns: (Column | undefined)[]): columns is Column[] {
    for (let i = 0; i < columns.length; i++) {
        if (columns[i] === undefined) {
            return false;
        }
    }
    return true;
}

// the values that row.values = values puts in the columns (see Row#values): a function of the column number
function insertedRowValues(worksheet: Worksheet, values: RowValues | null | undefined): (col: number) => CellValue | undefined {
    if (!values) {
        return () => undefined;
    }
    if (isValueArray(values)) {
        const offset = Object.hasOwn(values, "0") ? 1 : 0;
        return (col) => values[col - offset];
    }
    const byKey = values;
    // the column keys of the worksheet, as Row#values resolves them
    const byColumn = new Map<number, CellValue | undefined>();
    worksheet.eachColumnKey((column, key) => {
        if ((!(key in Object.prototype) || Object.hasOwn(byKey, key)) && byKey[key] !== undefined) {
            byColumn.set(column.number, byKey[key]);
        }
    });
    return (col) => byColumn.get(col);
}

/**
 * A worksheet of a workbook. Rows, cells and columns are created on demand; merged cells, images, tables, data
 * validations, conditional formattings, page setup and protection are kept here and written by the xlsx writer from
 * {@link Worksheet.model}.
 */
class Worksheet {
    /** The workbook (the one the worksheet was created in, even once removed from it). */
    _workbook: Workbook;
    /** The id of the worksheet in the workbook (`sheetId`, the index in `workbook._worksheets`). */
    id: number;
    /** The position of the worksheet among the others (see {@link WorksheetOptions}). */
    orderNo: number | undefined;
    declare _name: string;
    /** The visibility of the worksheet. */
    state: WorksheetState;
    /** The rows, indexed by row number - 1 (sparse). */
    _rows: (Row | undefined)[];
    /**
     * Column definitions: Column objects indexed by number - 1, created on demand (see getColumn), so the array may have
     * holes. A single cell in column XFD or a `<col min="1" max="16384">` on each worksheet used to create 16384 Column
     * objects per worksheet.
     */
    _columns: (Column | undefined)[] | null;
    /** The ranges of loaded cols ({ranges, length}, see Column.rangesFromModel) whose columns are not created yet. */
    _colRanges: ColumnRanges | null;
    /** True when `_columns` may have holes or `_colRanges` columns not yet created. */
    _columnsSparse: boolean;
    /** Column keys (addRow convenience): key → column (a Map: keys come from data, e.g. CSV headers). */
    _keys: Map<string, Column>;
    /** The merges, by the address of their master (top-left) cell. */
    _merges: Record<string, Range>;
    /** The number of merged cells (the area of the merges). */
    _mergedCellCount: number;
    /** The manual page breaks after rows. */
    rowBreaks: RowBreakModel[];
    /** Tab color, default row height, outline levels, etc. */
    properties: WorksheetPropertiesModel;
    /** For all things printing. */
    pageSetup: PageSetupModel;
    /** Headers and footers of the printed pages (`null` texts by default, also when read from a file without them). */
    headerFooter: HeaderFooterModel;
    /** The data validations of the cells. */
    dataValidations: DataValidations;
    /** For freezepanes, split, zoom, gridlines, etc. (`[]` by default, also when read from a file without them). */
    views: SheetViewModel[];
    /** The auto filter of the worksheet (a range such as `"A1:C1"`, or its corners), if any. */
    autoFilter: AutoFilterModel | null | undefined;
    /** The images (and background image). */
    _media: Image[];
    /** The protection of the worksheet, or null when it is not protected. */
    sheetProtection: SheetProtectionModel | null | undefined;
    /** The tables by name (see setTable). */
    tables: Record<string, Table>;
    /** The pivot tables of the worksheet (`addPivotTable`). */
    pivotTables: PivotTable[];
    /** The conditional formattings (`addConditionalFormatting`). */
    conditionalFormattings: ConditionalFormattingModel[];
    /** The number of header rows of the columns set by the `columns` setter. */
    declare _headerRowCount?: number;

    /**
     * @throws Error for an invalid name (see the name setter)
     */
    constructor(options: WorksheetOptions) {
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

        // column keys (addRow convenience): key ==> column; a Map: keys come from data (e.g. CSV headers) and may be
        // '__proto__' or 'constructor'
        this._keys = new Map();

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
            } satisfies PageSetupModel,
            options.pageSetup
        );

        this.headerFooter = Object.assign(defaultHeaderFooter(), options.headerFooter);

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

    /** The name of the worksheet. */
    get name(): string {
        return this._name;
    }

    /**
     * Sets the name: nothing gives `sheet<id>`; a name longer than 31 characters is truncated (with a warning).
     * @throws Error for a name that is not a string, empty, "History", with a character among `* ? : \ / [ ]`, starting
     *   or ending with a quote, or of another worksheet of the workbook (case-insensitive)
     */
    set name(name: string | undefined) {
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

    /** The workbook. */
    get workbook(): Workbook {
        return this._workbook;
    }

    /** Removes the worksheet from its workbook (when you're done with it). */
    destroy(): void {
        this._workbook.removeWorksheetEx(this);
    }

    /** The bounding range of the non-empty cells (a new Range; `A1:A1` for an empty worksheet). */
    get dimensions(): Range {
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

    /** The columns up to the last one defined or created (all created by this getter); null when there are none. */
    get columns(): Column[] | null {
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
        const columns = this._columns;
        if (!columns) {
            return null;
        }
        // not sparse: no holes
        if (!isDense(columns)) {
            throw new Error("The columns of the worksheet have holes");
        }
        return columns;
    }

    /** The number of columns, created or defined by the loaded ranges. */
    _columnsLength(): number {
        return Math.max(this._columns ? this._columns.length : 0, this._colRanges ? this._colRanges.length : 0);
    }

    /** The definition of column n in the loaded ranges, if any. */
    _columnDefn(n: number): ColumnDefn | undefined {
        return Column.defnInRanges(this._colRanges, n);
    }

    /** The cols of the model, in time proportional to the created columns and loaded ranges. */
    _colsModel(): ColModel[] | undefined {
        if (!this._columnsSparse) {
            return Column.toModel(this._columns);
        }
        // the created columns, in order, with their numbers
        const created: [number, Column][] = [];
        eachEntry(this._columns || [], (column, index) => {
            created.push([index + 1, column]);
        });
        const ranges = this._colRanges ? this._colRanges.ranges : [];
        const length = this._columnsLength();
        const segments: ColumnSegment[] = [];
        let plainColumn: Column | undefined;
        let i = 0;
        let r = 0;
        let n = 1;
        while (n <= length) {
            const next = created[i];
            if (next && next[0] === n) {
                segments.push({ min: n, max: n, column: next[1] });
                i++;
                n++;
            } else {
                // columns n.. up to the next created column stand for the same range, or are plain columns
                const nextCreated = next ? next[0] : length + 1;
                let range = ranges[r];
                while (range && range.max < n) {
                    range = ranges[++r];
                }
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

    /**
     * Sets the columns from an array of column definitions (the existing columns are replaced).
     * Note: any headers defined will overwrite existing values.
     */
    set columns(value: readonly (ColumnDefn & { readonly headers?: readonly string[] })[]) {
        // calculate max header row count
        this._headerRowCount = value.reduce((pv, cv) => {
            const headerCount = (cv.header && 1) || (cv.headers && cv.headers.length) || 0;
            return Math.max(pv, headerCount);
        }, 0);

        // construct Column objects
        let count = 1;
        this._colRanges = null;
        this._columnsSparse = false;
        const columns: Column[] = (this._columns = []);
        value.forEach((defn) => {
            const column = new Column(this, count++, false);
            columns.push(column);
            column.defn = defn;
        });
    }

    /** The column of a key, if any. */
    getColumnKey(key: string): Column | undefined {
        return this._keys.get(key);
    }

    /** Registers `value` as the column of `key` (`column.key = …` calls it). */
    setColumnKey(key: string, value: Column): void {
        this._keys.set(key, value);
    }

    /** Unregisters the column of `key`. */
    deleteColumnKey(key: string): void {
        this._keys.delete(key);
    }

    /** Calls `f` with each keyed column and its key. */
    eachColumnKey(f: (column: Column, key: string) => void): void {
        for (const [key, column] of this._keys) {
            f(column, key);
        }
    }

    /**
     * A column by key, letters (`"B"`) or 1-based number, created if needed (the columns before it are created when they
     * are used, or by the `columns` getter).
     * @returns undefined for a number that is not an integer >= 1
     */
    getColumn(c: number | string): Column | undefined {
        if (typeof c === "string") {
            // if it matches a key'd column, return that
            const col = this._keys.get(c);
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

    /**
     * Column c if it exists, without creating it (and the ones before it) like getColumn: a column of a loaded range
     * that was not created yet is returned as a detached Column.
     */
    _findColumn(c: number): Column | undefined {
        const column = this._columns && this._columns[c - 1];
        if (column) {
            return column;
        }
        const defn = Number.isInteger(c) && c >= 1 ? this._columnDefn(c) : undefined;
        return defn ? new Column(this, c, defn) : undefined;
    }

    /**
     * Column n, created if needed (see {@link getColumn}).
     * @throws Error for a number that is not an integer >= 1
     */
    _getColumnEx(n: number): Column {
        const column = this.getColumn(n);
        if (!column) {
            throw new Error(`Invalid column number: ${n} (the columns are numbered from 1)`);
        }
        return column;
    }

    /**
     * Removes `count` columns from column `start` (1-based) and inserts columns with the `inserts` values (each an array
     * of the values of the column, indexed by row - 1). The cells, column definitions, merges and defined names move.
     * @throws Error for a start that is not a column number or a count that is not an integer >= 0 (nothing changes),
     *   when a value would be inserted inside a merge, or the merges would exceed the cap
     */
    spliceColumns(start: number, count: number, ...inserts: readonly (CellValue | undefined)[][]): void {
        if (!isColumnNumber(start)) {
            throw new Error(`Invalid column number: ${String(start)} (the columns are numbered from 1 to ${colCache.MAX_COLS})`);
        }
        if (!Number.isInteger(count) || count < 0) {
            throw new Error(`Invalid count of columns: ${count}`);
        }
        const merges = this._planMerges(
            (merge) => merge.right >= start,
            (merge) => {
                const span = spliceSpan(merge.left, merge.right, start, count, inserts.length);
                return span && [merge.top, span[0], merge.bottom, span[1]];
            }
        );
        for (const range of merges.moved) {
            for (let col = Math.max(start, range.left + 1); col <= Math.min(start + inserts.length - 1, range.right); col++) {
                const values = inserts[col - start];
                for (let row = range.top; row <= range.bottom; row++) {
                    checkInsertedValue(values && values[row - 1], row, col, range);
                }
            }
        }
        this._detachMerges(merges.touched);
        const rows = this._rows;
        if (inserts.length > 0) {
            // the rows that exist and the ones an insert gives a value (also past the last row; 0, false and "" are
            // values): a missing row without value would only get empty cells
            const spliced: boolean[] = [];
            eachEntry(rows, (row, i) => {
                spliced[i] = true;
            });
            inserts.forEach((insert) => {
                eachEntry(insert, (value, i) => {
                    if (value !== null && value !== undefined) {
                        spliced[i] = true;
                    }
                });
            });
            eachEntry(spliced, (value, i) => {
                const rowInserts = inserts.map((insert) => insert[i] ?? null);
                const row = this.getRow(i + 1);

                row.splice(start, count, ...rowInserts);
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
        // no columns when none was ever created (no cell)
        const { columns } = this;
        const nEnd = columns ? columns.length : 0;
        // the column numbers below are integers >= 1 (start is a column number, count an integer >= 0)
        if (nExpand < 0) {
            for (let i = start + inserts.length; i <= nEnd; i++) {
                this._getColumnEx(i).defn = this._getColumnEx(i - nExpand).defn;
            }
        } else if (nExpand > 0) {
            for (let i = nEnd; i >= nKeep; i--) {
                this._getColumnEx(i + nExpand).defn = this._getColumnEx(i).defn;
            }
        }
        for (let i = start; i < start + inserts.length; i++) {
            this._getColumnEx(i).defn = null;
        }

        this._attachMerges(merges.moved);

        // account for defined names
        this.workbook.definedNames.spliceColumns(this.name, start, count, inserts.length);
    }

    /** The last column (`getColumn(columnCount)`; undefined for an empty worksheet). */
    get lastColumn(): Column | undefined {
        return this.getColumn(this.columnCount);
    }

    /** The number of columns: the largest cell count of the rows with values. */
    get columnCount(): number {
        let maxCount = 0;
        this.eachRow((row) => {
            maxCount = Math.max(maxCount, row.cellCount);
        });
        return maxCount;
    }

    /** The number of columns that have a non-empty cell. */
    get actualColumnCount(): number {
        // performance nightmare - for each row, counts all the columns used
        const counts: boolean[] = [];
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

    /** No-op (allows the streaming reader to fill a document). */
    _commitRow(): void {
        // nop - allows streaming reader to fill a document
    }

    get _lastRowNumber(): number {
        // need to cope with results of splice
        return lastEntryIndex(this._rows) + 1;
    }

    get _nextRow(): number {
        return this._lastRowNumber + 1;
    }

    /** The last row (the row of the highest index, even empty), if any. */
    get lastRow(): Row | undefined {
        if (this._rows.length) {
            return this._rows[this._rows.length - 1];
        }
        return undefined;
    }

    /** A row (1-based number) if it exists (`undefined` for a number that is not a row: 0, -1, 1.5...). */
    findRow(r: number): Row | undefined {
        return this._rows[r - 1];
    }

    /** The rows `start`..`start + length - 1` (1-based; missing ones are holes / undefined). */
    findRows(start: number, length: number): (Row | undefined)[] {
        return this._rows.slice(start - 1, start - 1 + length);
    }

    /** The number of the last row (with or without values). */
    get rowCount(): number {
        return this._lastRowNumber;
    }

    /** The number of rows that have values. */
    get actualRowCount(): number {
        // counts actual rows that have actual data
        let count = 0;
        this.eachRow(() => {
            count++;
        });
        return count;
    }

    /**
     * A row (1-based number), created if needed.
     * @throws Error for a number that is not a row (an integer from 1 to 1,048,576)
     */
    getRow(r: number): Row {
        if (!isRowNumber(r)) {
            throw new Error(`Invalid row number: ${String(r)} (the rows are numbered from 1 to ${colCache.MAX_ROWS})`);
        }
        let row = this._rows[r - 1];
        if (!row) {
            row = this._rows[r - 1] = new Row(this, r);
        }
        return row;
    }

    /**
     * The rows `start`..`start + length - 1`, created if needed; undefined when `length` < 1.
     * @throws Error for a number that is not a row (see {@link Worksheet.getRow})
     */
    getRows(start: number, length: number): Row[] | undefined {
        if (length < 1) return undefined;
        const rows = [];
        for (let i = start; i < start + length; i++) {
            rows.push(this.getRow(i));
        }
        return rows;
    }

    /**
     * Adds a row after the last one, with values (see {@link RowValues}).
     * @param style - `"i"` / `"i+"` copies the style of the row above (other values: none)
     */
    addRow(value: RowValues | null | undefined, style: RowStyleOption = "n"): Row {
        const rowNo = this._nextRow;
        const row = this.getRow(rowNo);
        row.values = value;
        this._setStyleOption(rowNo, style[0] === "i" ? style : "n");
        return row;
    }

    /** Adds rows after the last one (see {@link Worksheet.addRow}). */
    addRows(value: readonly (RowValues | null | undefined)[], style: RowStyleOption = "n"): Row[] {
        const rows: Row[] = [];
        value.forEach((row) => {
            rows.push(this.addRow(row, style));
        });
        return rows;
    }

    /**
     * Inserts a row at `pos` (1-based), shifting the rows down (see {@link RowStyleOption}).
     * @throws Error when a value would be inserted inside a merge, or the merges would exceed the cap
     */
    insertRow(pos: number, value: RowValues | null | undefined, style: RowStyleOption = "n"): Row {
        this.spliceRows(pos, 0, value);
        this._setStyleOption(pos, style);
        return this.getRow(pos);
    }

    /**
     * Inserts rows at `pos` (1-based), shifting the rows down (see {@link RowStyleOption}).
     * @throws Error when a value would be inserted inside a merge, or the merges would exceed the cap
     */
    insertRows(pos: number, values: readonly (RowValues | null | undefined)[], style: RowStyleOption = "n"): Row[] | undefined {
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

    /** Sets the row at `pos` to the style of the previous row (option "i") or of the next row (option "o"). */
    _setStyleOption(pos: number, style: RowStyleOption = "n"): void {
        if (style[0] === "o" && this.findRow(pos + 1) !== undefined) {
            this._copyStyle(pos + 1, pos, style[1] === "+");
        } else if (style[0] === "i" && this.findRow(pos - 1) !== undefined) {
            this._copyStyle(pos - 1, pos, style[1] === "+");
        }
    }

    /** Copies the style (and height) of row `src` to row `dest`: the row style and the cells' (empty ones too with `styleEmpty`). */
    _copyStyle(src: number, dest: number, styleEmpty = false): void {
        const rSrc = this.getRow(src);
        const rDst = this.getRow(dest);
        rDst.style = copyStyle(rSrc.style);

        rSrc.eachCell({ includeEmpty: styleEmpty }, (cell, colNumber) => {
            rDst.getCell(colNumber).style = copyStyle(cell.style);
        });
        rDst.height = rSrc.height;
    }

    /**
     * Duplicates row `rowNum` `count` times below it: values, styles, height, and its single-row merges; the copies are
     * inserted (`insert`) or overwrite the rows below.
     * @throws Error when the merges would exceed the cap
     */
    duplicateRow(rowNum: number, count: number, insert = false): void {
        // create count duplicates of rowNum
        // either inserting new or overwriting existing rows

        // a row that does not exist (spliceColumns no longer creates the empty rows) is duplicated as an empty row
        const rSrc = this.getRow(rowNum);
        const inserts: RowValues[] = new Array<RowValues>(count).fill(rSrc.values);
        // the single-row merges of the row are copied too
        const rowMerges = Object.values(this._merges).filter((merge) => merge.top === rowNum && merge.bottom === rowNum);
        const copiedCells = rowMerges.reduce((cells, merge) => cells + merge.count, 0) * count;
        // the values of the row in a merge are the master's: the copies lose nothing when the merge widens over them
        this._spliceRows(rowNum + 1, insert ? 0 : count, inserts, copiedCells, true);

        // now copy styles...
        for (let i = 0; i < count; i++) {
            // the splice created the rows of the copies
            const rDst = this.getRow(rowNum + i + 1);
            rDst.style = rSrc.style;
            rDst.height = rSrc.height;

            rSrc._eachExistingCell((cell, colNumber) => {
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

    /**
     * Removes `count` rows from row `start` (1-based) and inserts rows with the `inserts` values. The cells (values,
     * styles, notes), merges and defined names move.
     * @throws Error when a value would be inserted inside a merge, or the merges would exceed the cap
     */
    spliceRows(start: number, count: number, ...inserts: (RowValues | null | undefined)[]): void {
        this._spliceRows(start, count, inserts);
    }

    /**
     * @param extraMergedCells - merged cells the caller adds after the splice (checked against the cap with the moved
     *   merges)
     * @param copiesOfMergedRow - the inserts are copies of a row, whose values inside a merge are the master's
     */
    _spliceRows(
        start: number,
        count: number,
        inserts: readonly (RowValues | null | undefined)[],
        extraMergedCells = 0,
        copiesOfMergedRow = false
    ): void {
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
            const insertedValues: ((col: number) => CellValue | undefined)[] = [];
            for (const range of merges.moved) {
                for (let row = Math.max(start, range.top + 1); row <= Math.min(start + inserts.length - 1, range.bottom); row++) {
                    const valueAt =
                        insertedValues[row - start] || (insertedValues[row - start] = insertedRowValues(this, inserts[row - start]));
                    for (let col = range.left; col <= range.right; col++) {
                        checkInsertedValue(valueAt(col), row, col, range);
                    }
                }
            }
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

                    rSrc._eachExistingCell((cell, colNumber) => {
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
            for (const r of existing) {
                if (r >= nClearStart && r <= nClearEnd) {
                    this._rows[r - 1] = undefined;
                }
            }
        } else if (nExpand > 0) {
            // insert new cells
            for (i of shiftSteps(existing, nKeep, nEnd, nExpand).reverse()) {
                rSrc = this._rows[i - 1];
                if (rSrc) {
                    const rDst = this.getRow(i + nExpand);
                    rDst.values = rSrc.values;
                    rDst.style = rSrc.style;
                    rDst.height = rSrc.height;

                    rSrc._eachExistingCell((cell, colNumber) => {
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

    /**
     * Calls `iteratee` for each row that has values, in order (with `{ includeEmpty: true }`: for every row up to the
     * last one, creating the missing ones).
     */
    eachRow(iteratee: RowIteratee): void;
    eachRow(options: EachRowOptions | null | undefined, iteratee: RowIteratee): void;
    eachRow(optionsOrIteratee: EachRowOptions | RowIteratee | null | undefined, iteratee?: RowIteratee): void {
        // eachRow(iteratee) or eachRow(options, iteratee)
        const options = typeof optionsOrIteratee === "function" ? undefined : optionsOrIteratee;
        const callback = typeof optionsOrIteratee === "function" ? optionsOrIteratee : iteratee;
        if (typeof callback !== "function") {
            throw new Error("eachRow needs an iteratee function");
        }
        if (options && options.includeEmpty) {
            const n = this._rows.length;
            for (let i = 1; i <= n; i++) {
                callback(this.getRow(i), i);
            }
        } else {
            eachEntry(this._rows, (row) => {
                if (row.hasValues) {
                    callback(row, row.number);
                }
            });
        }
    }

    /** The values of all the rows, as a sparse array indexed by row number (each a sparse array of `row.values`). */
    getSheetValues(): CellValue[][] {
        const rows: CellValue[][] = [];
        eachEntry(this._rows, (row) => {
            rows[row.number] = row.values;
        });
        return rows;
    }

    // =========================================================================
    // Cells

    /**
     * The cell at (row, col) (1-based) or at an address (`"B2"`, `"$B$2"`, `"Sheet1!B2"`) if it exists: `undefined` for
     * an address that is not a cell of this worksheet (invalid, out of bounds, or on another sheet), as for a cell that
     * does not exist.
     */
    findCell(address: string): Cell | undefined;
    findCell(r: number, c: number): Cell | undefined;
    findCell(r: number | string, c?: number): Cell | undefined {
        const address = findCellAddress(this.name, r, c);
        const row = address && this._rows[address.row - 1];
        return row ? row.findCell(address.col) : undefined;
    }

    /**
     * The cell at (row, col) (1-based) or at an address (`"B2"`, `"$B$2"`, `"Sheet1!B2"`: a sheet name must be this
     * worksheet's), created if needed.
     * @throws Error for an invalid address: rows are 1 to 1,048,576, columns 1 (`A`) to 16,384 (`XFD`), the column
     * letters upper case; or a cell of another sheet
     */
    getCell(address: string): Cell;
    getCell(r: number, c: number): Cell;
    getCell(r: number | string, c?: number): Cell {
        const address = resolveCellAddress(this.name, r, c);
        const row = this.getRow(address.row);
        return row.getCellEx(address);
    }

    // =========================================================================
    // Merge

    /**
     * Merges a range (`"A1:B2"`, `"A1", "B2"`, `1, 1, 2, 2`...: see {@link SheetRangeArgs}) into its top-left cell, which
     * gives its style to the other cells. A sheet name in the arguments must be this worksheet's (case-insensitive).
     * @throws Error for an invalid range or a range of another sheet (nothing is merged), when a cell is already merged,
     * or when the merges would exceed the cap
     */
    mergeCells(...cells: SheetRangeArgs): void {
        this._mergeCellsInternal(resolveRange(this.name, "merge", cells));
    }

    /**
     * Merges a range into its top-left cell, the cells keeping their style (see {@link Worksheet.mergeCells}).
     * @throws Error as {@link Worksheet.mergeCells}
     */
    mergeCellsWithoutStyle(...cells: SheetRangeArgs): void {
        this._mergeCellsInternal(resolveRange(this.name, "merge", cells), true);
    }

    /** @throws Error when a cell is already merged, or the merges would exceed the cap */
    _mergeCellsInternal(dimensions: Range, ignoreStyle?: boolean): void {
        // Every cell of a merge is materialised, so the total merged area is capped (utils/merge-limits)
        this._checkMergedCells(`Cannot merge ${dimensions.range}`, this._mergedCellCount + dimensions.count);

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
        this._addMergedCells(dimensions.count);
    }

    /**
     * Splices move merges as rectangles: the merges that `isTouched`, and where `move(merge)` puts each of them
     * ([top, left, bottom, right], or nothing when the splice removes it). A merge reduced to a single cell is not a merge
     * any more.
     * @throws Error before anything changes when the moved merges (and `extraCells` more merged cells) would exceed the cap
     */
    _planMerges(
        isTouched: (merge: Range) => boolean,
        move: (merge: Range) => [number, number, number, number] | undefined,
        extraCells = 0
    ): { touched: Range[]; moved: Range[] } {
        const touched = Object.values(this._merges).filter(isTouched);
        const moved = touched
            .map(move)
            .filter((span): span is [number, number, number, number] => !!span && (span[0] !== span[2] || span[1] !== span[3]))
            .map((span) => new Range(...span));
        const count = [
            this._mergedCellCount + extraCells,
            ...touched.map((merge) => -merge.count),
            ...moved.map((range) => range.count),
        ].reduce((sum, n) => sum + n);
        this._checkMergedCells("Cannot splice", count);
        return { touched, moved };
    }

    /** The workbook that holds this worksheet (not one it was removed from). */
    get _holdingWorkbook(): Workbook | undefined {
        const workbook = this._workbook;
        return workbook && workbook._worksheets[this.id] === this ? workbook : undefined;
    }

    /** @throws Error when the worksheet would hold `count` merged cells: more than a worksheet or its workbook may hold */
    _checkMergedCells(what: string, count: number): void {
        const workbook = this._holdingWorkbook;
        if (workbook) {
            // the workbook keeps the total of its worksheets: summing them on each merge was O(merges x worksheet ids)
            checkMergedCells(what, count, workbook._mergedCellCount - this._mergedCellCount + count, workbook.maxMergedCells);
        } else {
            checkMergedCells(what, count, count, Infinity);
        }
    }

    /** The worksheet holds `delta` more merged cells, and so does its workbook. */
    _addMergedCells(delta: number): void {
        this._mergedCellCount += delta;
        const workbook = this._holdingWorkbook;
        if (workbook) {
            workbook._mergedCellCount += delta;
        }
    }

    /**
     * The merges are released before the cells move, each cell keeping its own value (empty for all but the master) and
     * style, so that rows and columns move as plain cells...
     */
    _detachMerges(merges: readonly Range[]): void {
        for (const merge of merges) {
            for (let i = merge.top; i <= merge.bottom; i++) {
                for (let j = merge.left; j <= merge.right; j++) {
                    const cell = this.findCell(i, j);
                    if (cell) {
                        cell.unmerge(true);
                    }
                }
            }
            delete this._merges[merge.tl];
            this._addMergedCells(-merge.count);
        }
    }

    /** ...then merged again at their new place. */
    _attachMerges(ranges: readonly Range[]): void {
        ranges.forEach((range) => this._mergeCellsInternal(range, true));
    }

    /** Unmerges the merge of a master (top-left) cell. */
    _unMergeMaster(master: Cell): void {
        // master is always top left of a rectangle
        const merge = this._merges[master.address];
        if (merge) {
            for (let i = merge.top; i <= merge.bottom; i++) {
                for (let j = merge.left; j <= merge.right; j++) {
                    this.getCell(i, j).unmerge();
                }
            }
            delete this._merges[master.address];
            this._addMergedCells(-merge.count);
        }
    }

    /** Whether the worksheet has a merge. */
    get hasMerges(): boolean {
        // return true if this._merges has a merge object
        return Object.values(this._merges).some(Boolean);
    }

    /**
     * Unmerges the merges that have a cell in a range (see {@link Worksheet.mergeCells}): merges are atomic, either
     * they're all merged or all un-merged. Unmerging breaks the style links.
     * @throws Error for an invalid range or a range of another sheet (nothing is unmerged)
     */
    unMergeCells(...cells: SheetRangeArgs): void {
        const dimensions = resolveRange(this.name, "unmerge", cells);

        // find any cells in that range and unmerge them
        for (let i = dimensions.top; i <= dimensions.bottom; i++) {
            for (let j = dimensions.left; j <= dimensions.right; j++) {
                const cell = this.findCell(i, j);
                if (cell) {
                    if (cell.type === ValueType.Merge) {
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
    /**
     * Sets a formula on a range (`"A1:B3"`, `"C3"`; a sheet name must be this worksheet's): the top-left cell gets it (a
     * shared formula, or an array formula with `shareType` "array") with the range as `ref` (`"A1:B3"`, `"C3"`, without
     * sheet name nor `$`), the other cells are its clones (a shared formula) or get their result (an array formula).
     * @param results - the results of the cells (by cell, by row then column, or in a flat array)
     * @throws Error for an invalid range or a range of another sheet (nothing is filled)
     */
    fillFormula(range: string, formula: string, results?: FillFormulaResults, shareType: "shared" | "array" = "shared"): void {
        // Define formula for top-left cell and share to rest
        const bounds = resolveRange(this.name, "fill a formula", [range]);
        const { top, left, bottom, right } = bounds;
        const ref = bounds.count > 1 ? bounds.range : bounds.tl;
        const width = right - left + 1;
        const masterAddress = colCache.encodeAddress(top, left);
        const isShared = shareType === "shared";

        // work out result accessor
        let getResult: (row: number, col: number) => FormulaResult | undefined;
        if (typeof results === "function") {
            getResult = results;
        } else if (results) {
            if (isResultRows(results)) {
                // a row missing from the array has no results, as a missing cell
                getResult = (row, col) => results[row - top]?.[col - left];
            } else {
                getResult = (row, col) => results[(row - top) * width + (col - left)];
            }
        } else {
            getResult = () => undefined;
        }
        // every result is read (once) and checked before any cell is filled: an invalid one changes nothing
        const cellResults: (FormulaResult | undefined)[] = [];
        for (let r = top; r <= bottom; r++) {
            for (let c = left; c <= right; c++) {
                const result: unknown = getResult(r, c);
                if (result === undefined || result === null) {
                    cellResults.push(undefined);
                } else if (isFormulaResult(result)) {
                    cellResults.push(result);
                } else {
                    throw new Error(
                        `Cannot fill a formula: the result of ${colCache.encodeAddress(r, c)} must be a string, a number, a boolean, a Date or an error value`
                    );
                }
            }
        }
        let index = 0;
        for (let r = top; r <= bottom; r++) {
            for (let c = left; c <= right; c++) {
                const result = cellResults[index];
                if (index === 0) {
                    this.getCell(r, c).value = { shareType, formula, ref, result };
                } else {
                    this.getCell(r, c).value = isShared ? ({ sharedFormula: masterAddress, result } satisfies FormulaValueObject) : result;
                }
                index++;
            }
        }
    }

    // =========================================================================
    // Images
    /** Adds an image of the workbook (`workbook.addImage` id) over a range (`"B2:D6"`) or at a position. */
    addImage(imageId: number | string, range: string | ImageRangeInput): void {
        this._media.push(new Image(this, { type: "image", imageId, range }));
    }

    /** The images of the worksheet (not the background). */
    getImages(): Image[] {
        return this._media.filter((m) => m.type === "image");
    }

    /** Sets an image of the workbook (`workbook.addImage` id) as the background. */
    addBackgroundImage(imageId: number | string): void {
        this._media.push(new Image(this, { type: "background", imageId }));
    }

    /** The id of the background image, if any. */
    getBackgroundImageId(): number | string | undefined {
        const image = this._media.find((m) => m.type === "background");
        return image && image.imageId;
    }

    // =========================================================================
    // Worksheet Protection
    /**
     * Protects the worksheet, with a password (hashed with SHA-512 and a random salt, `spinCount` times: default 100000)
     * or without. `options.spinCount` is normalised in place (an integer >= 0, 100000 when not a finite number).
     */
    protect(password?: string, options?: Partial<WorksheetProtection>): Promise<void> {
        // TODO: make this function truly async
        // perhaps marshal to worker thread or something
        return new Promise((resolve) => {
            const protection: SheetProtectionModel = {
                sheet: true,
            };
            this.sheetProtection = protection;
            if (options && "spinCount" in options) {
                // force spinCount to be integer >= 0
                const { spinCount } = options;
                options.spinCount =
                    typeof spinCount === "number" && Number.isFinite(spinCount) ? Math.round(Math.max(0, spinCount)) : 100000;
            }
            if (password) {
                protection.algorithmName = "SHA-512";
                protection.saltValue = Encryptor.randomBytes(16).toString("base64");
                // allow user specified spinCount (normalised above: a number when present)
                const spinCount = options?.spinCount ?? 100000;
                protection.spinCount = spinCount;
                protection.hashValue = Encryptor.convertPasswordToHash(password, "SHA512", protection.saltValue, spinCount);
            }
            if (options) {
                this.sheetProtection = Object.assign(protection, options);
                if (!password && "spinCount" in options) {
                    delete protection.spinCount;
                }
            }
            resolve();
        });
    }

    /** Removes the protection. */
    unprotect(): void {
        this.sheetProtection = null;
    }

    // =========================================================================
    // Tables
    /**
     * Adds a table (written in the cells of the worksheet).
     * @throws Error for invalid table properties
     */
    addTable(model: TableModel): Table {
        const table = new Table(this, model);
        setTable(this.tables, model.name, table);
        return table;
    }

    /** A table by name, if any. */
    getTable(name: string): Table | undefined {
        return Object.hasOwn(this.tables, name) ? this.tables[name] : undefined;
    }

    /** Removes a table (its cells keep their values). */
    removeTable(name: string): void {
        delete this.tables[name];
    }

    /** The tables. */
    getTables(): Table[] {
        return Object.values(this.tables);
    }

    // =========================================================================
    // Pivot Tables
    /**
     * Adds a pivot table (experimental: one per workbook).
     * @throws Error for an invalid pivot table (see makePivotTable)
     */
    addPivotTable(model: PivotTableInput): PivotTable {
        console.warn(
            // (the message has a space at the end of its first line)
            "Warning: Pivot Table support is experimental. \nPlease leave feedback at https://github.com/exceljs/exceljs/discussions/2575"
        );

        const pivotTable = makePivotTable(this, model);

        this.pivotTables.push(pivotTable);
        this.workbook.pivotTables.push(pivotTable);

        return pivotTable;
    }

    // ===========================================================================
    // Conditional Formatting
    /** Adds conditional formatting rules to a range. */
    addConditionalFormatting(cf: ConditionalFormattingModel): void {
        this.conditionalFormattings.push(cf);
    }

    /**
     * Removes conditional formattings: the one at an index, the ones a filter rejects (keeps those it returns true for),
     * or all of them (any other argument).
     */
    removeConditionalFormatting(
        filter?: number | ((cf: ConditionalFormattingModel, index: number, all: ConditionalFormattingModel[]) => unknown) | null
    ): void {
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
    /** @deprecated use `properties.tabColor` */
    get tabColor(): WorksheetPropertiesModel["tabColor"] {
        // eslint-disable-next-line no-console -- a deprecation notice with its stack trace, as it always was
        console.trace("worksheet.tabColor property is now deprecated. Please use worksheet.properties.tabColor");
        return this.properties.tabColor;
    }

    /** @deprecated use `properties.tabColor` */
    set tabColor(value: WorksheetPropertiesModel["tabColor"]) {
        // eslint-disable-next-line no-console -- a deprecation notice with its stack trace, as it always was
        console.trace("worksheet.tabColor property is now deprecated. Please use worksheet.properties.tabColor");
        this.properties.tabColor = value;
    }

    // ===========================================================================
    // Model

    /**
     * The model of the worksheet (for the xlsx writer): the settings by reference, the rows and columns as models, the
     * used range (`dimensions`) and the merged ranges (`merges`).
     */
    get model(): DocumentWorksheetModel {
        // ==========================================================
        // Rows
        const rows: RowModel[] = [];
        const dimensions = new Range();
        const merges: string[] = [];
        const model: DocumentWorksheetModel = {
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
            // =================================================
            // columns
            cols: this._colsModel(),
            rows,
            dimensions,
            merges,
        };

        eachEntry(this._rows, (row) => {
            const rowModel = row.model;
            if (rowModel) {
                dimensions.expand(rowModel.number, rowModel.min, rowModel.number, rowModel.max);
                rows.push(rowModel);
            }
        });

        // ==========================================================
        // Merges
        for (const merge of Object.values(this._merges)) {
            merges.push(merge.range);
        }

        return model;
    }

    /** Replaces the rows with those of the model. */
    _parseRows(model: WorksheetModel): void {
        this._rows = [];
        model.rows.forEach((rowModel) => {
            const row = new Row(this, rowModel.number);
            this._rows[row.number - 1] = row;
            row.model = rowModel;
        });
    }

    /**
     * Replaces the merges with those of the model, the cells keeping their style: `mergeCells` (a model read from a
     * file), else `merges` (the getter's).
     * @throws Error for an invalid range, or a range of another sheet
     */
    _parseMergeCells(model: WorksheetModel): void {
        // the merges of the model replace those of the worksheet (whose rows were replaced)
        this._merges = {};
        this._addMergedCells(-this._mergedCellCount);
        (model.mergeCells ?? model.merges)?.forEach((merge) => {
            // Do not merge styles when importing an Excel file
            // since each cell may have different styles intentionally.
            this.mergeCellsWithoutStyle(merge);
        });
    }

    /**
     * Sets the worksheet from a model (the settings are adopted by reference). The columns of loaded cols are created on
     * demand, unless a col has a key or a header.
     * @throws Error for an invalid name, or when the merges exceed the cap
     */
    set model(value: WorksheetModel) {
        this.name = value.name;
        if ((value.cols || []).some((col: ColumnDefn) => col.key || col.header)) {
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
        // a model without page breaks has none (one built by hand: the getter and a file read always give the list)
        this.rowBreaks = value.rowBreaks ?? [];
        this.views = value.views;
        this.autoFilter = value.autoFilter;
        this._media = value.media.map((medium) => new Image(this, medium));
        this.sheetProtection = value.sheetProtection;
        this.tables = value.tables.reduce<Record<string, Table>>((tables, table) => {
            const t = new Table(this);
            t.model = table;
            setTable(tables, table.name, t);
            return tables;
        }, {});
        // a loaded worksheet has no pivot table
        this.pivotTables = value.pivotTables || [];
        this.conditionalFormattings = value.conditionalFormattings;
    }
}

export default Worksheet;
