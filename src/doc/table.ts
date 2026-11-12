// A table of a worksheet (`worksheet.addTable`): a header row, data rows and an optional totals row, written in the
// cells of the worksheet and described by its model (written as `xl/tables/tableN.xml`).

import colCache from "../utils/col-cache";
import type { TableColumnModel, TableModel } from "../types";
import type Cell from "./cell";
import type { DxfModel, StyleModel } from "../types";

/** What a {@link Table} reads of its worksheet: its rows, to write the table in their cells. */
export interface TableWorksheet {
    getRow(row: number): { getCell(col: number): Cell };
}

/** A data row of a table: the values of its cells, by column (`undefined` for a missing one). */
type TableRow = NonNullable<TableModel["rows"]>[number];

/** The position of the top-left cell of a table (1-based). */
interface TablePosition {
    row: number;
    col: number;
}

/** What a table created by `addTable` has, and a table read from a file has not: its worksheet, data rows and position. */
interface EditableTable {
    worksheet: TableWorksheet;
    rows: TableRow[];
    tl: TablePosition;
}

/** The state of a table before a change, to clear its previous footprint on commit. */
interface TableCache {
    tl: TablePosition;
    width: number;
    tableHeight: number;
}

/** The totals row functions, with their SUBTOTAL function number (`none`, `custom`: none). */
const SUBTOTAL_FUNCTIONS: Readonly<Record<NonNullable<TableColumnModel["totalsRowFunction"]>, number | null>> = {
    none: null,
    custom: null,
    average: 101,
    countNums: 102,
    count: 103,
    max: 104,
    min: 105,
    stdDev: 106,
    var: 107,
    sum: 109,
};

/** Whether a value is a totals row function (`"sum"`...). */
function isTotalsRowFunction(value: unknown): value is keyof typeof SUBTOTAL_FUNCTIONS {
    return typeof value === "string" && Object.prototype.hasOwnProperty.call(SUBTOTAL_FUNCTIONS, value);
}

/** `count` and the noun, plural unless the count is 1 (`"1 column"`, `"2 columns"`). */
function plural(count: number, noun: string): string {
    return `${count} ${noun}${count === 1 ? "" : "s"}`;
}

/** @throws Error with the message when the test is falsy */
function assert(test: unknown, message: string): asserts test {
    if (!test) {
        throw new Error(message);
    }
}

/** Sets `o[name]` to `dflt` when it is undefined. */
function assign<T extends object, K extends keyof T>(o: T, name: K, dflt: T[K]): void {
    if (o[name] === undefined) {
        o[name] = dflt;
    }
}

/**
 * Checks a data row of a table.
 * @param rowNumber - its 1-based number among the data rows (for the message)
 * @throws Error for a row that is not an array, or that has more values than the table has columns
 */
function checkRow(name: string, values: unknown, rowNumber: number, width: number): asserts values is TableRow {
    assert(Array.isArray(values), `Table "${name}": row ${rowNumber} is not an array of values`);
    assert(
        values.length <= width,
        `Table "${name}": row ${rowNumber} has ${plural(values.length, "value")}, but the table has ${plural(width, "column")}`
    );
}

/** Copies the properties of a column style into the style of a cell (no copy on write: as it always did). */
function assignStyle(cell: Cell, style: StyleModel | DxfModel | undefined): void {
    if (style) {
        Object.assign(cell.style, style);
    }
}

/** A column of a table (`table.getColumn(i)`): reads and writes the column model; most changes need `table.commit()`. */
class Column {
    /** The table. */
    table: Table;
    /** The column model (by reference). */
    column: TableColumnModel;
    /** 0-based index of the column in the table. */
    index: number;

    // wrapper around column model, allowing access and manipulation
    constructor(table: Table, column: TableColumnModel, index: number) {
        this.table = table;
        this.column = column;
        this.index = index;
    }

    /** Sets a field of the column model, after saving the table state (for commit). */
    _set<K extends keyof TableColumnModel>(name: K, value: TableColumnModel[K]): void {
        this.table.cacheState();
        this.column[name] = value;
    }

    /** The column name (its header text); needs `table.commit()` once changed. */
    get name(): string {
        return this.column.name;
    }
    set name(value: string) {
        this._set("name", value);
    }

    /** Whether the column shows the filter button of the auto filter. */
    get filterButton(): boolean | undefined {
        return this.column.filterButton;
    }
    set filterButton(value: boolean | undefined) {
        this.column.filterButton = value;
    }

    /** The style of the column cells (a differential format, written as a dxf of the styles part). */
    get style(): TableColumnModel["style"] {
        return this.column.style;
    }
    set style(value: TableColumnModel["style"]) {
        this.column.style = value;
    }

    /** The text of the totals row in this column (used for the first column); needs `table.commit()`. */
    get totalsRowLabel(): string | undefined {
        return this.column.totalsRowLabel;
    }
    set totalsRowLabel(value: string | undefined) {
        this._set("totalsRowLabel", value);
    }

    /** The function of the totals row (`"sum"`, `"average"`, `"custom"`…); needs `table.commit()`. */
    get totalsRowFunction(): TableColumnModel["totalsRowFunction"] {
        return this.column.totalsRowFunction;
    }
    set totalsRowFunction(value: TableColumnModel["totalsRowFunction"]) {
        this._set("totalsRowFunction", value);
    }

    /** The cached result of the totals row formula; needs `table.commit()`. */
    get totalsRowResult(): TableColumnModel["totalsRowResult"] {
        return this.column.totalsRowResult;
    }
    set totalsRowResult(value: TableColumnModel["totalsRowResult"]) {
        this._set("totalsRowResult", value);
    }

    /** The formula of the totals row for the `"custom"` function; needs `table.commit()`. */
    get totalsRowFormula(): string | undefined {
        return this.column.totalsRowFormula;
    }
    set totalsRowFormula(value: string | undefined) {
        this._set("totalsRowFormula", value);
    }
}

/**
 * A table of a worksheet. Created from index.d.ts `TableProperties` (validated and completed with the defaults, then
 * written in the cells), or from a loaded model (without data rows nor `ref`).
 */
class Table {
    /** The worksheet (none for a table set from a loaded model). */
    worksheet: TableWorksheet | undefined;
    /** The model (by reference: `validate` completes it in place). */
    declare table: TableModel;
    /** The state before the first change since the last commit. */
    declare _cache?: TableCache | undefined;

    /**
     * @param table - the table properties: validated, completed, and written in the cells of the worksheet; nothing for a
     *   table set from a model next
     * @throws Error for invalid table properties (see validate)
     */
    constructor(worksheet?: TableWorksheet, table?: TableModel) {
        this.worksheet = worksheet;
        if (table) {
            this.table = table;
            // check things are ok first
            this.validate();

            this.store();
        }
    }

    /**
     * The formula of the totals row of a column (`SUBTOTAL(109,Table[Column])`), or null for the function "none".
     * @throws Error for an unknown totals row function
     */
    getFormula(column: TableColumnModel): string | null | undefined {
        // get the correct formula to apply to the totals row
        const { totalsRowFunction } = column;
        if (!isTotalsRowFunction(totalsRowFunction)) {
            throw new Error(`Invalid Totals Row Function: ${String(totalsRowFunction)}`);
        }
        if (totalsRowFunction === "custom") {
            return column.totalsRowFormula;
        }
        const subtotal = SUBTOTAL_FUNCTIONS[totalsRowFunction];
        return subtotal === null ? null : `SUBTOTAL(${subtotal},${this.table.name}[${column.name}])`;
    }

    /** The number of columns. */
    get width(): number {
        // width of the table
        return this.table.columns.length;
    }

    /** The number of data rows (of a table read from a file: those of its range, less the header and totals rows). */
    get height(): number {
        const { rows, tableRef, headerRow, totalsRow } = this.table;
        if (rows) {
            return rows.length;
        }
        if (!tableRef) {
            return 0;
        }
        const range = colCache.decode(tableRef);
        const rangeHeight = "bottom" in range ? range.bottom - range.top + 1 : 1;
        return Math.max(0, rangeHeight - (headerRow ? 1 : 0) - (totalsRow ? 1 : 0));
    }

    /** The height of the data plus the header row, if any. */
    get filterHeight(): number {
        // height of the table data plus optional header row
        return this.height + (this.table.headerRow ? 1 : 0);
    }

    /** The full height of the table in the worksheet (header and totals rows included). */
    get tableHeight(): number {
        // full height of the table on the sheet
        return this.filterHeight + (this.table.totalsRow ? 1 : 0);
    }

    /**
     * Checks the model, then completes it with the defaults (header row, no totals row, style "TableStyleMedium2"...),
     * the position (`tl`, `autoFilterRef`, `tableRef`) and the totals row formulas. Nothing is changed when it throws.
     * @throws Error without ref, columns or rows, for a ref outside the sheet, for a column without name, for an unknown
     *   totals row function, for a row that is not an array or has more values than the table has columns, for a table
     *   that does not fit in the sheet
     */
    validate(): void {
        this._validate();
    }

    /** {@link validate}, returning the position of the table. */
    _validate(): TablePosition {
        const { table } = this;
        const { name } = table;
        assert(table.ref, "Table must have ref");
        assert(table.columns, "Table must have column definitions");
        assert(table.rows, "Table must have row definitions");
        const tl = colCache.decodeAddress(table.ref);
        const { row, col } = tl;
        assert(row > 0, "Table must be on valid row");
        assert(col > 0, "Table must be on valid col");
        table.columns.forEach((column, i) => {
            assert(column.name, `Column ${i} must have a name`);
            if (i > 0 && column.totalsRowFunction !== undefined) {
                // throws for an unknown function
                this.getFormula(column);
            }
        });
        const { width } = this;
        table.rows.forEach((values: unknown, i) => {
            checkRow(name, values, i + 1, width);
        });
        const filterHeight = this.height + (table.headerRow === false ? 0 : 1);
        const tableHeight = filterHeight + (table.totalsRow ? 1 : 0);
        // autoFilterRef is a range that includes optional headers only; tableRef includes headers and totals (both throw
        // for a table that does not fit in the sheet, before any change)
        const autoFilterRef = colCache.encode(row, col, row + filterHeight - 1, col + width - 1);
        const tableRef = colCache.encode(row, col, row + tableHeight - 1, col + width - 1);

        // set defaults
        assign(table, "headerRow", true);
        assign(table, "totalsRow", false);
        const style = this._style();
        assign(style, "theme", "TableStyleMedium2");
        assign(style, "showFirstColumn", false);
        assign(style, "showLastColumn", false);
        assign(style, "showRowStripes", false);
        assign(style, "showColumnStripes", false);

        table.tl = tl;
        table.autoFilterRef = autoFilterRef;
        table.tableRef = tableRef;

        table.columns.forEach((column, i) => {
            if (i === 0) {
                assign(column, "totalsRowLabel", "Total");
            } else {
                assign(column, "totalsRowFunction", "none");
                // getFormula gives null (no formula) for the function "none"
                column.totalsRowFormula = this.getFormula(column) ?? undefined;
            }
        });
        return { row, col };
    }

    /**
     * Writes the table in the cells of the worksheet: headers, data (with the column styles), totals row.
     * @throws Error for a table read from a file (see {@link commit})
     */
    store(): void {
        const { worksheet, rows, tl } = this._editable();
        this._write(worksheet, tl, rows, true);
    }

    /**
     * Writes the headers, data and totals of the table in the cells of a worksheet (without styles; unused).
     * @throws Error for a table read from a file (see {@link commit})
     */
    load(worksheet: TableWorksheet): void {
        const { rows, tl } = this._editable();
        this._write(worksheet, tl, rows, false);
    }

    /**
     * Writes the headers, data and totals row of the table in the cells of a worksheet: with the column styles (and
     * clearing the totals cells without formula), or without (`load`).
     */
    _write(worksheet: TableWorksheet, tl: TablePosition, rows: readonly TableRow[], withStyles: boolean): void {
        const { table } = this;
        const { row, col } = tl;
        let count = 0;
        if (table.headerRow) {
            const r = worksheet.getRow(row + count++);
            table.columns.forEach((column, j) => {
                const cell = r.getCell(col + j);
                cell.value = column.name;
                if (withStyles) {
                    assignStyle(cell, column.style);
                }
            });
        }
        rows.forEach((data) => {
            const r = worksheet.getRow(row + count++);
            // a row has no more values than columns (validate); the cells past its values are left as they are
            table.columns.forEach((column, j) => {
                if (j < data.length) {
                    const cell = r.getCell(col + j);
                    cell.value = data[j];
                    if (withStyles) {
                        assignStyle(cell, column.style);
                    }
                }
            });
        });

        if (table.totalsRow) {
            const r = worksheet.getRow(row + count++);
            table.columns.forEach((column, j) => {
                const cell = r.getCell(col + j);
                if (j === 0) {
                    cell.value = column.totalsRowLabel;
                } else if (this.getFormula(column)) {
                    cell.value = {
                        formula: column.totalsRowFormula,
                        result: column.totalsRowResult,
                    };
                } else if (withStyles) {
                    cell.value = null;
                }
                if (withStyles) {
                    assignStyle(cell, column.style);
                }
            });
        }
    }

    /**
     * The worksheet, data rows and position of a table created by `addTable`.
     * @throws Error for a table read from a file: it has none of them (its rows are the cells of its worksheet)
     */
    _editable(): EditableTable {
        const { worksheet } = this;
        const { rows, tl } = this.table;
        if (!worksheet || !rows || !tl) {
            throw new Error(`Table "${this.table.name}" was read from a file: change the cells of its worksheet`);
        }
        return { worksheet, rows, tl };
    }

    /** The style of the model, created (empty) for a table without style (one read from a file without table style). */
    _style(): TableStyle {
        const { table } = this;
        if (!table.style) {
            table.style = {};
        }
        return table.style;
    }

    /** The model (by reference). */
    get model(): TableModel {
        return this.table;
    }

    set model(value: TableModel) {
        this.table = value;
    }

    // ================================================================
    // Mutating methods: a table read from a file has no data rows (they are the cells of its worksheet), they throw

    /**
     * Saves the position and size of the table before a change (once until the next commit).
     * @throws Error for a table read from a file
     */
    cacheState(): void {
        const { tl } = this._editable();
        if (!this._cache) {
            this._cache = {
                tl: { row: tl.row, col: tl.col },
                width: this.width,
                tableHeight: this.tableHeight,
            };
        }
    }

    /**
     * Applies the changes made since the last commit: validates the table, clears the cells it no longer covers, and
     * writes it again. Does nothing without a change (a table read from a file cannot be changed).
     * @throws Error for invalid table properties (see validate): nothing is changed then
     */
    commit(): void {
        // changes may have been made that might have on-sheet effects
        const cache = this._cache;
        if (!cache) {
            return;
        }
        const { worksheet, rows } = this._editable();

        // check things are ok first
        const tl = this._validate();

        const previous = cache.tl;
        if (tl.row !== previous.row || tl.col !== previous.col) {
            // wipe out whole table footprint at previous location
            for (let i = 0; i < cache.tableHeight; i++) {
                const row = worksheet.getRow(previous.row + i);
                for (let j = 0; j < cache.width; j++) {
                    const cell = row.getCell(previous.col + j);
                    cell.value = null;
                }
            }
        } else {
            // clear out below table if it has shrunk
            for (let i = this.tableHeight; i < cache.tableHeight; i++) {
                const row = worksheet.getRow(previous.row + i);
                for (let j = 0; j < cache.width; j++) {
                    const cell = row.getCell(previous.col + j);
                    cell.value = null;
                }
            }

            // clear out to right of table if it has lost columns
            for (let i = 0; i < this.tableHeight; i++) {
                const row = worksheet.getRow(previous.row + i);
                for (let j = this.width; j < cache.width; j++) {
                    const cell = row.getCell(previous.col + j);
                    cell.value = null;
                }
            }
        }

        this._write(worksheet, tl, rows, true);
        // the next change starts from the table as committed
        this._cache = undefined;
    }

    /**
     * Adds a row of data, inserted at `rowNumber` (0-based) or appended (needs `commit()`).
     * @throws Error for values that are not an array, or more values than the table has columns (nothing is changed);
     *   for a table read from a file
     */
    addRow(values: TableRow, rowNumber?: number): void {
        const { rows } = this._editable();
        const index = rowNumber === undefined ? rows.length : rowNumber;
        checkRow(this.table.name, values, Math.min(Math.max(index, 0), rows.length) + 1, this.width);
        this.cacheState();
        rows.splice(index, 0, values);
    }

    /**
     * Removes `count` rows of data from `rowIndex` (0-based; needs `commit()`).
     * @throws Error for a table read from a file
     */
    removeRows(rowIndex: number, count = 1): void {
        const { rows } = this._editable();
        this.cacheState();
        rows.splice(rowIndex, count);
    }

    /**
     * A column of the table.
     * @param colIndex - 0-based
     * @throws Error for an index that is not a column of the table
     */
    getColumn(colIndex: number): Column {
        const column = Number.isInteger(colIndex) && colIndex >= 0 ? this.table.columns[colIndex] : undefined;
        if (!column) {
            throw new Error(`Table "${this.table.name}" has no column ${colIndex} (0-based; it has ${plural(this.width, "column")})`);
        }
        return new Column(this, column, colIndex);
    }

    /**
     * Adds a column with its values (one per data row), inserted at `colIndex` (0-based) or appended (needs `commit()`).
     * @throws Error for more values than the table has rows (nothing is changed); for a table read from a file
     */
    addColumn(column: TableColumnModel, values: TableRow, colIndex?: number): void {
        const { rows } = this._editable();
        assert(
            values.length <= rows.length,
            `Table "${this.table.name}": ${plural(values.length, "value")} for a column, but the table has ${plural(rows.length, "row")}`
        );
        this.cacheState();

        if (colIndex === undefined) {
            this.table.columns.push(column);
            rows.forEach((row, i) => {
                row.push(values[i]);
            });
        } else {
            this.table.columns.splice(colIndex, 0, column);
            rows.forEach((row, i) => {
                row.splice(colIndex, 0, values[i]);
            });
        }
    }

    /**
     * Removes `count` columns from `colIndex` (0-based) with their data (needs `commit()`).
     * @throws Error for a table read from a file
     */
    removeColumns(colIndex: number, count = 1): void {
        const { rows } = this._editable();
        this.cacheState();

        this.table.columns.splice(colIndex, count);
        rows.forEach((row) => {
            row.splice(colIndex, count);
        });
    }

    /** Sets a field of the model, after saving the table state (for commit). */
    _assign<K extends keyof TableModel>(target: TableModel, prop: K, value: TableModel[K]): void {
        this.cacheState();
        target[prop] = value;
    }

    /** The top-left cell (`"B2"`); needs `commit()` once changed. */
    get ref(): string | undefined {
        return this.table.ref;
    }
    set ref(value: string | undefined) {
        this._assign(this.table, "ref", value);
    }

    /** The table name (unique in the workbook, used in structured references). */
    get name(): string {
        return this.table.name;
    }
    set name(value: string) {
        this.table.name = value;
    }

    /** The display name (the name when there is none). */
    get displayName(): string {
        return this.table.displayName || this.table.name;
    }
    set displayName(value: string | undefined) {
        this.table.displayName = value;
    }

    /** @deprecated misspelled: use `displayName` */
    set displayNamename(value: string | undefined) {
        this.table.displayName = value;
    }

    /** Whether the table has a header row; needs `commit()` once changed. */
    get headerRow(): boolean | undefined {
        return this.table.headerRow;
    }
    set headerRow(value: boolean | undefined) {
        this._assign(this.table, "headerRow", value);
    }

    /** Whether the table has a totals row; needs `commit()` once changed. */
    get totalsRow(): boolean | undefined {
        return this.table.totalsRow;
    }
    set totalsRow(value: boolean | undefined) {
        this._assign(this.table, "totalsRow", value);
    }

    // the style is set by validate(); a table read from a file without table style has none (`style: null`)

    /** The table style (`"TableStyleMedium2"`), or null for none (also for a table read from a file without table style). */
    get theme(): TableStyle["theme"] {
        return this.table.style ? this.table.style.theme : null;
    }
    set theme(value: TableStyle["theme"]) {
        this._style().theme = value;
    }

    /** Whether the table style highlights the first column. */
    get showFirstColumn(): boolean | undefined {
        return this.table.style ? this.table.style.showFirstColumn : undefined;
    }
    set showFirstColumn(value: boolean | undefined) {
        this._style().showFirstColumn = value;
    }

    /** Whether the table style highlights the last column. */
    get showLastColumn(): boolean | undefined {
        return this.table.style ? this.table.style.showLastColumn : undefined;
    }
    set showLastColumn(value: boolean | undefined) {
        this._style().showLastColumn = value;
    }

    /** Whether the table style bands the rows. */
    get showRowStripes(): boolean | undefined {
        return this.table.style ? this.table.style.showRowStripes : undefined;
    }
    set showRowStripes(value: boolean | undefined) {
        this._style().showRowStripes = value;
    }

    /** Whether the table style bands the columns. */
    get showColumnStripes(): boolean | undefined {
        return this.table.style ? this.table.style.showColumnStripes : undefined;
    }
    set showColumnStripes(value: boolean | undefined) {
        this._style().showColumnStripes = value;
    }
}

/** The style of a table model. */
type TableStyle = NonNullable<TableModel["style"]>;

export default Table;
