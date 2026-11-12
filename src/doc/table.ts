// A table of a worksheet (`worksheet.addTable`): a header row, data rows and an optional totals row, written in the
// cells of the worksheet and described by its model (written as `xl/tables/tableN.xml`).

import colCache from "../utils/col-cache";
import type { TableColumnModel, TableModel } from "../types";
import type Cell from "./cell";
import type { CellValue } from "./cell";
import type { DxfModel, StyleModel } from "../types";

/** What a {@link Table} reads of its worksheet: its rows, to write the table in their cells. */
export interface TableWorksheet {
    getRow(row: number): { getCell(col: number): Cell };
}

/** The state of a table before a change, to clear its previous footprint on commit. */
interface TableCache {
    ref: string | undefined;
    width: number;
    tableHeight: number;
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
 * The data rows of a table. A loaded table has none (nor `ref`): the members that need them throw, as they always did.
 */
function rowsOf(table: TableModel): unknown[][] {
    return table.rows as unknown[][];
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
        switch (column.totalsRowFunction) {
            case "none":
                return null;
            case "average":
                return `SUBTOTAL(101,${this.table.name}[${column.name}])`;
            case "countNums":
                return `SUBTOTAL(102,${this.table.name}[${column.name}])`;
            case "count":
                return `SUBTOTAL(103,${this.table.name}[${column.name}])`;
            case "max":
                return `SUBTOTAL(104,${this.table.name}[${column.name}])`;
            case "min":
                return `SUBTOTAL(105,${this.table.name}[${column.name}])`;
            case "stdDev":
                return `SUBTOTAL(106,${this.table.name}[${column.name}])`;
            case "var":
                return `SUBTOTAL(107,${this.table.name}[${column.name}])`;
            case "sum":
                return `SUBTOTAL(109,${this.table.name}[${column.name}])`;
            case "custom":
                return column.totalsRowFormula;
            case undefined:
            default:
                throw new Error(`Invalid Totals Row Function: ${String(column.totalsRowFunction)}`);
        }
    }

    /** The number of columns. */
    get width(): number {
        // width of the table
        return this.table.columns.length;
    }

    /** The number of data rows. */
    get height(): number {
        // height of the table data
        return rowsOf(this.table).length;
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
     * Completes the model with the defaults (header row, no totals row, style "TableStyleMedium2"...), the position
     * (`tl`, `autoFilterRef`, `tableRef`) and the totals row formulas.
     * @throws Error without ref, columns or rows, for a column without name, for an unknown totals row function
     */
    validate(): void {
        const { table } = this;
        // set defaults and check is valid
        assign(table, "headerRow", true);
        assign(table, "totalsRow", false);

        assign(table, "style", {});
        // assigned just above
        const style = table.style!;
        assign(style, "theme", "TableStyleMedium2");
        assign(style, "showFirstColumn", false);
        assign(style, "showLastColumn", false);
        assign(style, "showRowStripes", false);
        assign(style, "showColumnStripes", false);

        assert(table.ref, "Table must have ref");
        assert(table.columns, "Table must have column definitions");
        assert(table.rows, "Table must have row definitions");

        table.tl = colCache.decodeAddress(table.ref);
        const { row, col } = table.tl;
        assert(row > 0, "Table must be on valid row");
        assert(col > 0, "Table must be on valid col");

        const { width, filterHeight, tableHeight } = this;

        // autoFilterRef is a range that includes optional headers only
        table.autoFilterRef = colCache.encode(row, col, row + filterHeight - 1, col + width - 1);

        // tableRef is a range that includes optional headers and totals
        table.tableRef = colCache.encode(row, col, row + tableHeight - 1, col + width - 1);

        table.columns.forEach((column, i) => {
            assert(column.name, `Column ${i} must have a name`);
            if (i === 0) {
                assign(column, "totalsRowLabel", "Total");
            } else {
                assign(column, "totalsRowFunction", "none");
                column.totalsRowFormula = this.getFormula(column) as string | undefined;
            }
        });
    }

    /** Writes the table in the cells of the worksheet: headers, data (with the column styles), totals row. */
    store(): void {
        // where the table needs to store table data, headers, footers in
        // the sheet...
        const { table } = this;
        // a table from properties has a worksheet and its position (validate)
        const worksheet = this.worksheet!;
        const { row, col } = table.tl!;
        let count = 0;
        if (table.headerRow) {
            const r = worksheet.getRow(row + count++);
            table.columns.forEach((column, j) => {
                const { style, name } = column;
                const cell = r.getCell(col + j);
                cell.value = name;
                assignStyle(cell, style);
            });
        }
        rowsOf(table).forEach((data) => {
            const r = worksheet.getRow(row + count++);
            data.forEach((value, j) => {
                const cell = r.getCell(col + j);
                cell.value = value as CellValue;

                assignStyle(cell, table.columns[j]!.style);
            });
        });

        if (table.totalsRow) {
            const r = worksheet.getRow(row + count++);
            table.columns.forEach((column, j) => {
                const cell = r.getCell(col + j);
                if (j === 0) {
                    cell.value = column.totalsRowLabel;
                } else {
                    const formula = this.getFormula(column);
                    if (formula) {
                        cell.value = {
                            formula: column.totalsRowFormula,
                            result: column.totalsRowResult,
                        };
                    } else {
                        cell.value = null;
                    }
                }

                assignStyle(cell, column.style);
            });
        }
    }

    /** Writes the headers, data and totals of the table in the cells of a worksheet (without styles; unused). */
    load(worksheet: TableWorksheet): void {
        // where the table will read necessary features from a loaded sheet
        const { table } = this;
        // validate() set the position
        const { row, col } = table.tl!;
        let count = 0;
        if (table.headerRow) {
            const r = worksheet.getRow(row + count++);
            table.columns.forEach((column, j) => {
                const cell = r.getCell(col + j);
                cell.value = column.name;
            });
        }
        rowsOf(table).forEach((data) => {
            const r = worksheet.getRow(row + count++);
            data.forEach((value, j) => {
                const cell = r.getCell(col + j);
                cell.value = value as CellValue;
            });
        });

        if (table.totalsRow) {
            const r = worksheet.getRow(row + count++);
            table.columns.forEach((column, j) => {
                const cell = r.getCell(col + j);
                if (j === 0) {
                    cell.value = column.totalsRowLabel;
                } else {
                    const formula = this.getFormula(column);
                    if (formula) {
                        cell.value = {
                            formula: column.totalsRowFormula,
                            result: column.totalsRowResult,
                        };
                    }
                }
            });
        }
    }

    /** The model (by reference). */
    get model(): TableModel {
        return this.table;
    }

    set model(value: TableModel) {
        this.table = value;
    }

    // ================================================================
    // TODO: Mutating methods
    /** Saves the position and size of the table before a change (once until the next commit). */
    cacheState(): void {
        if (!this._cache) {
            this._cache = {
                ref: this.ref,
                width: this.width,
                tableHeight: this.tableHeight,
            };
        }
    }

    /**
     * Applies the changes made since the last commit: validates the table, clears the cells it no longer covers, and
     * writes it again. Does nothing without a change.
     * @throws Error for invalid table properties (see validate)
     */
    commit(): void {
        // changes may have been made that might have on-sheet effects
        if (!this._cache) {
            return;
        }

        // check things are ok first
        this.validate();

        // validate() checked that the table has a ref (and a worksheet: it was created from properties)
        const worksheet = this.worksheet!;
        const ref = colCache.decodeAddress(this._cache.ref!);
        if (this.ref !== this._cache.ref) {
            // wipe out whole table footprint at previous location
            for (let i = 0; i < this._cache.tableHeight; i++) {
                const row = worksheet.getRow(ref.row + i);
                for (let j = 0; j < this._cache.width; j++) {
                    const cell = row.getCell(ref.col + j);
                    cell.value = null;
                }
            }
        } else {
            // clear out below table if it has shrunk
            for (let i = this.tableHeight; i < this._cache.tableHeight; i++) {
                const row = worksheet.getRow(ref.row + i);
                for (let j = 0; j < this._cache.width; j++) {
                    const cell = row.getCell(ref.col + j);
                    cell.value = null;
                }
            }

            // clear out to right of table if it has lost columns
            for (let i = 0; i < this.tableHeight; i++) {
                const row = worksheet.getRow(ref.row + i);
                for (let j = this.width; j < this._cache.width; j++) {
                    const cell = row.getCell(ref.col + j);
                    cell.value = null;
                }
            }
        }

        this.store();
    }

    /** Adds a row of data, inserted at `rowNumber` (0-based) or appended (needs `commit()`). */
    addRow(values: unknown[], rowNumber?: number): void {
        // Add a row of data, either insert at rowNumber or append
        this.cacheState();

        if (rowNumber === undefined) {
            rowsOf(this.table).push(values);
        } else {
            rowsOf(this.table).splice(rowNumber, 0, values);
        }
    }

    /** Removes `count` rows of data from `rowIndex` (0-based; needs `commit()`). */
    removeRows(rowIndex: number, count = 1): void {
        // Remove a rows of data
        this.cacheState();
        rowsOf(this.table).splice(rowIndex, count);
    }

    /** A column of the table (0-based index). */
    getColumn(colIndex: number): Column {
        // a missing column gives a Column without model, as it always did
        const column = this.table.columns[colIndex]!;
        return new Column(this, column, colIndex);
    }

    /** Adds a column with its values, inserted at `colIndex` (0-based) or appended (needs `commit()`). */
    addColumn(column: TableColumnModel, values: readonly unknown[], colIndex?: number): void {
        // Add a new column, including column defn and values
        // Inserts at colNumber or adds to the right
        this.cacheState();

        if (colIndex === undefined) {
            this.table.columns.push(column);
            rowsOf(this.table).forEach((row, i) => {
                row.push(values[i]);
            });
        } else {
            this.table.columns.splice(colIndex, 0, column);
            rowsOf(this.table).forEach((row, i) => {
                row.splice(colIndex, 0, values[i]);
            });
        }
    }

    /** Removes `count` columns from `colIndex` (0-based) with their data (needs `commit()`). */
    removeColumns(colIndex: number, count = 1): void {
        // Remove a column with data
        this.cacheState();

        this.table.columns.splice(colIndex, count);
        rowsOf(this.table).forEach((row) => {
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

    // the style is set by validate() (a loaded table has one too)

    /** The table style (`"TableStyleMedium2"`), or null for none. */
    get theme(): TableStyle["theme"] {
        return this.table.style!.theme;
    }
    set theme(value: TableStyle["theme"]) {
        this.table.style!.theme = value;
    }

    /** Whether the table style highlights the first column. */
    get showFirstColumn(): boolean | undefined {
        return this.table.style!.showFirstColumn;
    }
    set showFirstColumn(value: boolean | undefined) {
        this.table.style!.showFirstColumn = value;
    }

    /** Whether the table style highlights the last column. */
    get showLastColumn(): boolean | undefined {
        return this.table.style!.showLastColumn;
    }
    set showLastColumn(value: boolean | undefined) {
        this.table.style!.showLastColumn = value;
    }

    /** Whether the table style bands the rows. */
    get showRowStripes(): boolean | undefined {
        return this.table.style!.showRowStripes;
    }
    set showRowStripes(value: boolean | undefined) {
        this.table.style!.showRowStripes = value;
    }

    /** Whether the table style bands the columns. */
    get showColumnStripes(): boolean | undefined {
        return this.table.style!.showColumnStripes;
    }
    set showColumnStripes(value: boolean | undefined) {
        this.table.style!.showColumnStripes = value;
    }
}

/** The style of a table model. */
type TableStyle = NonNullable<TableModel["style"]>;

export default Table;
