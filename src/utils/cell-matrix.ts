import colCache, { type DecodedEx, type RangeBounds, type SheetCellAddress } from "./col-cache";
import { deepMerge } from "./under-dash";
import { createDictionary } from "./utils";
import { eachEntry, spliceEntries } from "./sparse-array";

/** A cell of a {@link CellMatrix}: where it is. Cells created from a location are that location object itself. */
export interface MatrixCell {
    sheetName?: string | undefined;
    /** The address without `$` signs, e.g. `"B3"`. */
    address: string;
    /** 1-based row number. */
    row: number;
    /** 1-based column number. */
    col: number;
    $col$row?: string;
}

/**
 * A location given to a {@link CellMatrix}: a cell, or what {@link colCache.decodeEx} returns. A range or an error
 * reference (#REF!) has no row and column: such a location is filed under the property "undefined" of the sparse
 * arrays, which their enumeration never visits (see indexOf below).
 */
export type MatrixLocation = MatrixCell | DecodedEx;

/** The sheet key of a location: its sheet name, or "undefined" for a location without one (as the JavaScript version did). */
const sheetKey = (location: MatrixLocation): string => String(location.sheetName);

/** A location that names no cell: a range (see {@link CellMatrix.addCellEx}) or an error reference (#REF!). */
type UnaddressedLocation = Exclude<MatrixLocation, MatrixCell | SheetCellAddress>;

/** A cell of a {@link CellMatrix}, with the template's properties. */
export type MatrixEntry<T> = MatrixCell & Partial<T>;

/**
 * What {@link CellMatrix.findCellEx} gives: a cell, or a location without row and column (a range, an error reference)
 * that was added as such.
 */
export type MatrixFound<T> = MatrixEntry<T> | (UnaddressedLocation & Partial<T>);

// Rows and cells are sparse arrays indexed by row and column number: they are traversed and spliced in time
// proportional to their entries (a single far cell, a defined name on $XFD$1048576, used to cost a million steps)
/**
 * A sparse set of cells per sheet name, e.g. the cells of a defined name.
 * @typeParam T - The properties that the template gives to each cell created from a location.
 */
class CellMatrix<T extends object = object> {
    /** A template for the cells created from a location: a deep copy of it is merged into each. */
    template: T | undefined;
    /** Sheet name → row number → column number → cell. No prototype: sheet names come from files. */
    sheets: Record<string, MatrixEntry<T>[][]>;
    /**
     * Sheet name → the location without row and column (a range, an error reference) added for that sheet: the first
     * one stays, it is never enumerated, and splicing the rows of the sheet drops it.
     */
    private readonly unaddressed = new Map<string, UnaddressedLocation & Partial<T>>();

    constructor(template?: T) {
        this.template = template;
        // null prototype: sheet names come from files (e.g. a defined name '__proto__'!$A$1)
        this.sheets = createDictionary<MatrixEntry<T>[][]>();
    }

    /**
     * Adds the cell or the cells of the range of a reference such as `Sheet1!$A$1:$B$2`.
     * @throws Error when a column or a row is out of bounds.
     */
    addCell(addressStr: string): void {
        this.addCellEx(colCache.decodeEx(addressStr));
    }

    /** The cell of a reference such as `Sheet1!$A$1`, created when missing. */
    getCell(addressStr: string): MatrixFound<T> {
        return this.findCellEx(colCache.decodeEx(addressStr), true);
    }

    /** The cell of a reference such as `Sheet1!$A$1`, or undefined. */
    findCell(addressStr: string): MatrixFound<T> | undefined {
        return this.findCellEx(colCache.decodeEx(addressStr), false);
    }

    /** The cell at a row and column of a sheet, or undefined. */
    findCellAt(sheetName: string, rowNumber: number, colNumber: number): MatrixEntry<T> | undefined {
        const sheet = this.sheets[sheetName];
        const row = sheet && sheet[rowNumber];
        return row && row[colNumber];
    }

    /**
     * Adds the cell of a location, or each cell of a range with a truthy top. A range starting at row 0 and an error
     * reference are added as they are (see {@link findCellEx}).
     */
    addCellEx(address: MatrixLocation): void {
        if ("top" in address && address.top) {
            for (let row = address.top; row <= address.bottom; row++) {
                for (let col = address.left; col <= address.right; col++) {
                    this.getCellAt(address.sheetName, row, col);
                }
            }
        } else {
            this.findCellEx(address, true);
        }
    }

    /** The cell of a location, created when missing. */
    getCellEx(address: MatrixLocation): MatrixFound<T> {
        return this.findCellEx(address, true);
    }

    /**
     * The cell of a location; when it is missing and create is true, the location itself becomes the cell (merged with
     * a copy of the template). A location without row and column (a range, an error reference) is found by its sheet
     * name only: one such location is kept per sheet.
     */
    findCellEx(address: MatrixLocation, create: true): MatrixFound<T>;
    findCellEx(address: MatrixLocation, create?: boolean): MatrixFound<T> | undefined;
    findCellEx(address: MatrixLocation, create?: boolean): MatrixFound<T> | undefined {
        if (!("row" in address)) {
            return this.findUnaddressed(address, create);
        }
        const sheet = this.findSheet(address, create);
        const row = sheet && this.findSheetRow(sheet, address, create);
        return row && this.findRowCell(row, address, create);
    }

    /**
     * The cell at a row and column of a sheet, created (without the template) when missing. A range without a sheet
     * name has its cells under the sheet "undefined" (and their sheetName undefined).
     */
    getCellAt(sheetName: string | undefined, rowNumber: number, colNumber: number): MatrixEntry<T> {
        const key = String(sheetName);
        const sheet = this.sheets[key] || (this.sheets[key] = []);
        const row = sheet[rowNumber] || (sheet[rowNumber] = []);
        // the template's properties are optional: these cells have none of them
        const noTemplate: Partial<T> = {};
        return (
            row[colNumber] ||
            (row[colNumber] = {
                ...noTemplate,
                sheetName,
                address: colCache.n2l(colNumber) + rowNumber,
                row: rowNumber,
                col: colNumber,
            })
        );
    }

    /** Removes the cell of a location. */
    removeCellEx(address: MatrixLocation): void {
        if (!("row" in address)) {
            this.unaddressed.delete(sheetKey(address));
            return;
        }
        const sheet = this.sheets[sheetKey(address)];
        const row = sheet && sheet[address.row];
        if (row) {
            // a hole, as in all the sparse arrays (Object.keys must not list it)
            // eslint-disable-next-line @typescript-eslint/no-array-delete -- deliberate: no splice, the indexes are column numbers
            delete row[address.col];
        }
    }

    /**
     * Removes the cells of a rectangle of a sheet (a range without a sheet name: the cells of the sheet "undefined"), in
     * time proportional to the cells of the sheet, whatever the size of the rectangle.
     */
    removeCellsIn(range: RangeBounds & { readonly sheetName?: string | undefined }): void {
        const sheet = this.sheets[String(range.sheetName)];
        if (!sheet) {
            return;
        }
        eachEntry(sheet, (row, rowNumber) => {
            if (rowNumber >= range.top && rowNumber <= range.bottom) {
                const removed: number[] = [];
                eachEntry(row, (_cell, colNumber) => {
                    if (colNumber >= range.left && colNumber <= range.right) {
                        removed.push(colNumber);
                    }
                });
                for (const colNumber of removed) {
                    // a hole, as in removeCellEx
                    // eslint-disable-next-line @typescript-eslint/no-array-delete -- deliberate: no splice, the indexes are column numbers
                    delete row[colNumber];
                }
            }
        });
    }

    /** Calls callback for each cell of a sheet, by row then column. */
    forEachInSheet(sheetName: string, callback: (cell: MatrixEntry<T>, row: number, col: number) => void): void {
        const sheet = this.sheets[sheetName];
        if (sheet) {
            eachEntry(sheet, (row, rowNumber) => {
                if (row) {
                    eachEntry(row, (cell, colNumber) => {
                        if (cell) {
                            callback(cell, rowNumber, colNumber);
                        }
                    });
                }
            });
        }
    }

    /** Calls callback for each cell, sheet by sheet (in Object.keys order), then by row and column. */
    forEach(callback: (cell: MatrixEntry<T>, row: number, col: number) => void): void {
        for (const sheetName of Object.keys(this.sheets)) {
            this.forEachInSheet(sheetName, callback);
        }
    }

    /** The results of callback for each cell, in the order of {@link forEach}. */
    map<R>(callback: (cell: MatrixEntry<T>) => R): R[] {
        const results: R[] = [];
        this.forEach((cell) => {
            results.push(callback(cell));
        });
        return results;
    }

    /** The rows of the location's sheet, created (empty) when missing and create is true. */
    findSheet(address: MatrixLocation, create?: boolean): MatrixEntry<T>[][] | undefined {
        const name = sheetKey(address);
        const sheet = this.sheets[name];
        if (sheet) {
            return sheet;
        }
        if (create) {
            return (this.sheets[name] = []);
        }
        return undefined;
    }

    /** The cells of the location's row in sheet, created (empty) when missing and create is true. */
    findSheetRow(sheet: MatrixEntry<T>[][], address: MatrixCell, create?: boolean): MatrixEntry<T>[] | undefined {
        const row = sheet[address.row];
        if (row) {
            return row;
        }
        if (create) {
            return (sheet[address.row] = []);
        }
        return undefined;
    }

    /**
     * The cell of the location's column in row; when missing and create is true, the location itself merged with a copy
     * of the template.
     */
    findRowCell(row: MatrixEntry<T>[], address: MatrixCell, create?: boolean): MatrixEntry<T> | undefined {
        const cell = row[address.col];
        if (cell) {
            return cell;
        }
        if (create) {
            return (row[address.col] = this.withTemplate(address));
        }
        return undefined;
    }

    /** The location without row and column kept for the location's sheet; when missing and create is true, the location. */
    private findUnaddressed(address: UnaddressedLocation, create?: boolean): (UnaddressedLocation & Partial<T>) | undefined {
        const key = sheetKey(address);
        const found = this.unaddressed.get(key);
        if (found) {
            return found;
        }
        if (create) {
            const created = this.withTemplate(address);
            this.unaddressed.set(key, created);
            return created;
        }
        return undefined;
    }

    /**
     * The location itself, with a deep copy of the template merged in (plain objects and arrays copied, undefined values
     * skipped: for a JSON template, what a JSON copy gives).
     */
    private withTemplate<L extends object>(location: L): L & Partial<T> {
        const copy: Partial<T> = this.template ? deepMerge({}, this.template) : {};
        return Object.assign(location, copy);
    }

    /** Moves the rows of a sheet as Array#splice(start, numDelete, numInsert empty rows) would. */
    spliceRows(sheetName: string, start: number, numDelete: number, numInsert: number): void {
        const sheet = this.sheets[sheetName];
        if (sheet) {
            this.sheets[sheetName] = spliceEntries(sheet, start, numDelete, numInsert);
        }
        // as the JavaScript version did: it kept such a location under the property "undefined" of the rows, which the
        // splice does not copy
        this.unaddressed.delete(sheetName);
    }

    /** Moves the cells of each row of a sheet as Array#splice(start, numDelete, numInsert empty cells) would. */
    spliceColumns(sheetName: string, start: number, numDelete: number, numInsert: number): void {
        const sheet = this.sheets[sheetName];
        if (sheet) {
            eachEntry(sheet, (row, rowNumber) => {
                if (row) {
                    sheet[rowNumber] = spliceEntries(row, start, numDelete, numInsert);
                }
            });
        }
    }
}

export default CellMatrix;
