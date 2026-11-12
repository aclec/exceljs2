import colCache, { type DecodedEx } from "./col-cache";
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

// The sheets, rows and columns of locations are property keys of the sparse arrays (and of sheets): a location without
// a sheet name, a row or a column (a range, an error reference) is filed under "undefined", as the JavaScript version
// did. These views name the keys of such locations without changing them.
const sheetKey = (location: MatrixLocation): string => String(location.sheetName);
const rowIndex = (location: MatrixLocation): number => (location as MatrixCell).row;
const colIndex = (location: MatrixLocation): number => (location as MatrixCell).col;

// Rows and cells are sparse arrays indexed by row and column number: they are traversed and spliced in time
// proportional to their entries (a single far cell, a defined name on $XFD$1048576, used to cost a million steps)
/**
 * A sparse set of cells per sheet name, e.g. the cells of a defined name.
 * @typeParam T - The properties that the template gives to each cell created from a location.
 */
class CellMatrix<T extends object = object> {
    /** A template for the cells created from a location: a JSON copy of it is merged into each. */
    template: T | undefined;
    /** Sheet name → row number → column number → cell. No prototype: sheet names come from files. */
    sheets: Record<string, (MatrixCell & Partial<T>)[][]>;

    constructor(template?: T) {
        this.template = template;
        // null prototype: sheet names come from files (e.g. a defined name '__proto__'!$A$1)
        this.sheets = Object.create(null) as Record<string, (MatrixCell & Partial<T>)[][]>;
    }

    /**
     * Adds the cell or the cells of the range of a reference such as `Sheet1!$A$1:$B$2`.
     * @throws Error when a column or a row is out of bounds.
     */
    addCell(addressStr: string): void {
        this.addCellEx(colCache.decodeEx(addressStr));
    }

    /** The cell of a reference such as `Sheet1!$A$1`, created when missing. */
    getCell(addressStr: string): MatrixCell & Partial<T> {
        return this.findCellEx(colCache.decodeEx(addressStr), true);
    }

    /** The cell of a reference such as `Sheet1!$A$1`, or undefined. */
    findCell(addressStr: string): (MatrixCell & Partial<T>) | undefined {
        return this.findCellEx(colCache.decodeEx(addressStr), false);
    }

    /** The cell at a row and column of a sheet, or undefined. */
    findCellAt(sheetName: string, rowNumber: number, colNumber: number): (MatrixCell & Partial<T>) | undefined {
        const sheet = this.sheets[sheetName];
        const row = sheet && sheet[rowNumber];
        return row && row[colNumber];
    }

    /** Adds the cell of a location, or each cell of a range (with a truthy top). */
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
    getCellEx(address: MatrixLocation): MatrixCell & Partial<T> {
        return this.findCellEx(address, true);
    }

    /**
     * The cell of a location; when it is missing and create is true, the location itself becomes the cell (merged with
     * a copy of the template).
     */
    findCellEx(address: MatrixLocation, create: true): MatrixCell & Partial<T>;
    findCellEx(address: MatrixLocation, create?: boolean): (MatrixCell & Partial<T>) | undefined;
    findCellEx(address: MatrixLocation, create?: boolean): (MatrixCell & Partial<T>) | undefined {
        const sheet = this.findSheet(address, create);
        const row = this.findSheetRow(sheet, address, create);
        return this.findRowCell(row, address, create);
    }

    /**
     * The cell at a row and column of a sheet, created (without the template) when missing. A range without a sheet
     * name has its cells under the sheet "undefined" (and their sheetName undefined).
     */
    getCellAt(sheetName: string | undefined, rowNumber: number, colNumber: number): MatrixCell & Partial<T> {
        const key = String(sheetName);
        const sheet = this.sheets[key] || (this.sheets[key] = []);
        const row = sheet[rowNumber] || (sheet[rowNumber] = []);
        const cell =
            row[colNumber] ||
            (row[colNumber] = {
                sheetName,
                address: colCache.n2l(colNumber) + rowNumber,
                row: rowNumber,
                col: colNumber,
            } as MatrixCell & Partial<T>); // the template's properties are optional
        return cell;
    }

    /** Removes the cell of a location. */
    removeCellEx(address: MatrixLocation): void {
        const sheet = this.findSheet(address);
        if (!sheet) {
            return;
        }
        const row = this.findSheetRow(sheet, address);
        if (!row) {
            return;
        }
        // a hole, as in all the sparse arrays (Object.keys must not list it)
        // eslint-disable-next-line @typescript-eslint/no-array-delete -- deliberate: no splice, the indexes are column numbers
        delete row[colIndex(address)];
    }

    /** Calls callback for each cell of a sheet, by row then column. */
    forEachInSheet(sheetName: string, callback: (cell: MatrixCell & Partial<T>, row: number, col: number) => void): void {
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
    forEach(callback: (cell: MatrixCell & Partial<T>, row: number, col: number) => void): void {
        for (const sheetName of Object.keys(this.sheets)) {
            this.forEachInSheet(sheetName, callback);
        }
    }

    /** The results of callback for each cell, in the order of {@link forEach}. */
    map<R>(callback: (cell: MatrixCell & Partial<T>) => R): R[] {
        const results: R[] = [];
        this.forEach((cell) => {
            results.push(callback(cell));
        });
        return results;
    }

    /** The rows of the location's sheet, created (empty) when missing and create is true. */
    findSheet(address: MatrixLocation, create?: boolean): (MatrixCell & Partial<T>)[][] | undefined {
        const name = sheetKey(address);
        if (this.sheets[name]) {
            return this.sheets[name];
        }
        if (create) {
            return (this.sheets[name] = []);
        }
        return undefined;
    }

    /** The cells of the location's row in sheet, created (empty) when missing and create is true (sheet is then set). */
    findSheetRow(
        sheet: (MatrixCell & Partial<T>)[][] | undefined,
        address: MatrixLocation,
        create?: boolean
    ): (MatrixCell & Partial<T>)[] | undefined {
        const row = rowIndex(address);
        if (sheet && sheet[row]) {
            return sheet[row];
        }
        if (create) {
            return (sheet![row] = []);
        }
        return undefined;
    }

    /**
     * The cell of the location's column in row; when missing and create is true (row is then set), the location itself
     * merged with a copy of the template.
     */
    findRowCell(
        row: (MatrixCell & Partial<T>)[] | undefined,
        address: MatrixLocation,
        create?: boolean
    ): (MatrixCell & Partial<T>) | undefined {
        const col = colIndex(address);
        if (row && row[col]) {
            return row[col];
        }
        if (create) {
            // the location becomes the cell (a range or an error reference too, filed under "undefined")
            const cell = address as MatrixCell & Partial<T>;
            return (row![col] = this.template ? Object.assign(cell, JSON.parse(JSON.stringify(this.template)) as Partial<T>) : cell);
        }
        return undefined;
    }

    /** Moves the rows of a sheet as Array#splice(start, numDelete, numInsert empty rows) would. */
    spliceRows(sheetName: string, start: number, numDelete: number, numInsert: number): void {
        const sheet = this.sheets[sheetName];
        if (sheet) {
            this.sheets[sheetName] = spliceEntries(sheet, start, numDelete, numInsert);
        }
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
