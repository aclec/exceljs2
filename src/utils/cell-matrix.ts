/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import _ from "./under-dash";
import colCache from "./col-cache";
import sparseArrayModule from "./sparse-array";
const { eachEntry, spliceEntries } = sparseArrayModule;

// Rows and cells are sparse arrays indexed by row and column number: they are traversed and spliced in time
// proportional to their entries (a single far cell, a defined name on $XFD$1048576, used to cost a million steps)
class CellMatrix {
    constructor(template) {
        this.template = template;
        // null prototype: sheet names come from files (e.g. a defined name '__proto__'!$A$1)
        this.sheets = Object.create(null);
    }

    addCell(addressStr) {
        this.addCellEx(colCache.decodeEx(addressStr));
    }

    getCell(addressStr) {
        return this.findCellEx(colCache.decodeEx(addressStr), true);
    }

    findCell(addressStr) {
        return this.findCellEx(colCache.decodeEx(addressStr), false);
    }

    findCellAt(sheetName, rowNumber, colNumber) {
        const sheet = this.sheets[sheetName];
        const row = sheet && sheet[rowNumber];
        return row && row[colNumber];
    }

    addCellEx(address) {
        if (address.top) {
            for (let row = address.top; row <= address.bottom; row++) {
                for (let col = address.left; col <= address.right; col++) {
                    this.getCellAt(address.sheetName, row, col);
                }
            }
        } else {
            this.findCellEx(address, true);
        }
    }

    getCellEx(address) {
        return this.findCellEx(address, true);
    }

    findCellEx(address, create) {
        const sheet = this.findSheet(address, create);
        const row = this.findSheetRow(sheet, address, create);
        return this.findRowCell(row, address, create);
    }

    getCellAt(sheetName, rowNumber, colNumber) {
        const sheet = this.sheets[sheetName] || (this.sheets[sheetName] = []);
        const row = sheet[rowNumber] || (sheet[rowNumber] = []);
        const cell =
            row[colNumber] ||
            (row[colNumber] = {
                sheetName,
                address: colCache.n2l(colNumber) + rowNumber,
                row: rowNumber,
                col: colNumber,
            });
        return cell;
    }

    removeCellEx(address) {
        const sheet = this.findSheet(address);
        if (!sheet) {
            return;
        }
        const row = this.findSheetRow(sheet, address);
        if (!row) {
            return;
        }
        delete row[address.col];
    }

    forEachInSheet(sheetName, callback) {
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

    forEach(callback) {
        _.each(this.sheets, (sheet, sheetName) => {
            this.forEachInSheet(sheetName, callback);
        });
    }

    map(callback) {
        const results = [];
        this.forEach((cell) => {
            results.push(callback(cell));
        });
        return results;
    }

    findSheet(address, create) {
        const name = address.sheetName;
        if (this.sheets[name]) {
            return this.sheets[name];
        }
        if (create) {
            return (this.sheets[name] = []);
        }
        return undefined;
    }

    findSheetRow(sheet, address, create) {
        const { row } = address;
        if (sheet && sheet[row]) {
            return sheet[row];
        }
        if (create) {
            return (sheet[row] = []);
        }
        return undefined;
    }

    findRowCell(row, address, create) {
        const { col } = address;
        if (row && row[col]) {
            return row[col];
        }
        if (create) {
            return (row[col] = this.template ? Object.assign(address, JSON.parse(JSON.stringify(this.template))) : address);
        }
        return undefined;
    }

    spliceRows(sheetName, start, numDelete, numInsert) {
        const sheet = this.sheets[sheetName];
        if (sheet) {
            this.sheets[sheetName] = spliceEntries(sheet, start, numDelete, numInsert);
        }
    }

    spliceColumns(sheetName, start, numDelete, numInsert) {
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
