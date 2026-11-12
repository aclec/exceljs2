// The defined names of a workbook (`workbook.definedNames`): named cells and ranges, e.g. `MyRange` = `Sheet1!$A$1:$B$2`.

import colCache, { type RangeBounds } from "../utils/col-cache";
import CellMatrix, { type MatrixCell, type MatrixLocation } from "../utils/cell-matrix";
import Range from "./range";
import type { DefinedNameModel } from "../types";

// the references of the cells of a name (`$A$1`, `$A$1:$B$2`): only these ranges of a model are loaded
const rangeRegexp = /[$](\w+)[$](\d+)(:[$](\w+)[$](\d+))?/;

// Ranges are stored one entry per cell (so that names follow row/column splices).
// Ranges larger than this are kept as a single rectangle instead, so that a tiny
// definedName such as Sheet!$A$1:$XFD$1048576 cannot exhaust memory. When names are
// loaded (model setter), it also caps the cells expanded for all the names together:
// a cap per range would be multiplied by the number of names.
const MAX_EXPANDED_CELLS = 100000;

/** A cell of a name: marked while its ranges are rebuilt (see getRanges). */
type NameCell = MatrixCell & { mark?: boolean };

/** A range of a name too large to be expanded cell by cell (see MAX_EXPANDED_CELLS). */
export interface LargeNameRange {
    sheetName: string | undefined;
    top: number;
    left: number;
    bottom: number;
    right: number;
}

/**
 * A location given to {@link DefinedNames.addEx}: a cell, a decoded reference (`colCache.decodeEx`), or a range given by
 * its bounds only (index.d.ts `Location & { sheetName }`).
 */
export type NameLocation = MatrixLocation | (RangeBounds & { sheetName?: string | undefined });

/** The cells still allowed to be expanded, shared by several calls of {@link DefinedNames.addEx}. */
export interface ExpansionBudget {
    cells: number;
}

/**
 * Shifts a `[lo, hi]` interval after `numDelete` lines were removed and `numInsert` inserted at `start`.
 * @returns the new interval, or undefined when the whole interval was deleted
 */
function spliceInterval(lo: number, hi: number, start: number, numDelete: number, numInsert: number): [number, number] | undefined {
    const delta = numInsert - numDelete;
    const deletedEnd = start + numDelete - 1;
    const shift = (n: number): number => (n < start ? n : n + delta);
    if (numDelete > 0) {
        if (lo >= start && hi <= deletedEnd) {
            return undefined;
        }
        const newLo = lo >= start && lo <= deletedEnd ? start + numInsert : shift(lo);
        const newHi = hi >= start && hi <= deletedEnd ? start - 1 : shift(hi);
        return [newLo, newHi];
    }
    return [shift(lo), shift(hi)];
}

/** Whether a large range contains a cell location (a range or an error reference has no row: never contained). */
function rangeContains(range: LargeNameRange, address: MatrixLocation): boolean {
    const { row, col } = address as Partial<MatrixCell>;
    return (
        range.sheetName === address.sheetName &&
        row !== undefined &&
        col !== undefined &&
        row >= range.top &&
        row <= range.bottom &&
        col >= range.left &&
        col <= range.right
    );
}

/**
 * The defined names of a workbook. The cells of each name are kept one by one in a {@link CellMatrix} (so that names
 * follow the row and column splices); a range of more than 100,000 cells is kept as a rectangle (see
 * {@link LargeNameRange}): removing any cell of it removes the whole range.
 */
class DefinedNames {
    /** Name → its cells. No prototype: names come from files. */
    matrixMap: Record<string, CellMatrix<{ mark?: boolean }>>;
    /** Name → its ranges too large to expand. No prototype: names come from files. */
    rangeMap: Record<string, LargeNameRange[]>;

    constructor() {
        this.matrixMap = Object.create(null) as Record<string, CellMatrix<{ mark?: boolean }>>;
        // name -> [{sheetName, top, left, bottom, right}] for ranges too large to expand
        this.rangeMap = Object.create(null) as Record<string, LargeNameRange[]>;
    }

    /** The cells of a name (created empty when the name has none). */
    getMatrix(name: string): CellMatrix<{ mark?: boolean }> {
        const matrix = this.matrixMap[name] || (this.matrixMap[name] = new CellMatrix());
        return matrix;
    }

    /**
     * Adds a cell or a range to a name.
     * @param locStr - `SheetName!$col$row` or `SheetName!$c1$r1:$c2$r2`
     * @throws Error when a column or a row is out of bounds
     */
    add(locStr: string, name: string): void {
        const location = colCache.decodeEx(locStr);
        this.addEx(location, name);
    }

    /**
     * Adds a cell or a range to a name: a range larger than the budget (and than 100,000 cells) is kept as a rectangle.
     * @param budget - the cells still allowed to be expanded, shared by several calls (decremented)
     */
    addEx(location: NameLocation, name: string, budget: ExpansionBudget = { cells: MAX_EXPANDED_CELLS }): void {
        // a range (a range starting at row 0 is added as a cell, as it always was)
        const range = "top" in location && location.top ? location : undefined;
        const count = range ? (1 + range.bottom - range.top) * (1 + range.right - range.left) : 1;
        if (range && count > Math.min(budget.cells, MAX_EXPANDED_CELLS)) {
            const { sheetName, top, left, bottom, right } = range;
            (this.rangeMap[name] || (this.rangeMap[name] = [])).push({ sheetName, top, left, bottom, right });
            return;
        }
        budget.cells -= count;
        const matrix = this.getMatrix(name);
        if (range) {
            for (let col = range.left; col <= range.right; col++) {
                for (let row = range.top; row <= range.bottom; row++) {
                    const address = {
                        sheetName: range.sheetName,
                        address: colCache.n2l(col) + row,
                        row,
                        col,
                    };

                    matrix.addCellEx(address);
                }
            }
        } else {
            // a cell (a range starting at row 0, an invalid location, is filed as CellMatrix files it, as it always was)
            matrix.addCellEx(location as MatrixLocation);
        }
    }

    /** Removes a cell from a name (`SheetName!$col$row`). */
    remove(locStr: string, name: string): void {
        const location = colCache.decodeEx(locStr);
        this.removeEx(location, name);
    }

    /** Removes a cell from a name (and the large ranges of the name that contain it). */
    removeEx(location: MatrixLocation, name: string): void {
        const matrix = this.getMatrix(name);
        matrix.removeCellEx(location);
        this._removeFromRanges(location, name);
    }

    /** Removes a cell from all the names. */
    removeAllNames(location: MatrixLocation): void {
        for (const name of Object.keys(this.matrixMap)) {
            this.matrixMap[name]!.removeCellEx(location);
        }
        for (const name of Object.keys(this.rangeMap)) {
            this._removeFromRanges(location, name);
        }
    }

    /** Large ranges are atomic: removing any cell of one removes the whole range. */
    _removeFromRanges(location: MatrixLocation, name: string): void {
        const ranges = this.rangeMap[name];
        if (ranges) {
            this.rangeMap[name] = ranges.filter((range) => !this._rangeContains(range, location));
        }
    }

    /** Whether a large range contains a cell location. */
    _rangeContains(range: LargeNameRange, address: MatrixLocation): boolean {
        return rangeContains(range, address);
    }

    /**
     * Calls `callback` with each name and each of its cells. NB: the ranges kept as rectangles (too large to expand) are
     * not enumerated.
     */
    forEach(callback: (name: string, cell: NameCell) => void): void {
        for (const name of Object.keys(this.matrixMap)) {
            this.matrixMap[name]!.forEach((cell) => {
                callback(name, cell);
            });
        }
    }

    /** All the names of a cell (`Sheet1!$A$1`). */
    getNames(addressStr: string): string[] {
        return this.getNamesEx(colCache.decodeEx(addressStr));
    }

    /** All the names of a cell location. */
    getNamesEx(address: MatrixLocation): string[] {
        const names = new Set<string>();
        for (const name of Object.keys(this.matrixMap)) {
            if (this.matrixMap[name]!.findCellEx(address) && name) {
                names.add(name);
            }
        }
        for (const name of Object.keys(this.rangeMap)) {
            if (!names.has(name) && this.rangeMap[name]!.some((range) => this._rangeContains(range, address))) {
                names.add(name);
            }
        }
        return [...names];
    }

    /** The largest rectangle of marked cells around a marked cell, unmarking them. */
    _explore(matrix: CellMatrix<{ mark?: boolean }>, cell: NameCell): Range {
        cell.mark = false;
        // a cell without sheet name is filed under "undefined" (see CellMatrix)
        const sheetName = cell.sheetName as string;

        const range = new Range(cell.row, cell.col, cell.row, cell.col, sheetName);
        let x;
        let y;

        // grow vertical - only one col to worry about
        function vGrow(yy: number, edge: "top" | "bottom"): boolean {
            const c = matrix.findCellAt(sheetName, yy, cell.col);
            if (!c || !c.mark) {
                return false;
            }
            range[edge] = yy;
            c.mark = false;
            return true;
        }
        for (y = cell.row - 1; vGrow(y, "top"); y--);
        for (y = cell.row + 1; vGrow(y, "bottom"); y++);

        // grow horizontal - ensure all rows can grow
        function hGrow(xx: number, edge: "left" | "right"): boolean {
            const cells = [];
            for (y = range.top; y <= range.bottom; y++) {
                const c = matrix.findCellAt(sheetName, y, xx);
                if (c && c.mark) {
                    cells.push(c);
                } else {
                    return false;
                }
            }
            range[edge] = xx;
            for (let i = 0; i < cells.length; i++) {
                cells[i]!.mark = false;
            }
            return true;
        }
        for (x = cell.col - 1; hGrow(x, "left"); x--);
        for (x = cell.col + 1; hGrow(x, "right"); x++);

        return range;
    }

    /**
     * The ranges of a name (`$`-absolute references, `Sheet1!$A$1:$B$2`): its cells grouped into rectangles, then its
     * large ranges.
     * @param matrix - the cells of the name (default: its own)
     */
    getRanges(name: string, matrix?: CellMatrix<{ mark?: boolean }>): DefinedNameModel {
        matrix = matrix || this.matrixMap[name];
        const largeRanges = (this.rangeMap[name] || []).map(
            ({ sheetName, top, left, bottom, right }) => new Range(top, left, bottom, right, sheetName).$shortRange
        );

        if (!matrix) {
            return { name, ranges: largeRanges };
        }

        // mark and sweep!
        matrix.forEach((cell) => {
            cell.mark = true;
        });
        const cells = matrix;
        const ranges = matrix
            .map((cell) => cell.mark && this._explore(cells, cell))
            .filter((range): range is Range => !!range)
            .map((range) => range.$shortRange);

        return {
            name,
            ranges: ranges.concat(largeRanges),
        };
    }

    /** Updates the row, column and address of the cells of a sheet that a splice moved. */
    normaliseMatrix(matrix: CellMatrix<{ mark?: boolean }>, sheetName: string): void {
        // some of the cells might have shifted on specified sheet
        // need to reassign rows, cols
        matrix.forEachInSheet(sheetName, (cell, row, col) => {
            if (cell) {
                if (cell.row !== row || cell.col !== col) {
                    cell.row = row;
                    cell.col = col;
                    cell.address = colCache.n2l(col) + row;
                }
            }
        });
    }

    /** Moves the names after `numDelete` rows were removed and `numInsert` inserted at row `start` of a sheet. */
    spliceRows(sheetName: string, start: number, numDelete: number, numInsert: number): void {
        for (const name of Object.keys(this.matrixMap)) {
            const matrix = this.matrixMap[name]!;
            matrix.spliceRows(sheetName, start, numDelete, numInsert);
            this.normaliseMatrix(matrix, sheetName);
        }
        this._spliceRanges(sheetName, start, numDelete, numInsert, "top", "bottom");
    }

    /** Moves the names after `numDelete` columns were removed and `numInsert` inserted at column `start` of a sheet. */
    spliceColumns(sheetName: string, start: number, numDelete: number, numInsert: number): void {
        for (const name of Object.keys(this.matrixMap)) {
            const matrix = this.matrixMap[name]!;
            matrix.spliceColumns(sheetName, start, numDelete, numInsert);
            this.normaliseMatrix(matrix, sheetName);
        }
        this._spliceRanges(sheetName, start, numDelete, numInsert, "left", "right");
    }

    /** Moves the large ranges of a sheet (a range whose lines were all deleted is removed). */
    _spliceRanges(
        sheetName: string,
        start: number,
        numDelete: number,
        numInsert: number,
        lo: "top" | "left",
        hi: "bottom" | "right"
    ): void {
        for (const name of Object.keys(this.rangeMap)) {
            this.rangeMap[name] = this.rangeMap[name]!.map((range) => {
                if (range.sheetName !== sheetName) {
                    return range;
                }
                const interval = spliceInterval(range[lo], range[hi], start, numDelete, numInsert);
                return interval && { ...range, [lo]: interval[0], [hi]: interval[1] };
            }).filter((range): range is LargeNameRange => !!range);
        }
    }

    /** The names with their ranges (see getRanges); names without range are left out. */
    get model(): DefinedNameModel[] {
        // To get names per cell - just iterate over all names finding cells if they exist
        const names = new Set([...Object.keys(this.matrixMap), ...Object.keys(this.rangeMap)]);
        return [...names].map((name) => this.getRanges(name)).filter((definedName) => definedName.ranges.length);
    }

    /**
     * Replaces the names. Only the references with `$` (`Sheet1!$A$1:$B$2`) are loaded; the cells expanded for all the
     * names together are capped (see MAX_EXPANDED_CELLS).
     * @throws Error when a column or a row is out of bounds
     */
    set model(value: readonly DefinedNameModel[]) {
        // value is [ { name, ranges }, ... ]
        this.matrixMap = Object.create(null) as Record<string, CellMatrix<{ mark?: boolean }>>;
        this.rangeMap = Object.create(null) as Record<string, LargeNameRange[]>;
        const budget = { cells: MAX_EXPANDED_CELLS };
        value.forEach((definedName) => {
            this.getMatrix(definedName.name);
            definedName.ranges.forEach((rangeStr) => {
                if (rangeRegexp.test(rangeStr.split("!").pop() || "")) {
                    this.addEx(colCache.decodeEx(rangeStr), definedName.name, budget);
                }
            });
        });
    }
}

export default DefinedNames;
