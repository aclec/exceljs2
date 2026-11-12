// Types of the I/O orchestration (src/xlsx/xlsx.ts, src/csv, src/stream): the streaming reader / writer and the
// views they take of the document classes.
//
// TS-WAVE2: src/doc/row.ts, column.ts, range.ts and cell.ts are typed in parallel (wave 2). Until then, the stream
// reader / writer see them through the views below, built from the public types of index.d.ts plus the internal
// members they use; each importer asserts the constructor once (`as unknown as Doc...Constructor`). Once those files are
// typed, the views and the assertions go (the classes are used directly).

import type { Cell, CellValue, Column, Range, Row, Style } from "../../index";
import type { CellAddress } from "../utils/col-cache";
import type { SparseArray } from "../utils/sparse-array";
import type { ColModel } from "./cell";

/**
 * The value of a cell read by the streaming reader when the shared strings are not cached (`sharedStrings: "emit"`
 * or `"ignore"`): the index of the string in the table. Not in index.d.ts `CellValue`.
 */
export interface SharedStringRefValue {
    sharedString: number;
}

/** A cell as the stream reader / writer use it: index.d.ts `Cell` plus the internal note. */
export type DocCell = Cell & {
    /** The note of the cell (a Note), if any. */
    readonly _comment?: unknown;
};

/** A row as the stream reader / writer use it: index.d.ts `Row` plus its internals. */
export type DocRow = Omit<Row, "getCell" | "getCellEx" | "findCell"> & {
    /** The style of the row (applied to its new cells). */
    style: Partial<Style>;
    /** The cells, indexed by column number - 1 (sparse). */
    readonly _cells: SparseArray<DocCell>;
    getCell(indexOrKey: number | string): DocCell;
    getCellEx(address: CellAddress): DocCell;
    findCell(colNumber: number): DocCell | undefined;
};

/** The constructor of src/doc/row.ts. */
export type DocRowConstructor = new (worksheet: object, number: number) => DocRow;

/** The `<col>` ranges of a worksheet, as Column.rangesFromModel returns them: sorted, disjoint, clamped to the grid. */
export interface ColumnRanges {
    ranges: { min: number; max: number; defn: ColModel }[];
    /** The last column covered (0 when there is none). */
    length: number;
}

/** The constructor and statics of src/doc/column.ts. */
export interface DocColumnConstructor {
    /** @param defn - the definition (`false`: none, set later) */
    new (worksheet: object, number: number, defn?: unknown): Column;
    /** The `<col>` ranges of columns (equivalent neighbours merged); `undefined` when there is none. */
    toModel(columns: readonly Column[] | null): ColModel[] | undefined;
    rangesFromModel(cols: readonly ColModel[] | null): ColumnRanges;
    /** The definition of column `n` (1-based) in the ranges, if any. */
    defnInRanges(colRanges: ColumnRanges | null, n: number): ColModel | undefined;
}

/** A Range (src/doc/range.ts) as the stream reader / writer use it. */
export type DocRange = Range & {
    /** Extends the range to the cells of a row. */
    expandRow(row: DocRow): void;
};

/** The constructor of src/doc/range.ts: no argument (an empty range) or the arguments of `Range.decode`. */
export type DocRangeConstructor = new (...args: unknown[]) => DocRange;

/** A value assigned to a cell by the stream reader: a cell value, or a shared string index. */
export type ReadCellValue = CellValue | SharedStringRefValue;
