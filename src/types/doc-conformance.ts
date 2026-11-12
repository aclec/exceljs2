// Type-level checks (no runtime content): the classes of src/doc against the public API of index.d.ts.
//
// Each check compiles only when every public member of the index.d.ts type, except the ones listed, conforms to the
// runtime class (its type is assignable to the declared one). On failure the error names the members that do not
// conform. The comparison goes through {@link Erase}:
// - the back-references `worksheet`, `workbook` and `master` are not compared: a row or a cell of the streaming API
//   belongs to a stream worksheet, which index.d.ts does not model as a Worksheet; the classes type them as the narrow
//   interfaces they use (`CellWorksheet`, `RowWorksheet`...);
// - members whose name starts with `_` (internal by convention) are dropped;
// - methods are compared as TypeScript compares methods (bivariant parameters), to a depth of 3 levels.
//
// The lists of exceptions are the remaining known mismatches. Every entry must say why it stays; aligning index.d.ts
// with the runtime removes entries.

import type * as Public from "../../index";
import type Anchor from "../doc/anchor";
import type Cell from "../doc/cell";
import type Row from "../doc/row";
import type Column from "../doc/column";
import type Range from "../doc/range";
import type DataValidations from "../doc/data-validations";
import type DefinedNames from "../doc/defined-names";
import type Table from "../doc/table";
import type Worksheet from "../doc/worksheet";
import type Workbook from "../doc/workbook";

/** The back-references that are not compared (see the file comment). */
type BackReference = "worksheet" | "workbook" | "master";

/** Depth counter of {@link Erase}. */
type Previous = [never, 0, 1, 2, 3];

/** A function type compared as a method (bivariant parameters). */
type Method<A extends readonly unknown[], R> = { method(...args: A): R }["method"];

/**
 * The structure of a type as the checks compare it: without back-references nor `_` members, recursively to depth
 * `D` (`unknown` beyond).
 */
type Erase<T, D extends number = 3> = [D] extends [never]
    ? unknown
    : T extends (...args: infer A) => infer R
      ? Method<{ [I in keyof A]: Erase<A[I], Previous[D]> }, Erase<R, Previous[D]>>
      : T extends Date | string | number | boolean | null | undefined
        ? T
        : T extends readonly (infer E)[]
          ? Erase<E, Previous[D]>[]
          : T extends object
            ? { [K in keyof T as K extends BackReference | `_${string}` ? never : K]: Erase<T[K], Previous[D]> }
            : T;

/** The members of `P` that the runtime class `I` lacks (`missing:name`) or does not conform to. */
type Mismatches<I, P> = {
    [K in Exclude<keyof P, BackReference>]-?: K extends keyof I
        ? [Erase<I[K]>] extends [Erase<P[K]>]
            ? never
            : K
        : `missing:${K & string}`;
}[Exclude<keyof P, BackReference>];

/** `true` when `I` conforms to `P` but for the members `Except`; else the names of the other mismatches. */
type Check<I, P, Except extends keyof P = never> = [Mismatches<I, Omit<P, Except>>] extends [never] ? true : Mismatches<I, Omit<P, Except>>;

/** Compiles only for `true`: a failing check names its mismatches in the error. */
type Assert<T extends true> = T;

/** The runtime Anchor conforms to index.d.ts `Anchor`. */
export type AnchorConforms = Assert<Check<Anchor, Public.Anchor>>;
/** The runtime Range conforms to index.d.ts `Range`. */
export type RangeConforms = Assert<Check<Range, Public.Range>>;
/** The runtime DataValidations conforms to index.d.ts `DataValidations`. */
export type DataValidationsConforms = Assert<Check<DataValidations, Public.DataValidations>>;

/** The runtime Cell conforms to index.d.ts `Cell`, but for the known mismatches listed. */
export type CellConforms = Assert<
    Check<
        Cell,
        Public.Cell,
        // TODO(align)
        "value" | "merge" | "isMergedTo"
    >
>;
/** The runtime Row conforms to index.d.ts `Row`, but for the known mismatches listed. */
export type RowConforms = Assert<
    Check<
        Row,
        Public.Row,
        // TODO(align)
        "values" | "findCell" | "getCellEx" | "getCell" | "eachCell"
    >
>;
/** The runtime Column conforms to index.d.ts `Column`, but for the known mismatches listed. */
export type ColumnConforms = Assert<
    Check<
        Column,
        Public.Column,
        // TODO(align)
        "values" | "eachCell"
    >
>;
/** The runtime DefinedNames conforms to index.d.ts `DefinedNames`, but for the known mismatches listed. */
export type DefinedNamesConforms = Assert<
    Check<
        DefinedNames,
        Public.DefinedNames,
        // TODO(align)
        "addEx" | "removeEx"
    >
>;
/** The runtime Table conforms to index.d.ts `Table`, but for the known mismatches listed. */
export type TableConforms = Assert<
    Check<
        Table,
        Public.Table,
        // TODO(align)
        "getColumn"
    >
>;
/** The runtime Worksheet conforms to index.d.ts `Worksheet`, but for the known mismatches listed. */
export type WorksheetConforms = Assert<
    Check<
        Worksheet,
        Public.Worksheet,
        // TODO(align) worksheet settings
        | "properties"
        | "views"
        | "model"
        | "pageSetup"
        // rows, columns and cells: undefined only for an invalid column number (not an integer >= 1), typed `Column`
        // as documented (and also the Column cascade below)
        | "getColumn"
        // TODO(align) rows, columns and cells: their signatures are aligned; they fail only through the Row / Cell /
        // Column / CellValue types they return or take (removed once Row, Cell, Column and CellValue conform)
        | "findCell"
        | "getCell"
        | "columns"
        | "insertRows"
        | "lastColumn"
        | "getColumnKey"
        | "setColumnKey"
        | "lastRow"
        | "findRow"
        | "findRows"
        | "addRow"
        | "addRows"
        | "insertRow"
        | "getRow"
        | "getRows"
        | "eachRow"
        | "getSheetValues"
    >
>;
/** The runtime Workbook conforms to index.d.ts `Workbook`, but for the known mismatches listed. */
export type WorkbookConforms = Assert<
    Check<
        Workbook,
        Public.Workbook,
        // TODO(align)
        | "properties"
        | "calcProperties"
        | "worksheets"
        | "definedNames"
        | "views"
        | "company"
        | "manager"
        | "model"
        | "creator"
        | "title"
        | "subject"
        | "description"
        | "keywords"
        | "category"
        | "lastModifiedBy"
        | "lastPrinted"
        | "created"
        | "modified"
        | "xlsx"
        | "csv"
        | "addWorksheet"
        | "removeWorksheetEx"
        | "getWorksheet"
        | "eachSheet"
        | "getImage"
    >
>;
