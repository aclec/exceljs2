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
import type ModelContainer from "../doc/modelcontainer";

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

/**
 * The members of `P` that the runtime class `I` lacks (`missing:name`) or does not conform to (their type, or an
 * optional member where `P` declares a required one).
 */
type Mismatches<I, P> = {
    [K in Exclude<keyof P, BackReference>]-?: K extends keyof I
        ? [Erase<Pick<I, K>>] extends [Erase<Pick<P, K>>]
            ? never
            : K
        : `missing:${K & string}`;
}[Exclude<keyof P, BackReference>];

/**
 * `true` when `I` conforms to `P` but for the members `Except`, and every member of `Except` still mismatches; else
 * the names of the other mismatches, and `stale:name` for an exception that conforms now (remove it from the list).
 */
type Check<I, P, Except extends keyof P = never> = [Mismatches<I, Omit<P, Except>> | Stale<I, P, Except>] extends [never]
    ? true
    : Mismatches<I, Omit<P, Except>> | Stale<I, P, Except>;

/** The members of `Except` that conform (`stale:name`). */
type Stale<I, P, Except extends keyof P> = `stale:${Exclude<Except, Mismatches<I, P>> & string}`;

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
        // `value` returns two kinds of values that index.d.ts `CellValue` keeps out on purpose (they come only
        // from invalid input; describing them would make every cell value an untyped `object`):
        // - a JSON value: any object that is not a cell value (only from a value outside `CellValue`, such as an
        //   untyped one, or from the falsy-marker quirk: `{ text: "", hyperlink }`, `{ sharedString: 0 }`...);
        // - a formula read from a malformed file with neither `formula` nor `sharedFormula` (a shared formula clone
        //   whose `si` has no master): its value is `{ result }`.
        // `merge` / `isMergedTo` take a Cell: they mismatch only through the `value` of that Cell.
        "value" | "merge" | "isMergedTo"
    >
>;
/** The runtime Row conforms to index.d.ts `Row`, but for the known mismatches listed. */
export type RowConforms = Assert<
    Check<
        Row,
        Public.Row,
        // Only through `Cell.value` (see CellConforms): `values` holds cell values, the others give Cells.
        "values" | "findCell" | "getCellEx" | "getCell"
    >
>;
/** The runtime Column conforms to index.d.ts `Column`, but for the known mismatches listed. */
export type ColumnConforms = Assert<
    Check<
        Column,
        Public.Column,
        // Only through `Cell.value` (see CellConforms): `values` holds cell values.
        "values"
    >
>;
/** The runtime DefinedNames conforms to index.d.ts `DefinedNames`. */
export type DefinedNamesConforms = Assert<Check<DefinedNames, Public.DefinedNames>>;
/** The runtime Table conforms to index.d.ts `Table`. */
export type TableConforms = Assert<Check<Table, Public.Table>>;
/** The runtime Worksheet conforms to index.d.ts `Worksheet`, but for the known mismatches listed. */
export type WorksheetConforms = Assert<
    Check<
        Worksheet,
        Public.Worksheet,
        // undefined only for an invalid column number (not an integer >= 1), typed `Column` as documented; it
        // also mismatches through `Cell.value`, like the next ones
        | "getColumn"
        // only through `Cell.value` (see CellConforms): they give Cells, or Rows / Columns whose cells hold values
        | "findCell"
        | "getCell"
        | "lastColumn"
        | "lastRow"
    >
>;
/** The runtime Workbook conforms to index.d.ts `Workbook`. */
export type WorkbookConforms = Assert<Check<Workbook, Public.Workbook>>;
/** The runtime ModelContainer conforms to index.d.ts `ModelContainer`, but for the known mismatches listed. */
export type ModelContainerConforms = Assert<
    Check<
        ModelContainer,
        Public.ModelContainer,
        // `model` holds the model given to the constructor or parsed by `xlsx`: a parsed worksheet model lacks what only
        // the Worksheet getter sets (`state`, `rowBreaks`, `merges`, `dimensions` as a Range)
        "model"
    >
>;
