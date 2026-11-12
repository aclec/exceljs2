// Type-level checks (no runtime content): the classes of src/doc against the public API of index.d.ts.
//
// Each check fails to compile when a class does not conform to its public type, except for the members listed in its
// Omit<>: the known mismatches, where index.d.ts does not describe the runtime (which wins) or where a member involves
// another class that does not conform (Cell, Worksheet...). Fixing index.d.ts should shrink these lists.

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

/** Compiles only when `T` is assignable to `U`. */
type Conforms<T extends U, U> = T;

export type AnchorConforms = Conforms<Anchor, Omit<Public.Anchor, "worksheet">>;
export type CellConforms = Conforms<
    Cell,
    Omit<
        Public.Cell,
        | "worksheet"
        | "workbook"
        | "master"
        | "hyperlink"
        | "model"
        | "name"
        | "dataValidation"
        | "value"
        | "note"
        | "formula"
        | "result"
        | "formulaType"
        | "toCsvString"
        | "release"
        | "addMergeRef"
        | "releaseMergeRef"
        | "merge"
        | "isMergedTo"
        | "numFmt"
        | "font"
        | "alignment"
        | "protection"
        | "border"
        | "fill"
        | "sheetName"
    >
>;
export type RowConforms = Conforms<
    Row,
    Omit<
        Public.Row,
        | "worksheet"
        | "model"
        | "height"
        | "values"
        | "getCell"
        | "findCell"
        | "getCellEx"
        | "eachCell"
        | "numFmt"
        | "font"
        | "alignment"
        | "protection"
        | "border"
        | "fill"
    >
>;
export type ColumnConforms = Conforms<Column, Omit<Public.Column, "values" | "worksheet" | "headers" | "eachCell">>;
export type RangeConforms = Conforms<Range, Omit<Public.Range, "sheetName" | "decode">>;
export type DataValidationsConforms = Conforms<DataValidations, Omit<Public.DataValidations, "model" | "add" | "find">>;
export type DefinedNamesConforms = Conforms<
    DefinedNames,
    Omit<Public.DefinedNames, "getMatrix" | "addEx" | "removeEx" | "removeAllNames" | "forEach" | "getNamesEx" | "getRanges">
>;
export type TableConforms = Conforms<
    Table,
    Omit<Public.Table, "getColumn" | "ref" | "headerRow" | "totalsRow" | "style" | "columns" | "rows">
>;
export type WorksheetConforms = Conforms<
    Worksheet,
    Omit<
        Public.Worksheet,
        | "workbook"
        // the views are the flat SheetViewModel (any field of the normal / frozen / split variants): a parsed view is not
        // narrowed to the variant of its `state`
        | "views"
        // a parsed worksheet without <pageMargins> has `pageSetup.margins === null`
        | "pageSetup"
        | "dataValidations"
        | "dimensions"
        | "headerFooter"
        | "properties"
        | "autoFilter"
        | "lastColumn"
        | "getColumnKey"
        | "setColumnKey"
        | "eachColumnKey"
        | "getColumn"
        | "columns"
        | "lastRow"
        | "findRow"
        | "findRows"
        | "addRow"
        | "addRows"
        | "insertRow"
        | "insertRows"
        | "getRow"
        | "getRows"
        | "eachRow"
        | "getSheetValues"
        | "findCell"
        | "getCell"
        | "getBackgroundImageId"
        | "getImages"
        | "commit"
        | "model"
        | "addTable"
        | "getTable"
        | "getTables"
    >
>;
export type WorkbookConforms = Conforms<
    Workbook,
    Omit<
        Public.Workbook,
        | "category"
        | "company"
        | "creator"
        | "description"
        | "keywords"
        | "lastModifiedBy"
        | "created"
        | "manager"
        | "modified"
        | "lastPrinted"
        | "properties"
        | "subject"
        | "title"
        | "calcProperties"
        // their read methods resolve with this Workbook, whose document properties may be undefined (see above)
        | "xlsx"
        | "csv"
        | "definedNames"
        | "model"
        | "views"
        | "worksheets"
        | "addWorksheet"
        | "removeWorksheetEx"
        | "getWorksheet"
        | "eachSheet"
        | "getImage"
    >
>;
