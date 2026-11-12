// Pivot tables (experimental): `worksheet.addPivotTable(...)`. One pivot table per workbook, summing one value field.

import { objectFromProps, range, toSortedArray } from "../utils/utils";
import type { CellValue } from "./cell";
import type Worksheet from "./worksheet";

// TK(2023-10-10): turn this into a class constructor.

/**
 * What `worksheet.addPivotTable` takes: the source worksheet (its whole range: row 1 holds the field names), and the
 * names of the fields of the rows, columns and values.
 *
 * @example
 * {
 *   sourceSheet: worksheet1,
 *   rows: ['A', 'B'],
 *   columns: ['C'],
 *   values: ['E'], // only 1 item possible for now
 *   metric: 'sum', // only 'sum' possible for now
 * }
 */
export interface PivotTableInput {
    sourceSheet: Worksheet;
    rows: readonly string[];
    columns: readonly string[];
    values: readonly string[];
    metric?: string | undefined;
}

/**
 * A field of the pivot cache: a column of the source worksheet, its name (row 1), and for the fields of the rows and
 * columns, their distinct values (sorted).
 */
export interface PivotCacheField {
    name: CellValue | undefined;
    sharedItems: CellValue[] | null;
}

/**
 * A pivot table (`worksheet.pivotTables`, `workbook.pivotTables`): the fields of its rows, columns and values as indexes
 * into `cacheFields` (whose names are those of the source worksheet).
 */
export interface PivotTable {
    sourceSheet: Worksheet;
    rows: number[];
    columns: number[];
    values: number[];
    metric: "sum";
    cacheFields: PivotCacheField[];
    /** Defined in `<pivotTableDefinition>` of xl/pivotTables/pivotTable1.xml; also used in xl/workbook.xml. */
    cacheId: string;
    /** Set when the workbook is written (xlsx.ts): the relationship id of the pivot cache definition from the workbook. */
    rId?: string;
}

/**
 * Creates a pivot table.
 * @throws Error when the workbook has one already, for another metric than "sum", for a field name that is not in row 1
 *   of the source worksheet, without rows or columns, or without exactly one value
 */
function makePivotTable(worksheet: Worksheet, model: PivotTableInput): PivotTable {
    validate(worksheet, model);

    const { sourceSheet } = model;

    const cacheFields = makeCacheFields(sourceSheet, [...model.rows, ...model.columns]);

    // let {rows, columns, values} use indices instead of names;
    // names can then be accessed via `pivotTable.cacheFields[index].name`.
    // no prototype: the names come from the data (see objectFromProps)
    const nameToIndex = cacheFields.reduce<Record<string, number>>(
        (result, cacheField, index) => {
            // a name is a property key (coerced to a string)
            result[cacheField.name as string] = index;
            return result;
        },
        Object.create(null) as Record<string, number>
    );
    // validate() checked that the names are those of the cache fields
    const indexOf = (name: string): number => nameToIndex[name]!;

    // form pivot table object
    return {
        sourceSheet,
        rows: model.rows.map(indexOf),
        columns: model.columns.map(indexOf),
        values: model.values.map(indexOf),
        metric: "sum",
        cacheFields,
        // defined in <pivotTableDefinition> of xl/pivotTables/pivotTable1.xml;
        // also used in xl/workbook.xml
        cacheId: "10",
    };
}

/** @throws Error for an invalid pivot table (see makePivotTable) */
function validate(worksheet: Worksheet, model: PivotTableInput): void {
    if (worksheet.workbook.pivotTables.length === 1) {
        throw new Error("A pivot table was already added. At this time, ExcelJS supports at most one pivot table per file.");
    }

    if (model.metric && model.metric !== "sum") {
        throw new Error('Only the "sum" metric is supported at this time.');
    }

    // the header names are property keys (coerced to strings)
    const headerNames = model.sourceSheet.getRow(1).values.slice(1) as string[];
    const isInHeaderNames = objectFromProps(headerNames, true);
    for (const name of [...model.rows, ...model.columns, ...model.values]) {
        if (!isInHeaderNames[name]) {
            throw new Error(`The header name "${name}" was not found in ${model.sourceSheet.name}.`);
        }
    }

    if (!model.rows.length) {
        throw new Error("No pivot table rows specified.");
    }

    if (!model.columns.length) {
        throw new Error("No pivot table columns specified.");
    }

    if (model.values.length !== 1) {
        throw new Error("Exactly 1 value needs to be specified at this time.");
    }
}

/**
 * The cache fields of the columns of a worksheet: their names (row 1), and the distinct values (sorted) of the columns
 * named in `fieldNamesWithSharedItems`.
 *
 * @example
 * // worksheet values [
 * //   ['A', 'B', 'C', 'D', 'E'],
 * //   ['a1', 'b1', 'c1', 4, 5],
 * //   ['a1', 'b2', 'c1', 4, 5],
 * //   ['a2', 'b1', 'c2', 14, 24],
 * //   ['a2', 'b2', 'c2', 24, 35],
 * //   ['a3', 'b1', 'c3', 34, 45],
 * //   ['a3', 'b2', 'c3', 44, 45]
 * // ];
 * makeCacheFields(worksheet, ['A', 'B', 'C']);
 * // [
 * //   { name: 'A', sharedItems: ['a1', 'a2', 'a3'] },
 * //   { name: 'B', sharedItems: ['b1', 'b2'] },
 * //   { name: 'C', sharedItems: ['c1', 'c2', 'c3'] },
 * //   { name: 'D', sharedItems: null },
 * //   { name: 'E', sharedItems: null }
 * // ]
 */
function makeCacheFields(worksheet: Worksheet, fieldNamesWithSharedItems: readonly string[]): PivotCacheField[] {
    // Cache fields are used in pivot tables to reference source data.
    const names = worksheet.getRow(1).values;
    const nameToHasSharedItems = objectFromProps(fieldNamesWithSharedItems, true);

    const aggregate = (columnIndex: number): CellValue[] => {
        // the column always exists: columnIndex >= 1
        const columnValues = worksheet.getColumn(columnIndex)!.values.splice(2);
        const columnValuesAsSet = new Set(columnValues);
        return toSortedArray(columnValuesAsSet);
    };

    // make result
    const result: PivotCacheField[] = [];
    for (const columnIndex of range(1, names.length)) {
        const name = names[columnIndex];
        // a name is a property key (coerced to a string)
        const sharedItems = nameToHasSharedItems[name as string] ? aggregate(columnIndex) : null;
        result.push({ name, sharedItems });
    }
    return result;
}

export { makePivotTable };

export default { makePivotTable };
