// Pivot tables (experimental): `worksheet.addPivotTable(...)`. One pivot table per workbook, summing one value field.

import { range, toSortedArray } from "../utils/utils";
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
    /**
     * Defined in `<pivotTableDefinition>` of xl/pivotTables/pivotTable1.xml; also used in xl/workbook.xml. Written as the
     * next free id when a pivot table read from the file has it (this object is not changed).
     */
    cacheId: string;
    /** Set when the workbook is written (xlsx.ts): the relationship id of the pivot cache definition from the workbook. */
    rId?: string;
}

/**
 * The field name of a header value (row 1 of the source worksheet): a string, or the text of a number or a boolean;
 * undefined for another value (an empty cell, a date, a rich text...), which no field name designates.
 */
function fieldName(value: CellValue | undefined): string | undefined {
    if (typeof value === "string") {
        return value;
    }
    if (typeof value === "number" || typeof value === "boolean") {
        return String(value);
    }
    return undefined;
}

/**
 * Creates a pivot table.
 * @throws Error when the workbook has one already, when the worksheet has pivot tables read from the file, for another
 *   metric than "sum", for a field name that is not in row 1
 *   of the source worksheet, without rows or columns, or without exactly one value
 */
function makePivotTable(worksheet: Worksheet, model: PivotTableInput): PivotTable {
    validate(worksheet, model);

    const { sourceSheet } = model;

    const cacheFields = makeCacheFields(sourceSheet, [...model.rows, ...model.columns]);

    // let {rows, columns, values} use indices instead of names;
    // names can then be accessed via `pivotTable.cacheFields[index].name`.
    const nameToIndex = new Map<string, number>();
    cacheFields.forEach((cacheField, index) => {
        const name = fieldName(cacheField.name);
        if (name !== undefined) {
            nameToIndex.set(name, index);
        }
    });
    const indexOf = (name: string): number => {
        const index = nameToIndex.get(name);
        if (index === undefined) {
            // validate() checked that the names are those of the cache fields
            throw new Error(`The header name "${name}" was not found in ${sourceSheet.name}.`);
        }
        return index;
    };

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
    // the pivot table added has a fixed name and location (A3:E15), which could be those of the pivot tables read
    if (worksheet._pivotTableParts.length) {
        throw new Error(
            `The worksheet ${worksheet.name} already has pivot tables read from the file: add the pivot table to another worksheet.`
        );
    }

    if (model.metric && model.metric !== "sum") {
        throw new Error('Only the "sum" metric is supported at this time.');
    }

    const headerNames = new Set<string>();
    model.sourceSheet
        .getRow(1)
        .values.slice(1)
        .forEach((value) => {
            const name = fieldName(value);
            if (name !== undefined) {
                headerNames.add(name);
            }
        });
    for (const name of [...model.rows, ...model.columns, ...model.values]) {
        if (!headerNames.has(name)) {
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
    const withSharedItems = new Set(fieldNamesWithSharedItems);

    const aggregate = (columnIndex: number): CellValue[] => {
        const columnValues = worksheet._getColumnEx(columnIndex).values.splice(2);
        const columnValuesAsSet = new Set(columnValues);
        return toSortedArray(columnValuesAsSet);
    };

    // make result
    const result: PivotCacheField[] = [];
    for (const columnIndex of range(1, names.length)) {
        const name = names[columnIndex];
        const key = fieldName(name);
        const sharedItems = key !== undefined && withSharedItems.has(key) ? aggregate(columnIndex) : null;
        result.push({ name, sharedItems });
    }
    return result;
}

export { makePivotTable };

export default { makePivotTable };
