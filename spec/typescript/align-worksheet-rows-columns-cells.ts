// Compile-only checks for the worksheet rows, columns and cells API of index.d.ts (aligned with the runtime).
// Nothing here is executed. Lines marked @ts-expect-error must stay errors: the runtime rejects or mishandles them.

import * as ExcelJS from "../..";

// true only if A and B are the same type (holds with and without strictNullChecks)
type Equals<A, B> = (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2 ? true : false;

declare const worksheet: ExcelJS.Worksheet;
declare const range: ExcelJS.Range;

// --- getCell / findCell: an address, or (row, col) numbers -------------------------------------
function cells(): void {
    const byAddress: ExcelJS.Cell = worksheet.getCell("B2");
    const byNumbers: ExcelJS.Cell = worksheet.getCell(2, 2);
    const found: Equals<ReturnType<typeof worksheet.findCell>, ExcelJS.Cell | undefined> = true;
    worksheet.findCell("B2");
    worksheet.findCell(2, 2);
    // @ts-expect-error a column letter with a row number throws at run time
    worksheet.getCell(2, "B");
    // @ts-expect-error a row number without a column throws at run time
    worksheet.getCell(2);
    // @ts-expect-error a column letter with a row number is never found
    worksheet.findCell(2, "B");
    void [byAddress, byNumbers, found];
}

// --- rows --------------------------------------------------------------------------------------
function rows(rowValues: readonly ExcelJS.RowValues[]): void {
    worksheet.eachRow((row, rowNumber) => void [row.number, rowNumber]);
    worksheet.eachRow({ includeEmpty: true }, (row) => void row.number);
    worksheet.eachRow({}, (row) => void row.number);
    worksheet.eachRow(null, (row) => void row.number);
    worksheet.addRow(null);
    worksheet.addRow([1, "a"], "i+");
    worksheet.addRows(rowValues);
    // @ts-expect-error not a style option
    worksheet.addRow([1], "x");
    const inserted: Equals<ReturnType<typeof worksheet.insertRows>, ExcelJS.Row[] | undefined> = true;
    worksheet.insertRows(1, rowValues, "o+");
    worksheet.duplicateRow(1, 2);
    worksheet.duplicateRow(1, 2, true);
    const sheetValues: Equals<ReturnType<typeof worksheet.getSheetValues>, (ExcelJS.CellValue[] | undefined)[]> = true;
    const rowsOrNot: Equals<ReturnType<typeof worksheet.getRows>, ExcelJS.Row[] | undefined> = true;
    void [inserted, sheetValues, rowsOrNot];
}

// --- columns -----------------------------------------------------------------------------------
function columns(definitions: readonly ExcelJS.ColumnDefinition[], values: readonly ExcelJS.CellValue[]): void {
    worksheet.columns = definitions;
    const list: Equals<typeof worksheet.columns, ExcelJS.Column[] | null> = true;
    // undefined only for an invalid number (not an integer >= 1)
    const column: Equals<ReturnType<typeof worksheet.getColumn>, ExcelJS.Column> = true;
    worksheet.spliceColumns(1, 0, values, [1, "a"]);
    void [list, column];
}

// --- merges: the arguments of a range ----------------------------------------------------------
function merges(location: ExcelJS.Location): void {
    worksheet.mergeCells("A1:B2");
    worksheet.mergeCells("Sheet1!A1:B2");
    worksheet.mergeCells("A1", "B2");
    worksheet.mergeCells(1, 1, 2, 2);
    worksheet.mergeCells(1, 1, 2, 2, "Sheet1");
    worksheet.mergeCells(["A1", "B2"]);
    worksheet.mergeCells([1, 1, 2, 2]);
    worksheet.mergeCells(location);
    worksheet.mergeCells(range);
    worksheet.mergeCellsWithoutStyle("A1:B2");
    worksheet.unMergeCells("A1");
    worksheet.unMergeCells(1, 1, 2, 2);
    // @ts-expect-error two numbers are not a range (throws at run time)
    worksheet.mergeCells(1, 2);
    // @ts-expect-error three numbers are not a range
    worksheet.unMergeCells(1, 2, 3);
    // @ts-expect-error a merge needs a range (throws at run time)
    worksheet.mergeCells();
    // @ts-expect-error an unmerge needs a range (throws at run time)
    worksheet.unMergeCells();
    // @ts-expect-error a merge needs a range (throws at run time)
    worksheet.mergeCellsWithoutStyle();
}

// --- fillFormula: a range reference, results by cell, by row or flat, shared or array ----------
function formulas(location: ExcelJS.Location): void {
    worksheet.fillFormula("A1:B2", "C1*2");
    worksheet.fillFormula("A1:A3", "C1*2", [1, "two", true]);
    worksheet.fillFormula("A1:B2", "C1*2", [
        [1, 2],
        [3, 4],
    ]);
    worksheet.fillFormula("A1:B2", "C1*2", (row, col) => row * col);
    worksheet.fillFormula("A1:B2", "C1*2", () => undefined, "array");
    // @ts-expect-error fillFormula takes a reference string only (a Location throws at run time)
    worksheet.fillFormula(location, "C1*2");
    // @ts-expect-error fillFormula takes a reference string only (a Range throws at run time)
    worksheet.fillFormula(range, "C1*2");
}

// --- Range -------------------------------------------------------------------------------------
function ranges(location: ExcelJS.Location): void {
    const intersects: boolean = range.intersects(location) || range.intersects(range);
    const contains: boolean = range.containsEx({ row: 1, col: 1 }) || range.containsEx({ row: 1, col: 1, sheetName: "S" });
    // @ts-expect-error a cell needs its row and its column
    range.containsEx({ row: 1 });
    range.setTLBR("A1", "B2");
    range.setTLBR(1, 1, 2, 2, "Sheet1");
    range.forEachAddress((address, row, col) => void [address, row, col]);
    range.expandRow(undefined);
    range.expandRow(worksheet.getRow(1));
    void [intersects, contains];
}

void [cells, rows, columns, merges, formulas, ranges];
