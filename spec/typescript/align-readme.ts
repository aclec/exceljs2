// Compile-only checks (bun run test:types) of the README flows that the other align-*.ts files do not cover, written as
// a TypeScript user would write them (with the narrowing that the README JavaScript examples skip; the examples
// themselves are type-checked by `bun run check:readme`). Nothing here is executed. Lines marked @ts-expect-error must
// stay errors.

import ExcelJSDefault from "../..";
import * as ExcelJS from "../..";

// true only if A and B are the same type (holds with and without strictNullChecks)
type Equals<A, B> = (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2 ? true : false;
type Assert<T extends true> = T;

declare const workbook: ExcelJS.Workbook;
declare const worksheet: ExcelJS.Worksheet;

// --- Importing: the default import (esModuleInterop) is the namespace ------------------------------------------------
export type Importing = Assert<Equals<typeof ExcelJSDefault, typeof ExcelJS>>;

// --- Modifying Tables: getTable() is undefined for an unknown name ---------------------------------------------------
export function tables(): void {
    type GetTable = Assert<Equals<ReturnType<ExcelJS.Worksheet["getTable"]>, ExcelJS.Table | undefined>>;
    worksheet.addTable({
        name: "MyTable",
        ref: "A1",
        totalsRow: true,
        columns: [
            { name: "Date", totalsRowLabel: "Totals:", filterButton: true },
            // the cached result of a custom totals formula
            { name: "Amount", totalsRowFunction: "custom", totalsRowFormula: "SUM(B2:B3)", totalsRowResult: 140.7 },
        ],
        rows: [
            [new Date("2019-07-20"), 70.1],
            [new Date("2019-07-21"), 70.6],
        ],
    });
    const table = worksheet.getTable("MyTable");
    // @ts-expect-error -- undefined when the worksheet has no table of that name
    table.commit();
    if (!table) return;

    table.headerRow = true;
    table.totalsRow = false;
    table.ref = "D4";
    table.removeRows(0, 2);
    table.addRow([new Date("2019-08-05"), 5], 5);
    table.addRow([new Date("2019-08-10"), 10]);
    table.removeColumns(1, 1);
    table.addColumn({ name: "Letter", totalsRowFunction: "custom", totalsRowFormula: "ROW()", totalsRowResult: 6 }, ["a", "b"], 1);
    // @ts-expect-error -- a totals row result is a formula result
    table.addColumn({ name: "Bad", totalsRowResult: { formula: "ROW()" } }, []);

    const column = table.getColumn(1);
    column.name = "Code";
    column.filterButton = true;
    column.style = { font: { bold: true, name: "Comic Sans MS" } };
    column.totalsRowLabel = "Totals";
    column.totalsRowFunction = "custom";
    column.totalsRowFormula = "ROW()";
    column.totalsRowResult = 10;
    table.commit();
}

// --- Cell Comments: properties of a note given in its object form ----------------------------------------------------
export function notes(cell: ExcelJS.Cell, loaded: ExcelJS.Cell): void {
    type Note = Assert<Equals<typeof cell.note, string | ExcelJS.Comment | undefined>>;
    cell.note = "plain text";
    cell.note = {
        texts: [{ font: { size: 12, color: { theme: 0 }, name: "Calibri", family: 2, scheme: "minor" }, text: "This is " }],
        margins: { insetmode: "custom", inset: [0.25, 0.25, 0.35, 0.35] },
        protection: { locked: "True", lockText: "False" },
        editAs: "twoCells",
    };
    const note = loaded.note;
    // @ts-expect-error -- a plain-text note is a string, a cell without a note has none
    note.editAs = "oneCells";
    if (note === undefined || typeof note === "string") return;
    note.margins = { insetmode: "auto" };
    note.protection = { locked: "False", lockText: "False" };
    // @ts-expect-error -- the VML values are the strings "True" / "False"
    note.protection = { locked: false };
    note.editAs = "absolute";
}

// --- Headers and Footers: the default (null texts) also for a worksheet read from a file without them ---------------
export function headersAndFooters(): void {
    const { headerFooter } = worksheet;
    headerFooter.oddFooter = "&Lexceljs&C&F&RPage &P";
    headerFooter.differentFirst = true;
    headerFooter.firstHeader = "Hello Exceljs";
}

// --- Auto filters: a range, or its corners as addresses or row / column numbers -------------------------------------
export function autoFilters(): void {
    worksheet.autoFilter = "A1:C1";
    worksheet.autoFilter = { from: "A1", to: "C1" };
    worksheet.autoFilter = { from: { row: 3, column: 1 }, to: { row: 5, column: 12 } };
    worksheet.autoFilter = { from: "D3", to: { row: 7, column: 5 } };
    // @ts-expect-error -- a corner needs both row and column
    worksheet.autoFilter = { from: "D3", to: { row: 7 } };
}

// --- Rows: lastRow is undefined for an empty worksheet; values is a sparse array --------------------------------------
export function rows(): void {
    const lastRow = worksheet.lastRow;
    // @ts-expect-error -- undefined when the worksheet has no rows
    lastRow.height = 42.5;
    if (lastRow) lastRow.height = 42.5;
    const values = worksheet.getRow(4).values;
    const fifth: ExcelJS.CellValue | undefined = values[5];
    // addRow / insertRow: by key, by array, by sparse array, with an inherited ('i') or original ('o') style
    const sparse: ExcelJS.CellValue[] = [];
    sparse[5] = "Kyle";
    const added: ExcelJS.Row = worksheet.addRow(sparse, "i");
    const inserted: ExcelJS.Row = worksheet.insertRow(1, { id: 2, name: "Jane Doe", dob: new Date(1965, 1, 7) }, "o");
    const addedRows: ExcelJS.Row[] = worksheet.addRows([[5, "Bob", new Date()], { id: 6, name: "Barbara" }], "i");
    void [fifth, added, inserted, addedRows];
}

// --- Value Types: the type and formula type of a cell, the error values ----------------------------------------------
export function valueTypes(cell: ExcelJS.Cell): void {
    const isDate: boolean = cell.type === ExcelJS.ValueType.Date;
    // undefined for a cell that holds no formula
    type FormulaTypeOfCell = Assert<Equals<typeof cell.formulaType, ExcelJS.FormulaType | undefined>>;
    const isMaster: boolean = cell.formulaType === ExcelJS.FormulaType.Master;
    const formula: string | undefined = cell.formula;
    const result: ExcelJS.CellFormulaResult | undefined = cell.result;
    cell.value = { error: ExcelJS.ErrorValue.NotApplicable };
    cell.value = { error: "#VALUE!" };
    // @ts-expect-error -- not an error value of Excel
    cell.value = { error: "#SPILL!" };
    cell.value = { text: "Sheet2", hyperlink: "#'Sheet2'!A1" };
    cell.value = { sharedFormula: "A2", result: 10 };
    void [isDate, isMaster, formula, result];
}

// --- CSV options declared apart from the call ------------------------------------------------------------------------
export async function csvOptions(): Promise<void> {
    const readOptions: Partial<ExcelJS.CsvReadOptions> = {
        dateFormats: ["DD/MM/YYYY"],
        map(value, index) {
            return index === 1 ? new Date(value) : value;
        },
        // quoting off: null (false works as well at run time, but is not a parser option of fast-csv)
        parserOptions: { delimiter: "\t", quote: null },
    };
    await workbook.csv.readFile("in.csv", readOptions);
    // @ts-expect-error -- the parser option is a quote character or null
    await workbook.csv.readFile("in.csv", { parserOptions: { quote: false } });

    const writeOptions: Partial<ExcelJS.CsvWriteOptions> = {
        sheetName: "Page name",
        dateFormat: "DD/MM/YYYY HH:mm:ss",
        dateUTC: true,
        map(value) {
            return value instanceof Date ? value.toISOString() : value;
        },
        formatterOptions: { delimiter: "\t", quote: false },
    };
    await workbook.csv.writeFile("out.csv", writeOptions);
}

// --- Streaming XLSX Reader: the option values are literal types -----------------------------------------------------
export async function streamingReaderOptions(): Promise<void> {
    const widened = { sharedStrings: "emit", hyperlinks: "emit", worksheets: "emit" };
    // @ts-expect-error -- `const options = {...}` widens "emit" to string: declare it `as const`, or pass it inline
    void new ExcelJS.stream.xlsx.WorkbookReader("./file.xlsx", widened);
    const options = { sharedStrings: "emit", hyperlinks: "emit", worksheets: "emit" } as const;
    const reader = new ExcelJS.stream.xlsx.WorkbookReader("./file.xlsx", options);
    for await (const { eventType, value } of reader.parse()) {
        if (eventType === "worksheet") {
            for await (const row of value) void row.number;
        }
    }
}
