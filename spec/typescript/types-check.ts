// Compile-only checks for the public typings (index.d.ts). Nothing here is executed.
//
// Run with each supported TypeScript major (no tsconfig, no `types` option, like a consumer
// whose tsconfig does not list @types/node explicitly):
//   bunx tsc --noEmit --strict --target es2022 --module commonjs spec/typescript/types-check.ts
//   (TypeScript >= 6: add --ignoreConfig)
//
// Lines marked @ts-expect-error must stay errors: they guard against typings becoming too loose.

import * as fs from "fs";
import * as ExcelJS from "../..";

// true only if A and B are the same type (holds with and without strictNullChecks)
type Equals<A, B> = (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2 ? true : false;

declare const workbook: ExcelJS.Workbook;
declare const worksheet: ExcelJS.Worksheet;

// --- the package resolves to index.d.ts and Node stream types are available -------------------
async function nodeStreams(): Promise<void> {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.read(fs.createReadStream("in.xlsx"));
    await wb.xlsx.write(fs.createWriteStream("out.xlsx"));
    await wb.csv.read(fs.createReadStream("in.csv"));
}

// --- binary inputs: Node Buffers (TS >= 5.7 typed arrays), ArrayBuffer and Uint8Array ---------
async function binaryInputs(): Promise<void> {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(fs.readFileSync("in.xlsx"));
    await wb.xlsx.load(new ArrayBuffer(8));
    await wb.xlsx.load(new Uint8Array(8));
    wb.addImage({ buffer: fs.readFileSync("image.png"), extension: "png" });
    const output = await wb.xlsx.writeBuffer();
    fs.writeFileSync("out.xlsx", Buffer.from(output));
    // a Node Buffer at run time (the Buffer polyfill, a Uint8Array, in the browser bundle)
    const nodeBuffer: Buffer = await wb.xlsx.writeBuffer();
    const cast = (await wb.xlsx.writeBuffer()) as Buffer;
    fs.writeFileSync("out.xlsx", output);
    const csv: Buffer = await wb.csv.writeBuffer();
    // browsers: a Blob part (TypeScript 5.7+ DOM typings want an ArrayBuffer-backed view)
    const blob = new Blob([output], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    void [nodeBuffer, cast, csv, blob];
}

// --- Address.col / Address.row are numbers (upstream #2720) -----------------------------------
function addresses(): void {
    const cell = worksheet.getCell("B3");
    const col: number = cell.col;
    const row: number = cell.row;
    const address: ExcelJS.Address = { address: "B3", col: 2, row: 3, $col$row: "$B$3" };
    // @ts-expect-error column numbers are not letters
    const letter: string = cell.col;
    void [col, row, address, letter];
}

// --- image editAs (upstream #2983) -------------------------------------------------------------
function images(imageId: number): void {
    const editAs: ExcelJS.ImageEditAs = "twoCell";
    worksheet.addImage(imageId, { tl: { col: 0, row: 0 }, ext: { width: 10, height: 10 }, editAs });
    worksheet.addImage(imageId, { tl: { col: 0, row: 0 }, ext: { width: 10, height: 10 }, editAs: "oneCell" });
    worksheet.addImage(imageId, { tl: { col: 0, row: 0 }, ext: { width: 10, height: 10 }, editAs: "absolute" });
    // @ts-expect-error 'twoCells' is a comment anchor value, not an image one
    worksheet.addImage(imageId, { tl: { col: 0, row: 0 }, ext: { width: 10, height: 10 }, editAs: "twoCells" });
}

// --- conditional formatting: all cellIs operators, optional priority, data bar colours ---------
// (upstream #2736, #2655)
function conditionalFormatting(): void {
    const operators: ExcelJS.CellIsOperators[] = [
        "equal",
        "notEqual",
        "greaterThan",
        "greaterThanOrEqual",
        "lessThan",
        "lessThanOrEqual",
        "between",
        "notBetween",
    ];
    worksheet.addConditionalFormatting({
        ref: "A1:A10",
        rules: operators.map((operator): ExcelJS.ConditionalFormattingRule => ({
            type: "cellIs",
            operator,
            formulae: [1],
            stopIfTrue: true,
        })),
    });
    worksheet.addConditionalFormatting({
        ref: "B1:B10",
        rules: [
            {
                type: "dataBar",
                cfvo: [{ type: "min" }, { type: "max" }],
                color: { argb: "FF638EC6" },
                borderColor: { argb: "FF638EC6" },
                negativeFillColor: { argb: "FFFF0000" },
                negativeBorderColor: { argb: "FFFF0000" },
                axisColor: { argb: "FF000000" },
            },
        ],
    });
}

// --- Column.values / Row.values / Row.dimensions (upstream #2587, #2596) -----------------------
function rowsAndColumns(): void {
    const column = worksheet.getColumn(1);
    const columnValues: ExcelJS.CellValue[] = column.values;
    column.values = [1, "two", null, new Date()];

    const row = worksheet.getRow(1);
    const rowValues: ExcelJS.CellValue[] = row.values;
    row.values = [1, 2, 3];
    row.values = { id: 1, name: "John" }; // the setter still accepts values keyed by column key

    const dimensions = row.dimensions;
    const nullable: Equals<typeof dimensions, { min: number; max: number } | null> = true;
    if (dimensions) {
        const width: number = dimensions.max - dimensions.min + 1;
        void width;
    }
    void [columnValues, rowValues, nullable];
}

// --- Worksheet.protect arguments are optional (upstream #2664) ---------------------------------
async function protection(): Promise<void> {
    await worksheet.protect();
    await worksheet.protect("secret");
    await worksheet.protect(undefined, { selectLockedCells: true });
    await worksheet.protect("secret", { spinCount: 1000 });
}

// --- DataValidationType (upstream #2697) -------------------------------------------------------
function dataValidation(): void {
    const type: ExcelJS.DataValidationType = "whole";
    worksheet.getCell("A1").dataValidation = { type, operator: "between", formulae: [1, 10] };
    // @ts-expect-error unknown validation type
    const bad: ExcelJS.DataValidationType = "integer";
    void bad;
    // ST_DataValidationType time (kept when read from a file)
    worksheet.getCell("A2").dataValidation = { type: "time", operator: "lessThan", formulae: ["0.5"] };

    // worksheet.dataValidations (upgrade guide: read range keys from its model)
    const dataValidations: ExcelJS.DataValidations = worksheet.dataValidations;
    const added: ExcelJS.DataValidation = dataValidations.add("B1:B10", { type: "list", formulae: ['"a,b"'] });
    const found: ExcelJS.DataValidation | undefined = dataValidations.find("B5");
    dataValidations.remove("B5");
    const model: { [address: string]: ExcelJS.DataValidation | undefined } = dataValidations.model;
    dataValidations.model["C1:C10"] = { type: "whole", operator: "greaterThan", formulae: [0] };
    void [added, found, model];
}

// --- tables: theme may be null for "no style" (upstream #2061) ---------------------------------
function tables(): void {
    worksheet.addTable({
        name: "Plain",
        ref: "A1",
        style: { theme: null },
        columns: [{ name: "Value" }],
        rows: [[1]],
    });
    worksheet.addTable({
        name: "Styled",
        ref: "C1",
        style: { theme: "TableStyleLight1" },
        columns: [{ name: "Value" }],
        rows: [[1]],
    });
}

// --- streaming reader (upstream #1922, #1901, #3097) -------------------------------------------
async function streamingReader(): Promise<void> {
    const reader = new ExcelJS.stream.xlsx.WorkbookReader("in.xlsx"); // options are optional
    const withOptions = new ExcelJS.stream.xlsx.WorkbookReader(fs.createReadStream("in.xlsx"), {
        sharedStrings: "emit",
    });
    // event interface documented in the README
    reader.on("worksheet", (worksheetReader: ExcelJS.stream.xlsx.WorksheetReader) => {
        worksheetReader.on("row", (row: ExcelJS.Row) => row.number);
    });
    reader.on("end", () => undefined);
    await reader.read();
    for await (const worksheetReader of reader) {
        const name: string = worksheetReader.name;
        const id: number | string = worksheetReader.id;
        const state: ExcelJS.WorksheetState | undefined = worksheetReader.state;
        const dimensions: ExcelJS.Range = worksheetReader.dimensions;
        // columns is null until the reader has seen <cols>
        const columns: Equals<typeof worksheetReader.columns, ExcelJS.Column[] | null> = true;
        worksheetReader.on("row", (row: ExcelJS.Row) => row.number);
        for await (const row of worksheetReader) {
            const rowNumber: number = row.number;
            void rowNumber;
        }
        void [name, id, state, dimensions, columns];
    }
    void withOptions;
}

// --- removed APIs ------------------------------------------------------------------------------
function removed(): void {
    // @ts-expect-error createInputStream has thrown since 4.0; use csv.read(stream)
    workbook.csv.createInputStream();
}

void [nodeStreams, binaryInputs, addresses, images, conditionalFormatting, rowsAndColumns, protection];
void [dataValidation, tables, streamingReader, removed];
