// Compile-only checks of the declarations corrected after the review of the index.d.ts alignment. Nothing here is
// executed. Lines marked @ts-expect-error must stay errors.

import * as ExcelJS from "../..";

declare const workbook: ExcelJS.Workbook;
declare const worksheet: ExcelJS.Worksheet;

// --- pageSetup.paperSize: any paper size number of the standard, not only the listed ones ------------
worksheet.pageSetup.paperSize = ExcelJS.PaperSize.A4;
worksheet.pageSetup.paperSize = 8; // A3
worksheet.pageSetup.paperSize = undefined;
// @ts-expect-error - a paper size is a number
worksheet.pageSetup.paperSize = "A3";

// --- Range: the addresses and the count are getters -------------------------------------------------
const range = worksheet.dimensions;
export const addresses: string[] = [
    range.tl,
    range.$t$l,
    range.br,
    range.$b$r,
    range.range,
    range.$range,
    range.shortRange,
    range.$shortRange,
];
export const cellCount: number = range.count;
// @ts-expect-error - getter only: setting it throws
range.tl = "B2";
// @ts-expect-error - getter only: setting it throws
range.count = 3;

// --- xlsx.load with base64: a string or a Node Buffer of the text, not a plain Uint8Array ---------
declare const base64Text: string;
export const fromString: Promise<ExcelJS.Workbook> = workbook.xlsx.load(base64Text, { base64: true });
export const fromBuffer: Promise<ExcelJS.Workbook> = workbook.xlsx.load(Buffer.from(base64Text), { base64: true });
// @ts-expect-error - a plain Uint8Array is not decoded as text (its toString() is "85,69,...")
export const fromBytes = workbook.xlsx.load(new Uint8Array(4), { base64: true });

// --- ModelContainer: created empty to read a file into it -------------------------------------------
const container = new ExcelJS.ModelContainer();
export const read: Promise<ExcelJS.ModelContainer> = container.xlsx.readFile("in.xlsx");
export const model: ExcelJS.WorkbookModel | undefined = container.model;
