// Compile-only checks of the workbook and file I/O declarations of index.d.ts, aligned with the runtime. Nothing here
// is executed. Lines marked @ts-expect-error must stay errors.

import * as fs from "fs";
import { PassThrough } from "stream";
import * as ExcelJS from "../..";

// true only if A and B are the same type
type Equals<A, B> = (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2 ? true : false;
/** Compiles only for `true`. */
type Assert<T extends true> = T;

declare const workbook: ExcelJS.Workbook;

// --- document properties: undefined after loading a file without them ---------------------------
export type DocumentProperties = [
    Assert<Equals<ExcelJS.Workbook["title"], string | undefined>>,
    Assert<Equals<ExcelJS.Workbook["creator"], string | undefined>>,
    Assert<Equals<ExcelJS.Workbook["created"], Date | undefined>>,
    Assert<Equals<ExcelJS.Workbook["lastPrinted"], Date | undefined>>,
    Assert<Equals<ExcelJS.Workbook["revision"], number | undefined>>,
    Assert<Equals<ExcelJS.Workbook["language"], string | undefined>>,
    Assert<Equals<ExcelJS.Workbook["contentStatus"], string | undefined>>,
    Assert<Equals<ExcelJS.Workbook["properties"], Partial<ExcelJS.WorkbookProperties>>>,
    Assert<Equals<ExcelJS.Workbook["calcProperties"], Partial<ExcelJS.CalculationProperties>>>,
    Assert<Equals<ExcelJS.Workbook["views"], ExcelJS.WorkbookView[]>>,
    Assert<Equals<ExcelJS.WorkbookModel["views"], ExcelJS.WorkbookView[]>>,
    Assert<Equals<ExcelJS.WorkbookView["activeTab"], number | undefined>>,
    Assert<Equals<ExcelJS.WorkbookModel["revision"], number | undefined>>,
    Assert<Equals<ExcelJS.WorkbookModel["themes"], Record<string, string> | undefined>>,
];

function documentProperties(): void {
    // @ts-expect-error -- undefined after loading a file without a title
    workbook.title.trim();
    workbook.title?.trim();
    workbook.title = "Report";
    workbook.properties = {};
    workbook.calcProperties = { fullCalcOnLoad: true };
    // [] also after loading a file without <bookViews>
    workbook.views.push({ x: 0, y: 0, width: 10000, height: 20000, visibility: "visible" });
    workbook.views = [{ x: 0, y: 0, width: 10000, height: 20000, visibility: "visible" }];
    // @ts-expect-error -- a getter: assigning throws
    workbook.worksheets = [];
}

// --- images ---------------------------------------------------------------------------------------
export type Images = [
    Assert<Equals<ReturnType<ExcelJS.Workbook["getImage"]>, ExcelJS.Media | undefined>>,
    Assert<Equals<ExcelJS.Media["type"], "image">>,
    Assert<Equals<ExcelJS.Media["buffer"], ArrayBuffer | Uint8Array | null | undefined>>,
    Assert<Equals<ExcelJS.Workbook["media"], ExcelJS.Media[]>>,
];

function images(): void {
    workbook.addImage({ extension: "png", base64: "data:image/png;base64,AAAA" });
    workbook.addImage({ extension: "jpeg", buffer: new Uint8Array(8) });
    workbook.addImage({ extension: "bmp", filename: "image.bmp" });
    // @ts-expect-error -- the extension is required
    workbook.addImage({ buffer: new ArrayBuffer(8) });
    const image = workbook.getImage(0);
    // @ts-expect-error -- undefined for an unknown id
    image.extension.toUpperCase();
}

// --- xlsx: read methods resolve with the workbook, load accepts bytes, a Blob or base64 ------------
async function xlsx(): Promise<void> {
    const wb: ExcelJS.Workbook = await workbook.xlsx.load(new ArrayBuffer(8));
    await workbook.xlsx.load(fs.readFileSync("in.xlsx"));
    await workbook.xlsx.load(new DataView(new ArrayBuffer(8)));
    await workbook.xlsx.load(new Blob([new Uint8Array(8)]));
    await workbook.xlsx.load([0x50, 0x4b]);
    await workbook.xlsx.load("UEsDBA==", { base64: true });
    // @ts-expect-error -- a string is base64 text: the option is required
    await workbook.xlsx.load("UEsDBA==");
    await workbook.xlsx.read(fs.createReadStream("in.xlsx"));
    await workbook.xlsx.read((async function* () {})());
    const read: ExcelJS.Workbook = await workbook.xlsx.readFile("in.xlsx", { maxMergedCells: 10, ignoreNodes: ["dataValidations"] });
    const output = new PassThrough();
    await workbook.xlsx.write(output, { useStyles: false, zip: { level: 9 } });
    await workbook.xlsx.write({ write: (chunk: Buffer) => chunk.length, end: () => undefined });
    // @ts-expect-error -- the output must be writable
    await workbook.xlsx.write(fs.createReadStream("in.xlsx"));
    // @ts-expect-error -- not an option of xlsx.write (an option of the streaming WorkbookWriter)
    await workbook.xlsx.writeBuffer({ filename: "out.xlsx" });
    void wb;
    void read;
}

// --- csv: readable / writable streams, typed map functions ------------------------------------------
async function csv(): Promise<void> {
    const worksheet: ExcelJS.Worksheet = await workbook.csv.readFile("in.csv", {
        map: (value: string, index: number) => (index === 0 ? Number(value) : value),
        parserOptions: { headers: true, delimiter: ";" },
    });
    await workbook.csv.read(fs.createReadStream("in.csv"));
    // @ts-expect-error -- the input must be readable
    await workbook.csv.read(fs.createWriteStream("out.csv"));
    await workbook.csv.write(fs.createWriteStream("out.csv"), {
        map: (value: ExcelJS.CellValue) => (value instanceof Date ? value.getTime() : value),
        formatterOptions: {
            writeHeaders: false,
            transform: (row: unknown[]) => row.map(String),
        },
    });
    await workbook.csv.writeBuffer({
        formatterOptions: {
            transform: (row: unknown[], callback: (error?: Error | null, row?: unknown[]) => void) => callback(null, row),
        },
    });
    // @ts-expect-error -- the output must be writable
    await workbook.csv.write(fs.createReadStream("in.csv"));
    void worksheet;
}

// --- defined names: ranges given by their bounds (or a decoded reference) can be added and removed -
function definedNames(): void {
    const names = workbook.definedNames;
    names.addEx({ sheetName: "Sheet1", top: 1, left: 1, bottom: 2, right: 2 }, "range");
    names.addEx({ sheetName: "Sheet1", address: "A1", row: 1, col: 1 }, "cell");
    names.removeEx({ sheetName: "Sheet1", address: "A1", row: 1, col: 1 }, "cell");
    // a range given by its bounds is removed too
    names.removeEx({ sheetName: "Sheet1", top: 1, left: 1, bottom: 2, right: 2 }, "range");
    names.removeAllNames({ sheetName: "Sheet1", top: 1, left: 1, bottom: 2, right: 2 });
    names.remove("Sheet1!$A$1:$B$2", "range");
}

// --- ModelContainer: exported by the Node.js entry point ---------------------------------------------
async function modelContainer(): Promise<void> {
    const container = new ExcelJS.ModelContainer(workbook.model);
    const read: ExcelJS.ModelContainer = await container.xlsx.readFile("in.xlsx");
    const bytes: ExcelJS.Buffer = await container.xlsx.writeBuffer();
    void read;
    void bytes;
}

export { documentProperties, images, xlsx, csv, definedNames, modelContainer };
