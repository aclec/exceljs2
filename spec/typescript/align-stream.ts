// Compile-only checks of the streaming API declarations (index.d.ts `stream.xlsx`) aligned with the runtime. Nothing
// here is executed. Lines marked @ts-expect-error must stay errors.

import * as fs from "fs";
import * as ExcelJS from "../..";

// true only if A and B are the same type (holds with and without strictNullChecks)
type Equals<A, B> = (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2 ? true : false;

type WorkbookWriter = ExcelJS.stream.xlsx.WorkbookWriter;
type WorksheetWriter = ExcelJS.stream.xlsx.WorksheetWriter;
type WorksheetReader = ExcelJS.stream.xlsx.WorksheetReader;

// --- WorkbookWriter: its own class, not a Workbook ---------------------------------------------
async function workbookWriter(): Promise<void> {
    const empty = new ExcelJS.stream.xlsx.WorkbookWriter(); // options are optional
    const writer = new ExcelJS.stream.xlsx.WorkbookWriter({
        filename: "out.xlsx",
        useSharedStrings: true,
        useStyles: true,
        zip: { level: 9 },
        maxMergedCells: 10,
        creator: "me",
        created: new Date(),
    });
    new ExcelJS.stream.xlsx.WorkbookWriter({ stream: fs.createWriteStream("out.xlsx") });
    // @ts-expect-error the output is written to the stream: a readable stream is not one
    new ExcelJS.stream.xlsx.WorkbookWriter({ stream: fs.createReadStream("in.xlsx") });

    const worksheet = writer.addWorksheet("Data", { useSharedStrings: false, properties: { tabColor: { argb: "FF00FF00" } } });
    const isWorksheetWriter: Equals<typeof worksheet, WorksheetWriter> = true;
    const found: Equals<ReturnType<WorkbookWriter["getWorksheet"]>, WorksheetWriter | undefined> = true;
    const imageId: number = writer.addImage({ extension: "png", filename: "image.png" });
    const image: ExcelJS.stream.xlsx.WorkbookWriterImage | undefined = writer.getImage(imageId);
    writer.definedNames.add("Data!$A$1", "first");

    // commit resolves with the writer
    const committed: WorkbookWriter = await writer.commit();

    // @ts-expect-error a WorkbookWriter is not a Workbook: no xlsx / csv
    void writer.xlsx;
    // @ts-expect-error a WorkbookWriter is not a Workbook: no eachSheet
    writer.eachSheet(() => undefined);
    // @ts-expect-error a WorkbookWriter is not a Workbook: no worksheets array
    void writer.worksheets;

    void [empty, isWorksheetWriter, found, image, committed];
}

// --- WorksheetWriter: rows are committed in order --------------------------------------------
async function worksheetWriter(worksheet: WorksheetWriter): Promise<void> {
    worksheet.columns = [
        { header: "Id", key: "id", width: 10 },
        { header: "Name", key: "name", width: 32 },
    ];
    const row: ExcelJS.Row = worksheet.addRow({ id: 1, name: "a" });
    await row.commit();
    const lastRow: Equals<WorksheetWriter["lastRow"], ExcelJS.Row | undefined> = true;
    const cell: ExcelJS.Cell = worksheet.getCell("B3");
    const maybeCell: ExcelJS.Cell | undefined = worksheet.findCell(3, 2);
    const column: ExcelJS.Column = worksheet.getColumn("id");
    worksheet.mergeCells("C3:D4");
    worksheet.mergeCells("E5", "F6");
    worksheet.mergeCells(7, 1, 8, 2);
    worksheet.mergeCells(["G1", "H2"]);
    worksheet.dataValidations.add("A1", { type: "whole", allowBlank: true, formulae: [1] });
    worksheet.removeConditionalFormatting();
    await worksheet.protect("secret", { spinCount: 10 });
    const committed: Promise<void> = worksheet.commit();
    const destroy: Equals<ReturnType<WorksheetWriter["destroy"]>, never> = true;

    // @ts-expect-error a WorksheetWriter is not a Worksheet: no addRows
    worksheet.addRows([[1]]);
    // @ts-expect-error a WorksheetWriter is not a Worksheet: no getSheetValues
    worksheet.getSheetValues();
    // @ts-expect-error addRow takes no style
    worksheet.addRow([1], "i");

    void [lastRow, cell, maybeCell, column, committed, destroy];
}

// --- WorkbookReader: an EventEmitter with typed events, async iterable --------------------------
async function workbookReader(): Promise<void> {
    const later = new ExcelJS.stream.xlsx.WorkbookReader(); // the input can be given to read() / parse()
    await later.read("in.xlsx", { sharedStrings: "emit" });
    const reader = new ExcelJS.stream.xlsx.WorkbookReader(fs.createReadStream("in.xlsx"), {
        sharedStrings: "emit",
        hyperlinks: "emit",
        entries: "emit",
        maxEntrySize: 1024,
    });
    const options: Equals<typeof reader.options.sharedStrings, "cache" | "emit" | "ignore"> = true;
    reader.on("worksheet", (worksheet) => {
        const isReader: Equals<typeof worksheet, WorksheetReader> = true;
        worksheet.on("row", (row) => {
            const isRow: Equals<typeof row, ExcelJS.Row> = true;
            void isRow;
        });
        worksheet.on("hyperlink", (hyperlink: ExcelJS.stream.xlsx.WorksheetHyperlink) => hyperlink.ref);
        void isReader;
    });
    reader.on("shared-strings", ({ index, text }) => {
        const value: string | ExcelJS.CellRichTextValue = text;
        void [index, value];
    });
    reader.on("hyperlinks", (hyperlinks) => {
        hyperlinks.on("hyperlink", (relationship) => relationship.target);
    });
    reader.on("entry", (entry) => entry.type);
    reader.on("error", (error: unknown) => error);
    // @ts-expect-error unknown event
    reader.on("rows", () => undefined);
    // @ts-expect-error a worksheet event gives a WorksheetReader
    reader.on("worksheet", (worksheet: ExcelJS.Worksheet) => worksheet);

    for await (const event of reader.parse()) {
        if (event.eventType === "worksheet") {
            const isReader: Equals<typeof event.value, WorksheetReader> = true;
            void isReader;
        }
    }
    for await (const worksheet of reader) {
        // undefined only for an invalid number (below 1, not an integer): declared Column
        const column: ExcelJS.Column = worksheet.getColumn(1);
        const hyperlinks: Record<string, ExcelJS.stream.xlsx.WorksheetHyperlink> | undefined = worksheet.hyperlinks;
        // hyperlinks: "cache" gives the targets of the external links
        const target: Equals<ExcelJS.stream.xlsx.WorksheetHyperlink["target"], string | undefined> = true;
        void target;
        const outline: number = worksheet.properties.outlineLevelRow;
        for await (const events of worksheet.parse()) {
            const batch: ExcelJS.stream.xlsx.WorksheetReaderEvent[] = events;
            void batch;
        }
        void [column, hyperlinks, outline];
    }
    void options;
}

// --- charts and chartsheets of the streaming writer: the options of the document API -----------------
function streamCharts(writer: WorkbookWriter, worksheet: WorksheetWriter): void {
    const chart: ExcelJS.WorksheetChart = worksheet.addChart({ type: "column", anchor: "D2:K16", series: [{ values: "B2:B10" }] });
    const charts: ExcelJS.WorksheetChart[] = worksheet.getCharts();
    chart.remove();
    // @ts-expect-error -- a chart of a worksheet has an anchor
    worksheet.addChart({ type: "column", series: [{ values: "B2:B10" }] });
    const chartsheet: ExcelJS.Chartsheet = writer.addChartsheet(
        "Chart",
        { type: "line", series: [{ values: "Data!B2:B10" }] },
        { state: "veryHidden" }
    );
    void [charts, chartsheet];
}

// --- removed: the worksheet and hyperlink readers are not runtime values ----------------------
function removed(): void {
    // @ts-expect-error WorksheetReader is a type only: it is created by the WorkbookReader
    void ExcelJS.stream.xlsx.WorksheetReader;
    const zlib: ExcelJS.stream.xlsx.ZlibOptions = { level: 9 };
    // @ts-expect-error only `level` is read from the deprecated zlib options
    const chunked: ExcelJS.stream.xlsx.ZlibOptions = { chunkSize: 1024 };
    void [zlib, chunked];
}

void [workbookWriter, worksheetWriter, workbookReader, streamCharts, removed];
