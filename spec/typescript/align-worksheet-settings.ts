// Compile-only checks for the worksheet settings of index.d.ts (views, page setup, properties, protection, images,
// conditional formattings, model), aligned with the runtime. Nothing here is executed.
//
// Lines marked @ts-expect-error must stay errors: they guard against typings becoming too loose.

import * as ExcelJS from "../..";

// true only if A and B are the same type (holds with and without strictNullChecks)
type Equals<A, B> = (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2 ? true : false;

declare const workbook: ExcelJS.Workbook;
declare const worksheet: ExcelJS.Worksheet;

// --- views: [] also for a worksheet read without <sheetViews>; flat model with style / workbookViewId / tabSelected --
function views(): void {
    const list: Equals<typeof worksheet.views, ExcelJS.WorksheetViewModel[]> = true;
    worksheet.views.push({ state: "frozen", ySplit: 1 });
    worksheet.views = [{ state: "frozen", xSplit: 1, ySplit: 1, style: "pageLayout", tabSelected: true, workbookViewId: 0 }];
    worksheet.views = [{ state: "normal" }];
    const style: ExcelJS.WorksheetViewModel["style"] = "pageBreakPreview";
    const view: ExcelJS.WorksheetView = { ...viewCommon(), state: "normal" };
    worksheet.views = [view];
    // @ts-expect-error an unknown style
    worksheet.views = [{ style: "normal" }];
    void list;
    void style;
}

function viewCommon(): ExcelJS.WorksheetViewCommon {
    return {
        rightToLeft: false,
        activeCell: "A1",
        showRuler: true,
        showRowColHeaders: true,
        showGridLines: true,
        zoomScale: 100,
        zoomScaleNormal: 100,
    };
}

// --- header and footer: the default also for a worksheet read without <headerFooter> ----------------------------------
function headerFooter(): void {
    const defined: Equals<typeof worksheet.headerFooter, ExcelJS.WorksheetHeaderFooter> = true;
    worksheet.headerFooter.oddFooter = "&P";
    worksheet.headerFooter = { oddHeader: "&CPage &P", evenHeader: null };
    // the helpers of the header and footer texts
    worksheet.headerFooter.oddHeader = ExcelJS.headerFooter.build({ left: ExcelJS.headerFooter.escape("AT&T"), right: "&[Page]" });
    const sections: ExcelJS.HeaderFooterSections = ExcelJS.headerFooter.parse("&LLeft&RPage &P");
    const left: Equals<typeof sections.left, string | undefined> = true;
    // @ts-expect-error the sections are strings
    ExcelJS.headerFooter.build({ center: 1 });
    void defined;
    void left;
}

// --- page setup: margins null when read without <pageMargins>; printer defaults, copies, first page number ----------
function pageSetup(): void {
    const margins: Equals<typeof worksheet.pageSetup.margins, ExcelJS.Margins | null | undefined> = true;
    worksheet.pageSetup = { copies: 2, usePrinterDefaults: true, cellComments: "none", firstPageNumber: 3 };
    const used: boolean | undefined = worksheet.pageSetup.useFirstPageNumber;
    // @ts-expect-error not a value of cellComments
    worksheet.pageSetup.cellComments = "asDisplyed";
    // ST_Orientation: a file may give the default orientation of the printer
    const orientation: Equals<typeof worksheet.pageSetup.orientation, "portrait" | "landscape" | "default" | undefined> = true;
    // @ts-expect-error not a value of ST_Orientation
    worksheet.pageSetup.orientation = "upright";
    void orientation;
    void margins;
    void used;
}

// --- properties: outline properties set or read one by one; showGridLines did not exist ------------------------------
function properties(): void {
    worksheet.properties.outlineProperties = { summaryBelow: false };
    const below: Equals<NonNullable<typeof worksheet.properties.outlineProperties>["summaryBelow"], boolean | undefined> = true;
    // @ts-expect-error has no effect: the grid lines are a setting of the views
    worksheet.properties = { showGridLines: false };
    void below;
}

// --- addWorksheet options: views, header and footer as the worksheet holds them, auto filter ------------------------
function addWorksheetOptions(): void {
    workbook.addWorksheet("s", {
        views: [{ state: "split", xSplit: 2000, activePane: "topRight", style: "pageLayout" }],
        headerFooter: { oddFooter: "&P", firstHeader: null },
        autoFilter: "A1:C1",
        state: "hidden",
    });
}

// --- protection: what protect() keeps -----------------------------------------------------------------------------------
async function protection(): Promise<void> {
    await worksheet.protect("password", { selectLockedCells: false, spinCount: 1000 });
    const kept = worksheet.sheetProtection;
    if (kept) {
        const hash: string | undefined = kept.hashValue;
        const sheet: boolean | undefined = kept.sheet;
        void hash;
        void sheet;
    }
    worksheet.unprotect();
}

// --- page breaks: min is absent for the first column -------------------------------------------------------------------
function rowBreaks(): void {
    const breaks: ExcelJS.RowBreak[] = worksheet.rowBreaks;
    const min: Equals<ExcelJS.RowBreak["min"], number | undefined> = true;
    breaks.push({ id: 3, max: 16383, man: 1 });
    void min;
    // column breaks: id is the column number, min / max 0-based rows
    const colBreaks: ExcelJS.ColBreak[] = worksheet.colBreaks;
    const colMin: Equals<ExcelJS.ColBreak["min"], number | undefined> = true;
    colBreaks.push({ id: 3, max: 1048575, man: 1 });
    worksheet.getColumn(3).addPageBreak();
    worksheet.getColumn("D").addPageBreak(1, 10);
    // @ts-expect-error the rows are numbers
    worksheet.getColumn(3).addPageBreak("1");
    const modelBreaks: ExcelJS.ColBreak[] = worksheet.model.colBreaks;
    void colMin;
    void modelBreaks;
}

// --- images: corners as fractional col / row or as native anchors --------------------------------------------------
function images(imageId: number): void {
    worksheet.addImage(imageId, { tl: { col: 1.5, row: 1.5 }, br: { col: 3.5, row: 5.5 } });
    worksheet.addImage(imageId, { tl: { nativeCol: 1, nativeRow: 1 }, ext: { width: 100, height: 50 }, editAs: "absolute" });
    worksheet.addImage(imageId, "B2:D6");
    const [first] = worksheet.getImages();
    if (first && first.range) {
        // an Anchor of getImages() is a position
        worksheet.addImage(imageId, { tl: first.range.tl, br: first.range.tl });
    }
    // @ts-expect-error a corner needs both col and row
    worksheet.addImage(imageId, { tl: { col: 1 }, ext: { width: 1, height: 1 } });
}

// --- conditional formattings: by index, by filter, or all ------------------------------------------------------------
function conditionalFormattings(): void {
    const all: ExcelJS.ConditionalFormattingOptions[] = worksheet.conditionalFormattings;
    worksheet.removeConditionalFormatting(0);
    worksheet.removeConditionalFormatting((cf) => cf.ref !== "A1:A9");
    worksheet.removeConditionalFormatting();
    // @ts-expect-error not a filter
    worksheet.removeConditionalFormatting("A1:A9");
    void all;
}

// --- model: what worksheet.model contains --------------------------------------------------------------------------
function model(): void {
    const m = worksheet.model;
    const dimensions: Equals<typeof m.dimensions, ExcelJS.Range> = true;
    const merges: Equals<typeof m.merges, string[]> = true;
    const state: Equals<typeof m.state, ExcelJS.WorksheetState> = true;
    const views: Equals<typeof m.views, ExcelJS.WorksheetViewModel[]> = true;
    const headerFooter: Equals<typeof m.headerFooter, ExcelJS.WorksheetHeaderFooter> = true;
    const media: Equals<typeof m.media, ExcelJS.WorksheetMediumModel[]> = true;
    const rows: ExcelJS.RowModel[] = m.rows;
    const tables: ExcelJS.WorksheetTableModel[] = m.tables;
    const cols: ExcelJS.WorksheetColumnModel[] | undefined = m.cols;
    for (const medium of m.media) {
        if (medium.type === "image") {
            const col: number = medium.range.tl.nativeCol;
            void col;
        }
    }
    // a copy of a worksheet keeps its merges (read from merges when mergeCells is absent)
    workbook.addWorksheet("copy").model = { ...m, name: "copy" };
    workbook.addWorksheet("copy2").model = { ...m, name: "copy2", mergeCells: m.merges };
    void dimensions;
    void merges;
    void state;
    void views;
    void headerFooter;
    void media;
    void rows;
    void tables;
    void cols;
}

export { views, headerFooter, pageSetup, properties, addWorksheetOptions, protection, rowBreaks, images, conditionalFormattings, model };
