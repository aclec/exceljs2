import colCache from "./col-cache";

// Every cell of a merge is materialised (300 to 600 bytes of heap each), so that a few bytes of XML such as
// <mergeCell ref="A1:XFD1048576"/> cannot exhaust the heap. The same caps apply when a workbook is built and when it is
// read, so that a workbook ExcelJS writes with the defaults can be read back with the defaults:
//  - one worksheet: one full column (1,048,576 cells), not configurable;
//  - one workbook: workbook.maxMergedCells (load option maxMergedCells, WorkbookWriter option maxMergedCells),
//    one full worksheet by default (about 600 MB of heap); trusted callers raise it.

/** The maximum number of merged cells of a worksheet: one full column (1,048,576), not configurable. */
const MAX_MERGED_CELLS_PER_SHEET = colCache.MAX_ROWS;
/** The default maximum number of merged cells of a workbook (option maxMergedCells). */
const DEFAULT_MAX_MERGED_CELLS = MAX_MERGED_CELLS_PER_SHEET;

/**
 * Validates the maxMergedCells option.
 * @returns value, a number >= 0 (Infinity allowed).
 * @throws TypeError for anything else (NaN included).
 */
function checkMaxMergedCells(value: unknown): number {
    if (typeof value !== "number" || !(value >= 0)) {
        throw new TypeError(`maxMergedCells must be a number >= 0 (or Infinity), got ${String(value)}`);
    }
    return value;
}

/**
 * Checks the merged cell counts that a worksheet and its workbook would reach.
 * @param what - The operation, for the error message.
 * @param sheetCount - The merged cells the worksheet would hold.
 * @param workbookCount - The merged cells the workbook would hold.
 * @param maxMergedCells - The workbook limit (option maxMergedCells).
 * @throws Error when sheetCount exceeds {@link MAX_MERGED_CELLS_PER_SHEET} or workbookCount exceeds maxMergedCells.
 */
function checkMergedCells(what: string, sheetCount: number, workbookCount: number, maxMergedCells: number): void {
    if (sheetCount > MAX_MERGED_CELLS_PER_SHEET) {
        throw new Error(`${what}: worksheet would exceed ${MAX_MERGED_CELLS_PER_SHEET} merged cells`);
    }
    if (workbookCount > maxMergedCells) {
        throw new Error(`${what}: the workbook would exceed ${maxMergedCells} merged cells (see the maxMergedCells option)`);
    }
}

export { MAX_MERGED_CELLS_PER_SHEET, DEFAULT_MAX_MERGED_CELLS, checkMaxMergedCells, checkMergedCells };
