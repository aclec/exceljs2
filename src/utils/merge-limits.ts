/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import colCache from "./col-cache";

// Every cell of a merge is materialised (300 to 600 bytes of heap each), so that a few bytes of XML such as
// <mergeCell ref="A1:XFD1048576"/> cannot exhaust the heap. The same caps apply when a workbook is built and when it is
// read, so that a workbook ExcelJS writes with the defaults can be read back with the defaults:
//  - one worksheet: one full column (1,048,576 cells), not configurable;
//  - one workbook: workbook.maxMergedCells (load option maxMergedCells, WorkbookWriter option maxMergedCells),
//    one full worksheet by default (about 600 MB of heap); trusted callers raise it.
const MAX_MERGED_CELLS_PER_SHEET = colCache.MAX_ROWS;
const DEFAULT_MAX_MERGED_CELLS = MAX_MERGED_CELLS_PER_SHEET;

function checkMaxMergedCells(value) {
    if (typeof value !== "number" || !(value >= 0)) {
        throw new TypeError(`maxMergedCells must be a number >= 0 (or Infinity), got ${value}`);
    }
    return value;
}

// throws when a worksheet would hold sheetCount merged cells, and the workbook workbookCount of them
function checkMergedCells(what, sheetCount, workbookCount, maxMergedCells) {
    if (sheetCount > MAX_MERGED_CELLS_PER_SHEET) {
        throw new Error(`${what}: worksheet would exceed ${MAX_MERGED_CELLS_PER_SHEET} merged cells`);
    }
    if (workbookCount > maxMergedCells) {
        throw new Error(`${what}: the workbook would exceed ${maxMergedCells} merged cells (see the maxMergedCells option)`);
    }
}

export { MAX_MERGED_CELLS_PER_SHEET, DEFAULT_MAX_MERGED_CELLS, checkMaxMergedCells, checkMergedCells };
