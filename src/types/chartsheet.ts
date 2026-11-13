// Internal models of the chartsheets (sheets that are a single chart: `xl/chartsheets/sheetN.xml` → its drawing → the
// chart), carried from a file read to the file written (src/xlsx/chartsheet-parts.ts): the `<chartsheet>` element as a
// tree of parsed elements, and the chart anchors of its drawing with their parts (src/xlsx/chart-parts.ts). There is no
// authoring API: a chartsheet is read, kept in its place among the sheets, and written back.

import type { WorksheetState } from "../../index";
import type { ChartAnchorModel, ChartXmlElement } from "./chart";
import type { RelationshipModel } from "./book";
import type { DrawingModel } from "./sheet";

/** A chartsheet part as parsed (src/xlsx/xform/chartsheet/chartsheet-xform.ts). */
export interface ChartsheetPartModel {
    /**
     * The `<chartsheet>` element, without the elements that reference a part that is not carried (`<legacyDrawing>`,
     * `<legacyDrawingHF>`, `<drawingHF>`, `<picture>`, any other element with a relationship id) and without the
     * relationship id of `<pageSetup>` (printer settings): its only attribute marked `rel` is the `r:id` of `<drawing>`.
     */
    readonly sheet: ChartXmlElement;
    /** The relationship id of its `<drawing>` (the id in the file). */
    readonly drawingRelId: string;
}

/** A chartsheet part read, with the name of its drawing (`drawing1`) when its relationship resolves to a drawing part. */
export interface LoadedChartsheetModel extends ChartsheetPartModel {
    readonly drawing: string;
}

/**
 * A chartsheet of a workbook model (`WorkbookModel.chartsheets`): its `<sheet>` (name, `sheetId`, state), its place among
 * all the sheets, its part and its charts.
 */
export interface ChartsheetModel extends ChartsheetPartModel {
    /** The `sheetId` (a new one is given on write when a worksheet has the same). */
    id: number;
    name: string;
    /** Absent: visible. */
    state?: WorksheetState;
    /**
     * The 0-based position of the chartsheet among all the sheets written (worksheets and chartsheets): the chartsheets
     * are inserted, by increasing index, into the list of the worksheets (an index beyond the end appends).
     */
    index: number;
    /** The charts of its drawing (one, usually), as a worksheet carries its charts. */
    readonly charts: readonly ChartAnchorModel[];
}

/** A chartsheet being written (xlsx.ts prepare): its `<sheet>`, part, relationships and drawing. */
export interface ChartsheetWriteModel extends ChartsheetPartModel {
    /** The `sheetId` written (unique among the sheets written). */
    id: number;
    name: string;
    state?: WorksheetState;
    /** Its position among the sheets written ({@link ChartsheetModel.index}). */
    readonly index: number;
    /** The relationship id of the workbook part: set by xlsx.ts addWorkbookRels. */
    rId?: string;
    /** The number of the part: `xl/chartsheets/sheet<partNo>.xml`. */
    readonly partNo: number;
    /** The relationship id written for its `<drawing>`. */
    readonly drawingRId: string;
    /** The charts of its drawing. */
    readonly charts: readonly ChartAnchorModel[];
    /** The relationships of the part (its drawing): set with {@link drawing}. */
    rels: RelationshipModel[];
    /** The drawing written, holding the chart anchors (chartsheet-parts.ts prepareChartsheetDrawings). */
    drawing?: DrawingModel;
}
