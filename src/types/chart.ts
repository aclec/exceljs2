// Internal models of the charts of a worksheet, carried opaque from a file read to the file written (src/xlsx/chart-parts.ts,
// drawing/chart-anchor.ts): the anchor of each chart in the drawing of its worksheet, as a tree of parsed elements, and
// the parts of the chart (chart, style, colors, embedded workbook...) as their bytes, with the graph of their relationships.
// A chart created by `worksheet.addChart` (src/doc/chart.ts) is turned into the same models when the workbook is written:
// its chart part generated from a ChartSpaceSpec (src/xlsx/chart-space.ts), the cached values read from the cells.

import type { ChartGrouping, ChartLegendPosition, ChartType, ImageEditAs } from "../../index";
import type { AnchorModel } from "./sheet";

/** An attribute of an element of a chart anchor, as parsed (name and decoded value from the file). */
export interface ChartXmlAttribute {
    readonly name: string;
    readonly value: string;
    /**
     * `rel`: a relationship id of the drawing (an attribute of the relationships namespace), replaced on write by the id of
     * the relationship written; `shapeId`: the `id` of an `<xdr:cNvPr>`, renumbered on write (unique in the drawing).
     */
    readonly role?: "rel" | "shapeId";
}

/** An element of a chart anchor, as parsed: its children are elements and text (decoded). */
export interface ChartXmlElement {
    readonly name: string;
    readonly attributes: readonly ChartXmlAttribute[];
    readonly children: readonly (ChartXmlElement | string)[];
}

/** The kinds of parts carried with a chart: each has its relationship type, part name and content type (chart-parts.ts). */
export type ChartPartKind =
    "chart" | "chartEx" | "chartStyle" | "chartColorStyle" | "chartUserShapes" | "themeOverride" | "package" | "oleObject" | "image";

/** A part of a chart, carried as its bytes (never parsed), with its relationships. */
export interface ChartPartModel {
    readonly kind: ChartPartKind;
    /** The extension of the part name written (from the file for an embedded package, an OLE object or an image). */
    readonly extension: string;
    readonly data: Uint8Array;
    /** Its relationships (ids as in the file, which the bytes reference); a graph: parts may be shared, or form a cycle. */
    readonly rels: ChartRelModel[];
}

/** A relationship of a chart part or of a chart anchor: to a carried part, or to an external target (kept as is). */
export interface ChartRelModel {
    /** The id in the file (`r:id` of the anchor, or referenced by the bytes of the source part). */
    readonly Id: string;
    readonly Type: string;
    /** The part targeted (an internal relationship). */
    readonly part?: ChartPartModel;
    /** The target of an external relationship (`TargetMode="External"`), as in the file. */
    readonly externalTarget?: string;
}

/** A chart of a worksheet: its anchor in the drawing (two-cell, one-cell or absolute) and the relationships it uses. */
export interface ChartAnchorModel {
    /** The anchor element, the namespaces in scope in the file declared on it. */
    readonly anchor: ChartXmlElement;
    /** The relationships of the drawing that the anchor references (by their id in the file). */
    readonly rels: readonly ChartRelModel[];
}

/** A chart anchor of a drawing being written: the anchor and the new relationship id of each id of the file. */
export interface DrawingChartModel {
    readonly anchor: ChartXmlElement;
    readonly ids: ReadonlyMap<string, string>;
}

/** Names the chart parts while a workbook is written (chart-parts.ts ChartPartWriter). */
export interface ChartPartRegistry {
    /** @returns the part name of a chart part in the package written (`xl/charts/chart1.xml`), assigned on first call */
    partName(part: ChartPartModel): string;
    /** Claims a chart part for one anchor of the workbook written. @returns `false` when another anchor claimed it already */
    claim(part: ChartPartModel): boolean;
}

/** A part to declare in `[Content_Types].xml` (an `<Override>`). */
export interface ContentTypeOverride {
    /** The part name, without the leading `/`. */
    readonly name: string;
    readonly contentType: string;
}

/** A point of the cached values of a chart data reference: its 0-based index in the range, and its value. */
export interface ChartPoint<T> {
    readonly index: number;
    readonly value: T;
}

/** The cached numbers of a chart data reference (`<c:numCache>`): the empty and non-numeric cells have no point. */
export interface ChartNumberCache {
    readonly type: "number";
    /** The number format of the values (`<c:formatCode>`), `General` by default. */
    readonly formatCode: string;
    /** The number of cells of the range (`<c:ptCount>`). */
    readonly count: number;
    /** By increasing index (iterable more than once). */
    readonly points: Iterable<ChartPoint<number>>;
}

/** The cached texts of a chart data reference (`<c:strCache>`): the empty cells have no point. */
export interface ChartStringCache {
    readonly type: "string";
    /** The number of cells of the range (`<c:ptCount>`). */
    readonly count: number;
    /** By increasing index (iterable more than once). */
    readonly points: Iterable<ChartPoint<string>>;
}

/**
 * A chart data reference resolved when the workbook is written: its formula (`'My Sheet'!$B$2:$B$5`) and its cache;
 * without a cache when the values are not known (the streaming writer: rows committed before the chart was added), the
 * chart then shows the values of the cells once recalculated.
 */
export interface ChartDataRef<Cache extends ChartNumberCache | ChartStringCache> {
    readonly formula: string;
    readonly cache?: Cache | undefined;
}

/** A series of a chart to generate (scatter: `categories` are the x values, `values` the y values). */
export interface ChartSeriesSpec {
    /** The name: a text (`<c:v>`), or a reference to a cell. */
    readonly name?: string | ChartDataRef<ChartStringCache> | undefined;
    readonly categories?: ChartDataRef<ChartNumberCache | ChartStringCache> | undefined;
    readonly values: ChartDataRef<ChartNumberCache>;
    /** The color, `RRGGBB` (upper case hexadecimal digits). */
    readonly color?: string | undefined;
}

/** An axis of a chart to generate. */
export interface ChartAxisSpec {
    readonly title?: string | undefined;
    /** The bounds of a value axis. */
    readonly min?: number | undefined;
    readonly max?: number | undefined;
    /** The number format of the labels (otherwise the format of the source cells). */
    readonly numFmt?: string | undefined;
    readonly majorGridlines: boolean;
    /** `false`: the axis is deleted (`<c:delete val="1"/>`). */
    readonly visible: boolean;
}

/**
 * A chart to generate as a `<c:chartSpace>` part (src/xlsx/chart-space.ts renderChartSpace): validated (see
 * src/doc/chart.ts), its data references resolved, with their cached values.
 */
export interface ChartSpaceSpec {
    readonly type: ChartType;
    /** bar / column / line / area: the grouping of the series (default: clustered for bar / column, standard otherwise). */
    readonly grouping?: ChartGrouping | undefined;
    /** line / scatter: whether the points have markers. */
    readonly marker: boolean;
    /** scatter: whether the points are joined by lines. */
    readonly line: boolean;
    readonly title?: string | undefined;
    /** The position of the legend; `undefined`: no legend. */
    readonly legend?: ChartLegendPosition | undefined;
    /** The date system of the workbook (`<c:date1904>`). */
    readonly date1904: boolean;
    readonly series: readonly ChartSeriesSpec[];
    /** The category axis (scatter: the x value axis). Ignored for pie and doughnut charts. */
    readonly categoryAxis: ChartAxisSpec;
    /** The value axis (scatter: the y axis). Ignored for pie and doughnut charts. */
    readonly valueAxis: ChartAxisSpec;
}

/** Where a chart is anchored in the drawing of its worksheet: two corners, or a corner and a size. */
export interface ChartAnchorSpec {
    readonly from: AnchorModel;
    /** The bottom-right corner (`<xdr:twoCellAnchor>`). */
    readonly to?: AnchorModel | undefined;
    /** The size in pixels (`<xdr:oneCellAnchor>`, without `to`). */
    readonly ext?: { readonly width: number; readonly height: number } | undefined;
    /** How a two-cell anchor moves with the cells (no attribute: `twoCell`, moved and sized with them). */
    readonly editAs?: ImageEditAs | undefined;
    /** The name of the graphic frame (`<xdr:cNvPr name>`), e.g. `Chart 1`. */
    readonly name: string;
}

/** Where the chart of a chartsheet is in its drawing: a position and a size in EMU (`<xdr:absoluteAnchor>`). */
export interface ChartAbsoluteAnchorSpec {
    readonly absolute: { readonly x: number; readonly y: number; readonly cx: number; readonly cy: number };
    /** The name of the graphic frame (`<xdr:cNvPr name>`), e.g. `Chart 1`. */
    readonly name: string;
}
