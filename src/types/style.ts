// Internal style models: the JSON exchanged between the cell / row / column styles of src/doc and the style xforms
// (src/xlsx/xform/style). They are the public types of index.d.ts where those match the code, and extend them where the
// xforms read or write more (each difference is documented).
//
// Literal unions (font scheme, fill pattern, border style...) are the values of the API contract: a parsed model holds the
// attribute strings of the file unvalidated, so it may hold other values at runtime. Code that maps them (e.g. to XML)
// must keep a default for unknown values.

import type {
    Alignment,
    Border,
    Borders,
    Color,
    DifferentialStyle,
    FillGradientAngle,
    FillGradientPath,
    FillPattern,
    FillPatterns,
    Font,
    GradientStop,
    Protection,
    Style,
} from "../../index";
import type { ValueType } from "../doc/enums";

/**
 * A color (color-xform): index.d.ts `Color`, all parts optional: `{ argb }`, `{ theme }`, `{ indexed }` or `{ auto: true }`,
 * each with an optional `tint`. A color without `argb`, `theme` and `indexed` (e.g. `{}`) renders `auto="1"`; an element
 * without any color attribute parses as `undefined`.
 *
 * Producer: color-xform parse, the application. Consumer: color-xform render.
 */
export type ColorModel = Partial<Color>;

/**
 * A font (font-xform): index.d.ts `Font`, all parts optional. Only the elements present in the file become keys when
 * parsing; `size` is parsed with `parseFloat` (`sz="10.5"` reads as 10.5); `underline` parses as `true` or
 * the raw `val` string.
 */
export type FontModel = Partial<Font>;

/**
 * A gradient stop of a {@link FillModel} (index.d.ts `GradientStop`: `position` 0..1, parsed with parseFloat). A parsed
 * stop has no `color` without a `<color>` child, and `color: undefined` for a `<color>` without color attribute; a stop
 * without color renders `<color/>` (CT_GradientStop requires the element).
 */
export type GradientStopModel = GradientStop;

/**
 * A fill (fill-xform): pattern or gradient, index.d.ts `Fill` but for the pattern, which a parsed `<patternFill>` lacks
 * without `patternType` (Excel writes the pattern fills of the dxfs so). Rendering a fill of any other `type` writes
 * nothing. A parsed angle gradient always has `degree`; a parsed gradient has no `gradient` key when it has neither
 * `degree` nor `type="path"`; a parsed path gradient always has `center`, whose `right` / `bottom` are present only when
 * they differ from `left` / `top`.
 */
export type FillModel = (Omit<FillPattern, "pattern"> & { pattern?: FillPatterns }) | FillGradientAngle | FillGradientPath;

/** One edge of a {@link BorderModel} (index.d.ts `Border`, all parts optional): an edge with a color but no style is `{ color }`. */
export type BorderEdgeModel = Partial<Border>;

/**
 * A border (border-xform): index.d.ts `Borders`, all parts optional. Render also reads `color`, the default color of the
 * edges (never parsed); a parsed `diagonal` always has `up` and `down`.
 */
export type BorderModel = Partial<Borders>;

/**
 * An alignment (alignment-xform): index.d.ts `Alignment`. The XML `center` vertical alignment is `"middle"` in the
 * model; `readingOrder` "2" parses as `"rtl"`, anything else as `"ltr"`. Parsing gives `null` when nothing is set.
 */
export type AlignmentModel = Partial<Alignment>;

/** Protection (protection-xform): index.d.ts `Protection`. Parsing gives `null` when both are the defaults. */
export type ProtectionModel = Partial<Protection>;

/**
 * The style of a cell, row or column as the document holds it (`cell.style`): index.d.ts `Style`, all parts optional.
 *
 * Producer: src/doc (cell/row/column `style`), `StyleLookup.getStyleModel` on read. Consumer: `StyleRegistry.addStyleModel`
 * on write, the document setters on read. On read, `font` / `border` / `fill` are shared between the cells of a style
 * (references to the style manager's model): copy before mutating.
 */
export type StyleModel = Partial<Style>;

/** A number format (numfmt-xform): `formatCode` is `undefined` when longer than 255 characters (rejected). */
export interface NumFmtModel {
    id: number;
    formatCode: string | undefined;
}

/**
 * A cell format, `<xf>` of `cellXfs` / `cellStyleXfs` (style-xform): indexes into the style manager's lists. `xfId` only
 * with the `{ xfId: true }` option; render writes `0` for a missing id (the style manager registers formats with only
 * the ids that are set). Parsing sets every id (`NaN` when absent or not a number); a parsed `alignment` / `protection`
 * is `null` when its element sets only defaults.
 */
export interface XfModel {
    numFmtId?: number;
    fontId?: number;
    fillId?: number;
    borderId?: number;
    xfId?: number;
    alignment?: AlignmentModel | null;
    protection?: ProtectionModel | null;
}

/**
 * A differential format, `<dxf>` (dxf-xform): index.d.ts `DifferentialStyle`, the style of a conditional formatting
 * rule or a table column. Parsing gives every part as a key: `null` when its element is absent (or sets only defaults);
 * `fill` may be `{}` (a `<fill>` without a known child) and `numFmt` `undefined` (a rejected format code). `numFmt` is
 * the format code; render needs `numFmtId`, which `StyleRegistry.addDxfStyle` adds (it mutates the style it is given).
 */
export type DxfModel = DifferentialStyle;

/**
 * The write side of the style manager (StylesXform, or StylesXform.Mock when styles are not written): registers styles
 * and returns their ids. Passed as `options.styles` to `prepare()`.
 */
export interface StyleRegistry {
    /**
     * @param cellType - the type of the cell (Number and Date cells get a default number format)
     * @returns the id of the cell format (`s` attribute): 0 for a falsy style
     */
    addStyleModel(style: StyleModel, cellType?: ValueType): number;
    /** @returns the index of the dxf; mutates `style` (adds `numFmtId` when it has a `numFmt`) */
    addDxfStyle(style: DxfModel): number;
}

/** The read side of the style manager (a parsed StylesXform). Passed as `options.styles` to `reconcile()`. */
export interface StyleLookup {
    /** @returns the style of a cell format id (cached, shared between callers), or `null` for an unknown id */
    getStyleModel(id: number): StyleModel | null;
    /** @returns the dxf at an index (`undefined` when there is none) */
    getDxfStyle(id: number): DxfModel | undefined;
}
