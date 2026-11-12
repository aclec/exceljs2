// Internal style models: the JSON exchanged between the cell / row / column styles of src/doc and the style xforms
// (src/xlsx/xform/style). They are the public types of index.d.ts where those match the code, and extend them where the
// xforms read or write more (each difference is documented).
//
// Literal unions (font scheme, fill pattern, border style...) are the values of the API contract: a parsed model holds the
// attribute strings of the file unvalidated, so it may hold other values at runtime. Code that maps them (e.g. to XML)
// must keep a default for unknown values.

import type { Alignment, BorderStyle, FillPatterns, Font, Protection, Style } from "../../index";
import type { ValueType } from "../doc/enums";

/**
 * A color: `{ argb }`, `{ theme, tint? }` or `{ indexed }` (color-xform). `{}` renders `auto="1"`; an `auto` color parses
 * as `undefined`.
 *
 * Producer: color-xform parse, the application. Consumer: color-xform render.
 * Differs from index.d.ts `Color`, which lacks `tint` and `indexed`.
 */
export interface ColorModel {
    argb?: string;
    theme?: number;
    /** -1..1, with `theme` */
    tint?: number;
    /** legacy palette index */
    indexed?: number;
}

/**
 * A font (font-xform). Only the elements present in the file become keys when parsing.
 *
 * Differs from index.d.ts `Font`: also `condense`, `extend`, `shadow` (parsed and rendered); `size` is parsed with
 * `parseInt` (IntegerXform: `sz="10.5"` reads as 10); `underline` parses as `true` or the raw `val` string; `color` may
 * have `tint` / `indexed`.
 */
export interface FontModel extends Partial<Omit<Font, "color">> {
    color?: ColorModel;
    condense?: boolean;
    extend?: boolean;
    shadow?: boolean;
}

/** A gradient stop of a {@link FillModel} (`position` 0..1, parsed with parseFloat). */
export interface GradientStopModel {
    position: number;
    color: ColorModel;
}

/**
 * A fill (fill-xform): pattern or gradient. Rendering a fill of any other `type` writes nothing.
 *
 * Differs from index.d.ts `Fill`: a parsed gradient has no `gradient` key when it has neither `degree` nor
 * `type="path"`; `center` may have `right` / `bottom` (present only when they differ from `left` / `top`).
 */
export type FillModel =
    | {
          type: "pattern";
          pattern: FillPatterns;
          fgColor?: ColorModel;
          bgColor?: ColorModel;
      }
    | {
          type: "gradient";
          gradient?: "angle" | "path";
          /** angle gradients, integer degrees */
          degree?: number;
          /** path gradients */
          center?: { left: number; top: number; right?: number; bottom?: number };
          stops: GradientStopModel[];
      };

/** One edge of a {@link BorderModel}: an edge with a color but no style is `{ color }`. */
export interface BorderEdgeModel {
    style?: BorderStyle;
    color?: ColorModel;
}

/**
 * A border (border-xform).
 *
 * Differs from index.d.ts `Borders`: render also reads `color`, the default color of the edges; a parsed `diagonal` always
 * has `up` and `down`.
 */
export interface BorderModel {
    left?: BorderEdgeModel;
    right?: BorderEdgeModel;
    top?: BorderEdgeModel;
    bottom?: BorderEdgeModel;
    diagonal?: BorderEdgeModel & { up?: boolean; down?: boolean };
    /** render only: default color of the edges */
    color?: ColorModel;
}

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
 * with the `{ xfId: true }` option; render writes `0` for a missing id.
 */
export interface XfModel {
    numFmtId: number;
    fontId: number;
    fillId: number;
    borderId: number;
    xfId?: number;
    alignment?: AlignmentModel;
    protection?: ProtectionModel;
}

/**
 * A differential format, `<dxf>` (dxf-xform): the style of a conditional formatting rule or a table column. Parsed parts
 * that are absent are present as `undefined` keys. `numFmt` is the format code; render needs `numFmtId`, which
 * `StyleRegistry.addDxfStyle` adds (it mutates the style it is given).
 */
export interface DxfModel {
    alignment?: AlignmentModel;
    border?: BorderModel;
    fill?: FillModel;
    font?: FontModel;
    numFmt?: string;
    numFmtId?: number;
    protection?: ProtectionModel;
}

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
