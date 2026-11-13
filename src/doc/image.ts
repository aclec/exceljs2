// An image of a worksheet: the background image, or an image placed over a range of cells.

import type { ImageEditAs } from "../../index";
import colCache from "../utils/col-cache";
import type { ImageHyperlinkModel, WorksheetMediumModel } from "../types";
import Anchor, { type AnchorPosition, type AnchorWorksheet } from "./anchor";

/** The position of an image as given to `worksheet.addImage` (or read from a model): corners and size. */
export interface ImageRangeInput {
    /** The top-left corner. */
    tl: AnchorPosition;
    /** The bottom-right corner (two-cell anchor). */
    br?: AnchorPosition | undefined;
    /** The size in pixels (one-cell anchor, with `tl`). */
    ext?: { width: number; height: number } | undefined;
    editAs?: ImageEditAs | undefined;
    hyperlinks?: ImageHyperlinkModel | undefined;
}

/**
 * What the {@link Image} model setter takes: a {@link WorksheetMediumModel}, or for an image a range given as a string
 * (`"B2:D6"`: the image covers the cells, `editAs` "oneCell") or as an {@link ImageRangeInput}. `imageId` is the index
 * of the image in the workbook media (`workbook.addImage`).
 */
export type ImageInput =
    | { type: "background"; imageId: number | string }
    | { type: "image"; imageId: number | string; range: string | ImageRangeInput; hyperlinks?: ImageHyperlinkModel | undefined };

/** The position of an image in the worksheet, as an {@link Image} holds it. */
export interface ImageRange {
    tl: Anchor;
    br?: Anchor | undefined;
    ext?: { width: number; height: number } | undefined;
    editAs?: ImageEditAs | undefined;
    hyperlinks?: ImageHyperlinkModel | undefined;
}

/** An image of a worksheet (`worksheet.getImages()`), or its background image. */
class Image {
    /** The worksheet, which gives the sizes of the cells to the anchors. */
    worksheet: AnchorWorksheet;
    /** `"image"` or `"background"`: set by the model setter. */
    declare type: "image" | "background";
    /** The index of the image in the workbook media: set by the model setter. */
    declare imageId: number | string;
    /** The position of an image (set by the model setter for the type `"image"` only). */
    declare range?: ImageRange;

    constructor(worksheet: AnchorWorksheet, model: ImageInput) {
        this.worksheet = worksheet;
        this.model = model;
    }

    /**
     * The model of the image (new objects; `ext` and `hyperlinks` are shared).
     * @throws Error when the type is neither "image" nor "background", or for an image without range
     */
    get model(): WorksheetMediumModel {
        switch (this.type) {
            case "background":
                return {
                    type: this.type,
                    imageId: this.imageId,
                };
            case "image": {
                // the model setter sets the range of an image (getImages() gives it: it may be removed)
                const { range } = this;
                if (!range) {
                    throw new Error(`The image ${this.imageId} has no range`);
                }
                return {
                    type: this.type,
                    imageId: this.imageId,
                    hyperlinks: range.hyperlinks,
                    range: {
                        tl: range.tl.model,
                        br: range.br && range.br.model,
                        ext: range.ext,
                        editAs: range.editAs,
                    },
                };
            }
            default:
                throw new Error("Invalid Image Type");
        }
    }

    set model(input: ImageInput) {
        this.type = input.type;
        this.imageId = input.imageId;

        if (input.type === "image") {
            const { range } = input;
            if (typeof range === "string") {
                const decoded = colCache.decode(range);
                // a single cell ("B2") is not a range: both anchors are then at A1, as they always were
                const corners = "top" in decoded ? decoded : undefined;
                this.range = {
                    tl: new Anchor(this.worksheet, corners && { col: corners.left, row: corners.top }, -1),
                    br: new Anchor(this.worksheet, corners && { col: corners.right, row: corners.bottom }, 0),
                    editAs: "oneCell",
                };
            } else {
                this.range = {
                    tl: new Anchor(this.worksheet, range.tl, 0),
                    br: range.br && new Anchor(this.worksheet, range.br, 0),
                    ext: range.ext,
                    editAs: range.editAs,
                    hyperlinks: input.hyperlinks || range.hyperlinks,
                };
            }
        }
    }
}

export default Image;
