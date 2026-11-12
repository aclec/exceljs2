import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { Alignment } from "../../../../index";
import type { AlignmentModel } from "../../../types";
import { ReadingOrder } from "../../../doc/enums";
import { parseBoolean, validInt } from "../../../utils/utils";
import BaseXform, { isOneOf } from "../base-xform";

// ECMA-376 ST_HorizontalAlignment, but general (the default: no alignment)
const HORIZONTAL_VALUES: Readonly<Record<Alignment["horizontal"], true>> = {
    left: true,
    center: true,
    right: true,
    fill: true,
    centerContinuous: true,
    distributed: true,
    justify: true,
};
// ECMA-376 ST_VerticalAlignment, with middle for center
const VERTICAL_VALUES: Readonly<Record<Alignment["vertical"], true>> = {
    top: true,
    middle: true,
    bottom: true,
    distributed: true,
    justify: true,
};

// model value -> attribute value (undefined: not written)
const validation = {
    horizontal(value: string | undefined): string | undefined {
        return isOneOf(HORIZONTAL_VALUES, value) ? value : undefined;
    },
    vertical(value: string | undefined): string | undefined {
        if (value === "middle") return "center";
        return isOneOf(VERTICAL_VALUES, value) ? value : undefined;
    },
    textRotation(value: number | string | undefined): number | "vertical" | undefined {
        if (value === "vertical") {
            return value;
        }
        const rotation = validInt(value);
        return rotation >= -90 && rotation <= 90 ? rotation : undefined;
    },
    indent(value: number | undefined): number {
        return Math.max(0, validInt(value));
    },
    readingOrder(value: string | undefined): ReadingOrder | undefined {
        if (value === "ltr") return ReadingOrder.LeftToRight;
        if (value === "rtl") return ReadingOrder.RightToLeft;
        return undefined;
    },
};

// textRotation: model (-90..90 degrees or "vertical") <-> XML (0..90 up, 91..180 down, 255 vertical)
const textRotationXform = {
    toXml(value: number | "vertical" | undefined): number | undefined {
        const textRotation = validation.textRotation(value);
        if (textRotation) {
            if (textRotation === "vertical") {
                return 255;
            }

            const tr = Math.round(textRotation);
            if (tr >= 0 && tr <= 90) {
                return tr;
            }

            if (tr < 0 && tr >= -90) {
                return 90 - tr;
            }
        }
        return undefined;
    },
    toModel(textRotation: string): number | "vertical" | undefined {
        const tr = validInt(textRotation);
        if (tr === 255) {
            return "vertical";
        }
        if (tr >= 0 && tr <= 90) {
            return tr;
        }
        if (tr > 90 && tr <= 180) {
            return 90 - tr;
        }
        return undefined;
    },
};

/**
 * The alignment of a cell format, `<alignment>`. Invalid values are not written, and nothing is written when no value is
 * valid. The XML `center` vertical alignment is `"middle"` in the model. Parsing gives `null` when no attribute is set;
 * a `horizontal` / `vertical` that is not one of the model (`general`, the default, or an invalid value) is absent; the
 * other values are not validated (except `textRotation`, which may parse as `undefined`).
 */
class AlignmentXform extends BaseXform<AlignmentModel> {
    get tag(): string {
        return "alignment";
    }

    override render(xmlStream: XmlStream, model: AlignmentModel): void {
        // the attributes, in this order, when truthy
        const horizontal = validation.horizontal(model.horizontal);
        const vertical = validation.vertical(model.vertical);
        const wrapText = model.wrapText ? "1" : undefined;
        const shrinkToFit = model.shrinkToFit ? "1" : undefined;
        const indent = validation.indent(model.indent) || undefined;
        const textRotation = textRotationXform.toXml(model.textRotation) || undefined;
        const readingOrder = validation.readingOrder(model.readingOrder);

        if (horizontal || vertical || wrapText || shrinkToFit || indent || textRotation || readingOrder) {
            xmlStream.leafNode("alignment", {
                horizontal,
                vertical,
                wrapText,
                shrinkToFit,
                indent,
                textRotation,
                readingOrder,
            });
        }
    }

    /** Reads `<alignment>` (returns nothing: the parent decides when to delegate to this xform). */
    override parseOpen(node: SaxTag): void {
        const { horizontal, vertical, wrapText, shrinkToFit, indent, textRotation, readingOrder } = node.attributes;
        const model: AlignmentModel = {};

        let valid = false;
        // a value that is not one of the model (general, an invalid one) is absent
        if (isOneOf(HORIZONTAL_VALUES, horizontal)) {
            model.horizontal = horizontal;
            valid = true;
        }
        const verticalValue = vertical === "center" ? "middle" : vertical;
        if (isOneOf(VERTICAL_VALUES, verticalValue)) {
            model.vertical = verticalValue;
            valid = true;
        }
        // "0"/"false" is the default: treat it as absent, as rendering does
        if (parseBoolean(wrapText)) {
            model.wrapText = true;
            valid = true;
        }
        if (parseBoolean(shrinkToFit)) {
            model.shrinkToFit = true;
            valid = true;
        }
        if (indent) {
            model.indent = parseInt(indent, 10);
            valid = true;
        }
        if (textRotation) {
            model.textRotation = textRotationXform.toModel(textRotation);
            valid = true;
        }
        if (readingOrder) {
            model.readingOrder = readingOrder === "2" ? "rtl" : "ltr";
            valid = true;
        }

        this.model = valid ? model : null;
    }

    override parseText(): void {}

    override parseClose(): boolean {
        return false;
    }
}

export default AlignmentXform;
