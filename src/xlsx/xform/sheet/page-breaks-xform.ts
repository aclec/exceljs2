import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { RowBreakModel } from "../../../types";
import colCache from "../../../utils/col-cache";
import BaseXform from "../base-xform";

/** An xsd:unsignedInt written in decimal digits (at most 10 of them: anchored, no backtracking). */
const unsignedIntRegex = /^[0-9]{1,10}$/;

/** The longest attribute value quoted in an error message. */
const MAX_QUOTED = 32;

/**
 * An unsigned integer attribute of a `<brk>` (xsd:unsignedInt), or its default: a value beyond the worksheet is its
 * last row or column (`max` is the last column, 16383, in Excel's files; 4.x wrote 16838).
 * @param last - the last row or 0-based column of a worksheet
 * @throws Error for anything but decimal digits
 */
function unsignedAttribute(node: SaxTag, name: "id" | "min" | "max", last: number): number | undefined {
    const value = node.attributes[name];
    if (value === undefined) {
        return undefined;
    }
    if (!unsignedIntRegex.test(value)) {
        throw new Error(`Invalid page break <brk ${name}=${JSON.stringify(value.slice(0, MAX_QUOTED))}>: an unsigned integer is expected`);
    }
    return Math.min(Number(value), last);
}

/**
 * The `man` attribute (xsd:boolean, default false) of a `<brk>`, as the model holds it: 1 or 0.
 * @throws Error for a value that is not an xsd:boolean
 */
function manualAttribute(node: SaxTag): number {
    const value = node.attributes.man;
    switch (value) {
        case "1":
        case "true":
            return 1;
        case undefined:
        case "0":
        case "false":
            return 0;
        default:
            throw new Error(`Invalid page break <brk man=${JSON.stringify(value.slice(0, MAX_QUOTED))}>: a boolean is expected`);
    }
}

/**
 * A page break between rows, `<brk id="5" max="16383" man="1"/>` (ECMA-376 CT_Break), rendered from a
 * {@link RowBreakModel} (its properties are the attributes). Parsing gives the CT_Break defaults for the
 * absent attributes (`id`, `max` 0; `man` 0, i.e. an automatic break; `min` absent), the `pt` attribute is not kept.
 */
class PageBreaksXform extends BaseXform<RowBreakModel> {
    get tag(): string {
        return "brk";
    }

    override render(xmlStream: XmlStream, model: RowBreakModel): void {
        // in the order of the properties of row.addPageBreak (an absent min is not written)
        xmlStream.leafNode("brk", { id: model.id, max: model.max, man: model.man, min: model.min });
    }

    /**
     * @returns `true` for `<brk>`
     * @throws Error for an attribute that is not of its xsd type (`id` above the last row, 1,048,576, and `min` / `max`
     * above the last 0-based column, 16,383, are read as these)
     */
    override parseOpen(node: SaxTag): boolean {
        if (node.name !== "brk") {
            return false;
        }
        const model: RowBreakModel = {
            id: unsignedAttribute(node, "id", colCache.MAX_ROWS) ?? 0,
            max: unsignedAttribute(node, "max", colCache.MAX_COLS - 1) ?? 0,
            man: manualAttribute(node),
        };
        const min = unsignedAttribute(node, "min", colCache.MAX_COLS - 1);
        if (min !== undefined) {
            model.min = min;
        }
        this.model = model;
        return true;
    }

    override parseText(): void {}

    override parseClose(): boolean {
        return false;
    }
}

export default PageBreaksXform;
