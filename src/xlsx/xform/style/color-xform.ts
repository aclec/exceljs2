import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { ColorModel } from "../../../types";
import { parseBoolean } from "../../../utils/utils";
import BaseXform from "../base-xform";

/**
 * A color element (CT_Color, `<color>` by default; `<fgColor>`, `<bgColor>`, `<tabColor>`... with another name):
 * `rgb`, `theme`, `indexed` or `auto`, each with an optional `tint`. A model without `argb`, `theme` and `indexed` renders
 * `auto="1"`; an `auto` color parses as `{ auto: true }`, an element without any of them as `undefined` (its `tint` alone
 * is no color).
 */
class ColorXform extends BaseXform<ColorModel> {
    /** The name of the element. */
    readonly name: string;

    /** @param name - the name of the element (default `color`) */
    constructor(name?: string) {
        super();

        // this.name controls the xm node name
        this.name = name || "color";
    }

    get tag(): string {
        return this.name;
    }

    /** @returns `true` when the element was written (the model is truthy) */
    override render(xmlStream: XmlStream, model: ColorModel | undefined): boolean {
        if (model) {
            xmlStream.openNode(this.name);
            if (model.argb) {
                xmlStream.addAttribute("rgb", model.argb);
            } else if (model.theme !== undefined) {
                xmlStream.addAttribute("theme", model.theme);
            } else if (model.indexed !== undefined) {
                xmlStream.addAttribute("indexed", model.indexed);
            } else {
                // `auto: true`, or a color without any (CT_Color would mean no color)
                xmlStream.addAttribute("auto", "1");
            }
            // CT_Color: tint applies to any of them
            if (model.tint !== undefined) {
                xmlStream.addAttribute("tint", model.tint);
            }
            xmlStream.closeNode();
            return true;
        }
        return false;
    }

    /** @returns `true` for its element */
    override parseOpen(node: SaxTag): boolean {
        if (node.name === this.name) {
            const { rgb, theme, tint, indexed, auto } = node.attributes;
            let model: ColorModel | undefined;
            if (rgb) {
                model = { argb: rgb };
            } else if (theme) {
                model = { theme: parseInt(theme, 10) };
            } else if (indexed) {
                model = { indexed: parseInt(indexed, 10) };
            } else if (parseBoolean(auto)) {
                model = { auto: true };
            }
            if (model && tint !== undefined) {
                model.tint = parseFloat(tint);
            }
            this.model = model;
            return true;
        }
        return false;
    }

    override parseText(): void {}

    override parseClose(): boolean {
        return false;
    }
}

export default ColorXform;
