import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { ColorModel } from "../../../types";
import BaseXform from "../base-xform";

/**
 * A color element (CT_Color, `<color>` by default; `<fgColor>`, `<bgColor>`, `<tabColor>`... with another name):
 * `rgb`, `theme` (with `tint`), `indexed` or `auto`. A model without `argb`, `theme` and `indexed` renders `auto="1"`;
 * an `auto` color parses as `{ auto: true }`, an element without any of them as `undefined`.
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
                if (model.tint !== undefined) {
                    xmlStream.addAttribute("tint", model.tint);
                }
            } else if (model.indexed !== undefined) {
                xmlStream.addAttribute("indexed", model.indexed);
            } else {
                // `auto: true`, or a color without any (CT_Color would mean no color)
                xmlStream.addAttribute("auto", "1");
            }
            xmlStream.closeNode();
            return true;
        }
        return false;
    }

    /** @returns `true` for its element */
    override parseOpen(node: SaxTag): boolean {
        if (node.name === this.name) {
            const { rgb, theme, tint, indexed } = node.attributes;
            if (rgb) {
                this.model = { argb: rgb };
            } else if (theme) {
                const model: ColorModel = { theme: parseInt(theme, 10) };
                if (tint) {
                    model.tint = parseFloat(tint);
                }
                this.model = model;
            } else if (indexed) {
                this.model = { indexed: parseInt(indexed, 10) };
            } else if (node.attributes.auto === "1" || node.attributes.auto === "true") {
                // xsd:boolean
                this.model = { auto: true };
            } else {
                this.model = undefined;
            }
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
