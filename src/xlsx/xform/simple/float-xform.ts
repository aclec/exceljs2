import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { XmlAttributes } from "../../../utils/xml-stream";
import BaseXform, { type ValueXformOptions } from "../base-xform";

/**
 * An element holding a number, in an attribute or as its text. Rendered unless the model is `undefined` (`0` is
 * written). Parsed with `parseFloat`: `NaN` when the value is absent or not a number.
 */
class FloatXform extends BaseXform<number> {
    /** The element. */
    readonly tag: string;
    /** The attribute holding the value; `undefined`: the text of the element. */
    readonly attr: string | undefined;
    /** Fixed attributes written on the element. */
    readonly attrs: XmlAttributes | undefined;
    /** The chunks of text of the element being parsed (text mode). */
    protected text: string[] = [];

    constructor(options: ValueXformOptions) {
        super();

        this.tag = options.tag;
        this.attr = options.attr;
        this.attrs = options.attrs;
    }

    override render(xmlStream: XmlStream, model: number | undefined): void {
        if (model !== undefined) {
            xmlStream.openNode(this.tag);
            if (this.attrs) {
                xmlStream.addAttributes(this.attrs);
            }
            if (this.attr) {
                xmlStream.addAttribute(this.attr, model);
            } else {
                xmlStream.writeText(model);
            }
            xmlStream.closeNode();
        }
    }

    /** Starts reading its element (attribute mode: the model is the parsed attribute). */
    override parseOpen(node: SaxTag): void {
        if (node.name === this.tag) {
            if (this.attr) {
                this.model = BaseXform.toFloatValue(node.attributes[this.attr], NaN);
            } else {
                this.text = [];
            }
        }
    }

    override parseText(text: string): void {
        if (!this.attr) {
            this.text.push(text);
        }
    }

    override parseClose(): boolean {
        if (!this.attr) {
            this.model = parseFloat(this.text.join(""));
        }
        return false;
    }
}

export default FloatXform;
