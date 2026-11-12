import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { XmlAttributes } from "../../../utils/xml-stream";
import BaseXform, { type ValueXformOptions } from "../base-xform";

/** The options of an {@link IntegerXform}. */
export interface IntegerXformOptions extends ValueXformOptions {
    /** Render the element also when the model is 0 (or `undefined`: then it must have an attribute). */
    readonly zero?: boolean;
}

/**
 * An element holding an integer, in an attribute or as its text. Unlike FloatXform, zero is not rendered (unless the
 * `zero` option is set). Parsed with `parseInt(_, 10)`: `NaN` when the attribute is absent or not a number; an empty text
 * is 0.
 */
class IntegerXform extends BaseXform<number> {
    /** The element. */
    readonly tag: string;
    /** The attribute holding the value; `undefined`: the text of the element. */
    readonly attr: string | undefined;
    /** Fixed attributes written on the element. */
    readonly attrs: XmlAttributes | undefined;
    /** Render zero. */
    readonly zero: boolean | undefined;
    /** The chunks of text of the element being parsed (text mode). */
    protected text: string[] = [];

    constructor(options: IntegerXformOptions) {
        super();

        this.tag = options.tag;
        this.attr = options.attr;
        this.attrs = options.attrs;

        // option to render zero
        this.zero = options.zero;
    }

    /** @throws Error in text mode with the `zero` option when the model is `undefined` (the callers always pass a number) */
    override render(xmlStream: XmlStream, model: number | undefined): void {
        // int is different to float in that zero is not rendered
        if (model || this.zero) {
            const { attr } = this;
            if (!attr && model === undefined) {
                throw new Error(`No integer to write in <${this.tag}>`);
            }
            xmlStream.openNode(this.tag);
            if (this.attrs) {
                xmlStream.addAttributes(this.attrs);
            }
            if (attr) {
                xmlStream.addAttribute(attr, model);
            } else if (model !== undefined) {
                xmlStream.writeText(model);
            }
            xmlStream.closeNode();
        }
    }

    /** @returns `true` for its element */
    override parseOpen(node: SaxTag): boolean {
        if (node.name === this.tag) {
            if (this.attr) {
                this.model = BaseXform.toIntValue(node.attributes[this.attr], NaN);
            } else {
                this.text = [];
            }
            return true;
        }
        return false;
    }

    override parseText(text: string): void {
        if (!this.attr) {
            this.text.push(text);
        }
    }

    override parseClose(): boolean {
        if (!this.attr) {
            const text = this.text.join("");
            this.model = text ? parseInt(text, 10) : 0;
        }
        return false;
    }
}

export default IntegerXform;
