import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import utils from "../../../utils/utils";
import BaseXform from "../base-xform";

/** The options of a {@link BooleanXform}. */
export interface BooleanXformOptions {
    /** The element, e.g. `b`. */
    readonly tag: string;
    /** The attribute holding the value, e.g. `val` (read when parsing; the element is rendered without it). */
    readonly attr: string;
}

/**
 * A boolean property element (ECMA-376 CT_BooleanProperty), e.g. `<b/>`: rendered as the empty element when the model is
 * truthy, nothing otherwise. Parsing: `<b/>` is `true`; `<b val="0"/>` or `<b val="false"/>` is `false`.
 */
class BooleanXform extends BaseXform<boolean> {
    /** The element. */
    readonly tag: string;
    /** The attribute read when parsing. */
    readonly attr: string;

    constructor(options: BooleanXformOptions) {
        super();

        this.tag = options.tag;
        this.attr = options.attr;
    }

    /** Writes the empty element when `model` is truthy. */
    override render(xmlStream: XmlStream, model: boolean | undefined): void {
        if (model) {
            xmlStream.openNode(this.tag);
            xmlStream.closeNode();
        }
    }

    /** Reads the value of its element (returns nothing: the parent decides when to delegate to this xform). */
    override parseOpen(node: SaxTag): void {
        if (node.name === this.tag) {
            // CT_BooleanProperty: <b/> is true, <b val="0"/> or <b val="false"/> is false
            this.model = utils.parseBoolean(node.attributes[this.attr], true);
        }
    }

    override parseText(): void {}

    override parseClose(): boolean {
        return false;
    }
}

export default BooleanXform;
