import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { XmlAttributes } from "../../../utils/xml-stream";
import type { FontModel } from "../../../types";
import BaseXform, { isOneOf } from "../base-xform";

/** The underline of a font: `true` (single) or an underline style. */
export type UnderlineModel = NonNullable<FontModel["underline"]>;

// ECMA-376 ST_UnderlineValues
const UNDERLINE_VALUES: Readonly<Record<Exclude<UnderlineModel, boolean>, true>> = {
    none: true,
    single: true,
    double: true,
    singleAccounting: true,
    doubleAccounting: true,
};

/** The underline styles rendered, with the attributes of their `<u>` element (`single` is the default: no `val`). */
const ATTRIBUTES: Readonly<Record<string, XmlAttributes>> = {
    single: {},
    double: { val: "double" },
    singleAccounting: { val: "singleAccounting" },
    doubleAccounting: { val: "doubleAccounting" },
};

/**
 * The underline of a font, `<u>` (CT_UnderlineProperty). Rendered as `<u/>` for `true` or `"single"`, with `val` for
 * the other known styles; nothing for `false`, `"none"` or an unknown style. Parsed as the `val` attribute, or `true`
 * (single, the default) when it is absent or not one of ST_UnderlineValues.
 */
class UnderlineXform extends BaseXform<UnderlineModel> {
    /** The attributes of `<u>` by underline style (looked up as own properties: model strings may come from a file). */
    static readonly Attributes = ATTRIBUTES;

    /** @param model - the default model, rendered when `render` gets none */
    constructor(model?: UnderlineModel) {
        super();

        this.model = model;
    }

    get tag(): string {
        return "u";
    }

    /** Writes `<u>` for `model` (or, when it is falsy, for the model of this xform). */
    override render(xmlStream: XmlStream, model: UnderlineModel | undefined): void {
        const underline = model || this.model;

        if (underline === true) {
            xmlStream.leafNode("u");
        } else if (typeof underline === "string" && Object.hasOwn(ATTRIBUTES, underline)) {
            xmlStream.leafNode("u", ATTRIBUTES[underline]);
        }
    }

    /** Reads `<u>` (returns nothing: the parent decides when to delegate to this xform). */
    override parseOpen(node: SaxTag): void {
        if (node.name === "u") {
            const { val } = node.attributes;
            this.model = isOneOf(UNDERLINE_VALUES, val) ? val : true;
        }
    }

    override parseText(): void {}

    override parseClose(): boolean {
        return false;
    }
}

export default UnderlineXform;
