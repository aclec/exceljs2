import type { SaxTag } from "../../../../utils/parse-sax";
import type XmlStream from "../../../../utils/xml-stream";
import type { CfvoModel } from "../../../../types";
import BaseXform from "../../base-xform";

// a number literal ("50", "-1.5", "1E3"); each digit run can only be matched one way, so testing a long
// non-numeric value ("111...1x") is linear instead of quadratic
const NUMBER = /^[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:E[-+]?\d+)?$/i;

/**
 * `<cfvo type="num" val="50"/>`: a threshold of a color scale, data bar or icon set. The parsed `type` is the attribute
 * (unvalidated).
 */
class CfvoXform extends BaseXform<CfvoModel> {
    get tag(): string {
        return "cfvo";
    }

    override render(xmlStream: XmlStream, model: CfvoModel): void {
        xmlStream.leafNode(this.tag, {
            type: model.type,
            val: model.value,
        });
    }

    /** Reads the attributes of the element (the parent delegates `<cfvo>` only). */
    override parseOpen(node: SaxTag): void {
        this.model = {
            // unvalidated ST_CfvoType
            type: node.attributes.type as CfvoModel["type"],
            value: CfvoXform.toValue(node.attributes.val),
        };
    }

    /** @returns `false` at the end of `<cfvo>` */
    override parseClose(name: string): boolean {
        return name !== this.tag;
    }

    /**
     * The value of a threshold: a number for a number literal, else the text kept as written (a formula or a cell
     * reference such as `$A$1`, `AVERAGE(A:A)`). Linear in the length of the text.
     *
     * @returns `undefined` for `undefined`
     */
    static toValue(this: void, text: string | undefined): number | string | undefined {
        if (text === undefined) {
            return undefined;
        }
        return NUMBER.test(text) ? parseFloat(text) : text;
    }
}

export default CfvoXform;
