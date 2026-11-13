import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { CustomFilterModel } from "../../../types";
import BaseXform from "../base-xform";

/** A custom filter of a filter column, `<customFilter val operator>`. */
class CustomFilterXform extends BaseXform<CustomFilterModel> {
    get tag(): string {
        return "customFilter";
    }

    override render(xmlStream: XmlStream, model: CustomFilterModel): void {
        xmlStream.leafNode(this.tag, {
            val: model.val,
            operator: model.operator,
        });
    }

    /** @returns `true` for `<customFilter>` */
    override parseOpen(node: SaxTag): boolean {
        if (node.name === this.tag) {
            this.model = {
                val: node.attributes.val,
                operator: node.attributes.operator,
            };
            return true;
        }
        return false;
    }

    override parseText(): void {}

    override parseClose(): boolean {
        return false;
    }
}

export default CustomFilterXform;
