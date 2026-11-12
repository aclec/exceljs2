import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { FilterModel } from "../../../types";
import BaseXform from "../base-xform";

/** A value of a filter column, `<filter val>`. */
class FilterXform extends BaseXform<FilterModel> {
    get tag(): string {
        return "filter";
    }

    override render(xmlStream: XmlStream, model: FilterModel): void {
        xmlStream.leafNode(this.tag, {
            val: model.val,
        });
    }

    /** @returns `true` for `<filter>` */
    override parseOpen(node: SaxTag): boolean {
        if (node.name === this.tag) {
            this.model = {
                val: node.attributes.val,
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

export default FilterXform;
