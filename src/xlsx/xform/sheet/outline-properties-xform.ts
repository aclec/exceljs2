import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { OutlinePropertiesModel } from "../../../types";
import { parseBoolean } from "../../../utils/utils";
import BaseXform from "../base-xform";

/**
 * `<outlinePr summaryBelow="0" summaryRight="1"/>` of `<sheetPr>`: where the summary rows / columns of the outlines are.
 * A parsed attribute that is absent is an `undefined` key.
 */
class OutlinePropertiesXform extends BaseXform<OutlinePropertiesModel> {
    get tag(): string {
        return "outlinePr";
    }

    /** @returns whether the element was written (only when `summaryBelow` or `summaryRight` is set) */
    override render(xmlStream: XmlStream, model: OutlinePropertiesModel | undefined): boolean {
        if (model && (model.summaryBelow !== undefined || model.summaryRight !== undefined)) {
            xmlStream.leafNode(this.tag, {
                summaryBelow: model.summaryBelow !== undefined ? Number(model.summaryBelow) : undefined,
                summaryRight: model.summaryRight !== undefined ? Number(model.summaryRight) : undefined,
            });
            return true;
        }
        return false;
    }

    /** @returns `true` for `<outlinePr>` */
    override parseOpen(node: SaxTag): boolean {
        if (node.name === this.tag) {
            const { summaryBelow, summaryRight } = node.attributes;
            this.model = {
                summaryBelow: summaryBelow !== undefined ? parseBoolean(summaryBelow) : undefined,
                summaryRight: summaryRight !== undefined ? parseBoolean(summaryRight) : undefined,
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

export default OutlinePropertiesXform;
