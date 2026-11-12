import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { WorkbookPivotCacheModel } from "../../../types";
import BaseXform from "../base-xform";

/** A `<pivotCache cacheId r:id>` of `xl/workbook.xml` (rendered for each pivot table). */
class WorkbookPivotCacheXform extends BaseXform<WorkbookPivotCacheModel> {
    override render(xmlStream: XmlStream, model: WorkbookPivotCacheModel): void {
        xmlStream.leafNode("pivotCache", {
            cacheId: model.cacheId,
            "r:id": model.rId,
        });
    }

    /** @returns `true` for `<pivotCache>` */
    override parseOpen(node: SaxTag): boolean {
        if (node.name === "pivotCache") {
            this.model = {
                cacheId: node.attributes.cacheId,
                rId: node.attributes["r:id"],
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

export default WorkbookPivotCacheXform;
