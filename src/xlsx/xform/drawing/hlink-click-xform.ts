import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { DrawingPictureModel } from "../../../types";
import BaseXform from "../base-xform";

/**
 * The hyperlink of a picture, `<a:hlinkClick r:id tooltip>`: rendered when the picture has a hyperlink relationship.
 * Parsed as `{ hyperlinks: { rId, tooltip } }`.
 */
class HLinkClickXform extends BaseXform<Pick<DrawingPictureModel, "hyperlinks">> {
    get tag(): string {
        return "a:hlinkClick";
    }

    override render(xmlStream: XmlStream, model: Pick<DrawingPictureModel, "hyperlinks">): void {
        if (!(model.hyperlinks && model.hyperlinks.rId)) {
            return;
        }
        xmlStream.leafNode(this.tag, {
            "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "r:id": model.hyperlinks.rId,
            tooltip: model.hyperlinks.tooltip,
        });
    }

    /** @returns `true` */
    override parseOpen(node: SaxTag): boolean {
        switch (node.name) {
            case this.tag:
                this.model = {
                    hyperlinks: {
                        rId: node.attributes["r:id"],
                        tooltip: node.attributes.tooltip,
                    },
                };
                return true;
            default:
                return true;
        }
    }

    override parseText(): void {}

    override parseClose(): boolean {
        return false;
    }
}

export default HLinkClickXform;
