import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { SheetPartRefModel } from "../../../types";
import BaseXform from "../base-xform";

/** `<picture r:id="…"/>`: the background image of a worksheet (nothing is rendered without a model). */
class PictureXform extends BaseXform<SheetPartRefModel> {
    get tag(): string {
        return "picture";
    }

    override render(xmlStream: XmlStream, model: SheetPartRefModel | undefined): void {
        if (model) {
            xmlStream.leafNode(this.tag, { "r:id": model.rId });
        }
    }

    /** @returns `true` for `<picture>` */
    override parseOpen(node: SaxTag): boolean {
        switch (node.name) {
            case this.tag:
                this.model = {
                    rId: node.attributes["r:id"],
                };
                return true;
            default:
                return false;
        }
    }

    override parseText(): void {}

    override parseClose(): boolean {
        return false;
    }
}

export default PictureXform;
