import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { SheetPartRefModel } from "../../../types";
import BaseXform from "../base-xform";

/**
 * `<drawing r:id="…"/>`: the drawing part of a worksheet (its images); nothing is rendered without a model. Subclassed
 * (with another `tag`) for `<legacyDrawing>`, the VML of the notes.
 */
class DrawingXform extends BaseXform<SheetPartRefModel> {
    get tag(): string {
        return "drawing";
    }

    override render(xmlStream: XmlStream, model: SheetPartRefModel | undefined): void {
        if (model) {
            xmlStream.leafNode(this.tag, { "r:id": model.rId });
        }
    }

    /** @returns `true` for its element ({@link tag}) */
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

export default DrawingXform;
