import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { SheetPartRefModel } from "../../../types";
import BaseXform from "../base-xform";

/** `<tablePart r:id="…"/>`: a table of a worksheet (an item of `<tableParts>`). */
class TablePartXform extends BaseXform<SheetPartRefModel> {
    get tag(): string {
        return "tablePart";
    }

    override render(xmlStream: XmlStream, model: SheetPartRefModel | undefined): void {
        if (model) {
            xmlStream.leafNode(this.tag, { "r:id": model.rId });
        }
    }

    /** @returns `true` for `<tablePart>` */
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

export default TablePartXform;
