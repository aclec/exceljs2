import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import BaseXform from "../base-xform";

/** `<dimension ref="A1:D10"/>`: the used range of a worksheet. The model is the `ref` string (not rendered when empty). */
class DimensionXform extends BaseXform<string> {
    get tag(): string {
        return "dimension";
    }

    override render(xmlStream: XmlStream, model: string | undefined): void {
        if (model) {
            xmlStream.leafNode("dimension", { ref: model });
        }
    }

    /** @returns `true` for `<dimension>` (the model is its `ref`, `undefined` when absent) */
    override parseOpen(node: SaxTag): boolean {
        if (node.name === "dimension") {
            this.model = node.attributes.ref;
            return true;
        }
        return false;
    }

    override parseText(): void {}

    override parseClose(): boolean {
        return false;
    }
}

export default DimensionXform;
