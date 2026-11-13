import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import BaseXform from "../base-xform";

/** `<mergeCell ref="A1:B2"/>`: one merged range. The model is the `ref` string (`undefined` when absent). */
class MergeCellXform extends BaseXform<string> {
    get tag(): string {
        return "mergeCell";
    }

    override render(xmlStream: XmlStream, model: string): void {
        xmlStream.leafNode("mergeCell", { ref: model });
    }

    /** @returns `true` for `<mergeCell>` */
    override parseOpen(node: SaxTag): boolean {
        if (node.name === "mergeCell") {
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

export default MergeCellXform;
