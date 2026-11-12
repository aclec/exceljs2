import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import BaseXform from "../base-xform";

/** The non-visual picture properties, `<xdr:cNvPicPr>`: the aspect ratio is locked. Nothing is parsed. */
class CNvPicPrXform extends BaseXform<never> {
    get tag(): string {
        return "xdr:cNvPicPr";
    }

    override render(xmlStream: XmlStream): void {
        xmlStream.openNode(this.tag);
        xmlStream.leafNode("a:picLocks", {
            noChangeAspect: "1",
        });
        xmlStream.closeNode();
    }

    /** @returns `true`: every element inside `<xdr:cNvPicPr>` is accepted (and skipped) */
    override parseOpen(node: SaxTag): boolean {
        switch (node.name) {
            case this.tag:
                return true;
            default:
                return true;
        }
    }

    override parseText(): void {}

    override parseClose(name: string): boolean {
        switch (name) {
            case this.tag:
                return false;
            default:
                // unprocessed internal nodes
                return true;
        }
    }
}

export default CNvPicPrXform;
