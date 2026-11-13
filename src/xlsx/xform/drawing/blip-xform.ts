import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { DrawingPictureModel } from "../../../types";
import BaseXform from "../base-xform";

/** The image of a picture, `<a:blip r:embed>`: its relationship id. Its children are skipped when parsing. */
class BlipXform extends BaseXform<DrawingPictureModel> {
    get tag(): string {
        return "a:blip";
    }

    override render(xmlStream: XmlStream, model: DrawingPictureModel): void {
        xmlStream.leafNode(this.tag, {
            "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "r:embed": model.rId,
            cstate: "print",
        });
        // TODO: handle children (e.g. a:extLst=>a:ext=>a14:useLocalDpi
    }

    /** @returns `true`: every element inside `<a:blip>` is accepted (and skipped) */
    override parseOpen(node: SaxTag): boolean {
        switch (node.name) {
            case this.tag:
                this.model = {
                    rId: node.attributes["r:embed"],
                };
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

export default BlipXform;
