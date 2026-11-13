import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { DrawingPictureModel } from "../../../types";
import BaseXform from "../base-xform";

/**
 * The extensions of the non-visual properties of a picture, `<a:extLst>`: its `a16:creationId`. Nothing is parsed (the
 * children are skipped).
 */
class ExtLstXform extends BaseXform<DrawingPictureModel> {
    get tag(): string {
        return "a:extLst";
    }

    /**
     * a16:creationId must be unique per picture in the workbook: derives a GUID-shaped id from the drawing number and the
     * picture index.
     *
     * @param drawingNumber - the N of `drawingN` (4 hex digits)
     * @param index - the 1-based index of the picture in the drawing (12 hex digits)
     */
    static creationId(this: void, drawingNumber: number, index: number): string {
        const hex = (value: number, length: number): string => value.toString(16).toUpperCase().padStart(length, "0");
        return `{00000000-0008-0000-${hex(drawingNumber, 4)}-${hex(index, 12)}}`;
    }

    /** Writes the creation id of the picture (or one derived from its index, when it has none). */
    override render(xmlStream: XmlStream, model: DrawingPictureModel | undefined): void {
        xmlStream.openNode(this.tag);
        xmlStream.openNode("a:ext", {
            uri: "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}",
        });
        xmlStream.leafNode("a16:creationId", {
            "xmlns:a16": "http://schemas.microsoft.com/office/drawing/2014/main",
            id: (model && model.creationId) || ExtLstXform.creationId(0, (model && model.index) || 0),
        });
        xmlStream.closeNode();
        xmlStream.closeNode();
    }

    /** @returns `true`: every element inside `<a:extLst>` is accepted (and skipped) */
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

export default ExtLstXform;
