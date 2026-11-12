import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { DrawingPictureModel } from "../../../types";
import BaseXform, { ownEntry } from "../base-xform";
import CNvPrXform from "./c-nv-pr-xform";
import CNvPicPrXform from "./c-nv-pic-pr-xform";

/**
 * The non-visual properties of a picture, `<xdr:nvPicPr>`: `<xdr:cNvPr>` then `<xdr:cNvPicPr>`. Parsed as the model of
 * `<xdr:cNvPr>` (its hyperlink).
 */
class NvPicPrXform extends BaseXform<Pick<DrawingPictureModel, "hyperlinks">> {
    declare map: { readonly "xdr:cNvPr": CNvPrXform; readonly "xdr:cNvPicPr": CNvPicPrXform };

    constructor() {
        super();

        this.map = {
            "xdr:cNvPr": new CNvPrXform(),
            "xdr:cNvPicPr": new CNvPicPrXform(),
        };
    }

    get tag(): string {
        return "xdr:nvPicPr";
    }

    override render(xmlStream: XmlStream, model: DrawingPictureModel): void {
        xmlStream.openNode(this.tag);
        this.map["xdr:cNvPr"].render(xmlStream, model);
        this.map["xdr:cNvPicPr"].render(xmlStream);
        xmlStream.closeNode();
    }

    /** @returns `true`: every element inside `<xdr:nvPicPr>` is accepted (unknown ones are skipped) */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }

        switch (node.name) {
            case this.tag:
                this.reset();
                break;
            default: {
                const parser = ownEntry(this.map, node.name);
                this.parser = parser;
                if (parser) {
                    parser.parseOpen(node);
                }
                break;
            }
        }
        return true;
    }

    override parseText(): void {}

    override parseClose(name: string): boolean {
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                this.parser = undefined;
            }
            return true;
        }
        switch (name) {
            case this.tag:
                this.model = this.map["xdr:cNvPr"].model;
                return false;
            default:
                return true;
        }
    }
}

export default NvPicPrXform;
