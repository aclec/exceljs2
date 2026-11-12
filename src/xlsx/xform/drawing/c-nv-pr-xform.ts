import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { DrawingPictureModel } from "../../../types";
import BaseXform, { ownEntry } from "../base-xform";
import HlickClickXform from "./hlink-click-xform";
import ExtLstXform from "./ext-lst-xform";

/**
 * The non-visual drawing properties of a picture, `<xdr:cNvPr id name>` (CT_NonVisualDrawingProps order: hlinkClick,
 * then extLst). Parsed as the model of its hyperlink (`null` without `<a:hlinkClick>`).
 */
class CNvPrXform extends BaseXform<Pick<DrawingPictureModel, "hyperlinks">> {
    declare map: { readonly "a:hlinkClick": HlickClickXform; readonly "a:extLst": ExtLstXform };

    constructor() {
        super();

        this.map = {
            "a:hlinkClick": new HlickClickXform(),
            "a:extLst": new ExtLstXform(),
        };
    }

    get tag(): string {
        return "xdr:cNvPr";
    }

    /** @param model - the picture: its `index` (prepared) is the id */
    override render(xmlStream: XmlStream, model: DrawingPictureModel): void {
        xmlStream.openNode(this.tag, {
            id: model.index,
            name: `Picture ${model.index}`,
        });
        this.map["a:hlinkClick"].render(xmlStream, model);
        this.map["a:extLst"].render(xmlStream, model);
        xmlStream.closeNode();
    }

    /** @returns `true`: every element inside `<xdr:cNvPr>` is accepted (unknown ones are skipped) */
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
                this.model = this.map["a:hlinkClick"].model;
                return false;
            default:
                return true;
        }
    }
}

export default CNvPrXform;
