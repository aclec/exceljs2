import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { DrawingPictureModel } from "../../../types";
import BaseXform, { ownEntry } from "../base-xform";
import BlipXform from "./blip-xform";

/**
 * The fill of a picture, `<xdr:blipFill>`: its image (`<a:blip>`), stretched (CT_BlipFillProperties order: blip, then
 * stretch). Parsed as the model of the blip (`null` without `<a:blip>`).
 */
class BlipFillXform extends BaseXform<DrawingPictureModel> {
    declare map: { readonly "a:blip": BlipXform };

    constructor() {
        super();

        this.map = {
            "a:blip": new BlipXform(),
        };
    }

    get tag(): string {
        return "xdr:blipFill";
    }

    override render(xmlStream: XmlStream, model: DrawingPictureModel): void {
        xmlStream.openNode(this.tag);

        this.map["a:blip"].render(xmlStream, model);

        // TODO: options for this + parsing
        xmlStream.openNode("a:stretch");
        xmlStream.leafNode("a:fillRect");
        xmlStream.closeNode();

        xmlStream.closeNode();
    }

    /** @returns `true`: every element inside `<xdr:blipFill>` is accepted (unknown ones are skipped) */
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
                this.model = this.map["a:blip"].model;
                return false;

            default:
                return true;
        }
    }
}

export default BlipFillXform;
