import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { DrawingPictureModel } from "../../../types";
import BaseXform, { ownEntry } from "../base-xform";
import StaticXform from "../static-xform";

import BlipFillXform from "./blip-fill-xform";
import NvPicPrXform from "./nv-pic-pr-xform";
import ExtLstXform from "./ext-lst-xform";

import spPrJSON from "./sp-pr";

/** The prepare options of a picture (from DrawingXform.prepare). */
export interface PicPrepareOptions {
    /** The 0-based index of the anchor in its drawing. */
    readonly index: number;
    /** The N of `drawingN`. */
    readonly drawingNumber?: number;
}

/**
 * A picture, `<xdr:pic>` (CT_Picture order: nvPicPr, blipFill, spPr). Parsed as the merge of the models of its children:
 * `{ hyperlinks?, rId }`.
 */
class PicXform extends BaseXform<DrawingPictureModel> {
    declare map: {
        readonly "xdr:nvPicPr": NvPicPrXform;
        readonly "xdr:blipFill": BlipFillXform;
        readonly "xdr:spPr": StaticXform;
    };

    constructor() {
        super();

        this.map = {
            "xdr:nvPicPr": new NvPicPrXform(),
            "xdr:blipFill": new BlipFillXform(),
            "xdr:spPr": new StaticXform(spPrJSON),
        };
    }

    get tag(): string {
        return "xdr:pic";
    }

    /** Sets the 1-based `index` of the picture and its `creationId`. */
    override prepare(model: DrawingPictureModel, options: PicPrepareOptions): void {
        const index = options.index + 1;
        model.index = index;
        model.creationId = ExtLstXform.creationId(options.drawingNumber || 0, index);
    }

    override render(xmlStream: XmlStream, model: DrawingPictureModel): void {
        xmlStream.openNode(this.tag);

        this.map["xdr:nvPicPr"].render(xmlStream, model);
        this.map["xdr:blipFill"].render(xmlStream, model);
        this.map["xdr:spPr"].render(xmlStream);

        xmlStream.closeNode();
    }

    /** @returns `true`: every element inside `<xdr:pic>` is accepted (unknown ones are skipped) */
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
                // the static <xdr:spPr> has no model; a child without its element's content has a null model
                const { parser } = this;
                if (parser === this.map["xdr:nvPicPr"]) {
                    this.mergeModel(this.map["xdr:nvPicPr"].model ?? {});
                } else if (parser === this.map["xdr:blipFill"]) {
                    this.mergeModel(this.map["xdr:blipFill"].model ?? {});
                } else {
                    this.mergeModel({});
                }
                this.parser = undefined;
            }
            return true;
        }
        switch (name) {
            case this.tag:
                return false;
            default:
                // not quite sure how we get here!
                return true;
        }
    }
}

export default PicXform;
