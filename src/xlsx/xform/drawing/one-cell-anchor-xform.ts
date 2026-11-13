import type XmlStream from "../../../utils/xml-stream";
import type { DrawingAnchorModel, DrawingReconcileOptions } from "../../../types";
import BaseCellAnchorXform from "./base-cell-anchor-xform";
import StaticXform from "../static-xform";

import CellPositionXform from "./cell-position-xform";
import ExtXform from "./ext-xform";
import PicXform, { type PicPrepareOptions } from "./pic-xform";

/**
 * A picture anchored by one cell, `<xdr:oneCellAnchor>` (CT_OneCellAnchor order: from, ext, pic, clientData): its range
 * is `{ editAs, tl, ext }` (`ext` in pixels). Parsing gives no model (`null`) for an anchor without `<xdr:from>` or
 * `<xdr:ext>`.
 */
class OneCellAnchorXform extends BaseCellAnchorXform {
    declare map: {
        readonly "xdr:from": CellPositionXform;
        readonly "xdr:ext": ExtXform;
        readonly "xdr:pic": PicXform;
        readonly "xdr:clientData": StaticXform;
    };

    constructor() {
        super();

        this.map = {
            "xdr:from": new CellPositionXform({ tag: "xdr:from" }),
            "xdr:ext": new ExtXform({ tag: "xdr:ext" }),
            "xdr:pic": new PicXform(),
            "xdr:clientData": new StaticXform({ tag: "xdr:clientData" }),
        };
    }

    get tag(): string {
        return "xdr:oneCellAnchor";
    }

    /** Prepares the picture (its index and creation id). */
    override prepare(model: DrawingAnchorModel, options: PicPrepareOptions): void {
        // the anchors of a drawing are built with their picture (worksheet-xform prepare)
        if (model.picture) {
            this.map["xdr:pic"].prepare(model.picture, options);
        }
    }

    /**
     * Writes the anchor, with an `editAs` attribute as for a two-cell anchor (CT_OneCellAnchor declares none).
     *
     * @throws Error for an anchor without size (`range.ext`) or picture
     */
    override render(xmlStream: XmlStream, model: DrawingAnchorModel): void {
        const { range, picture } = model;
        if (!range.ext || !picture) {
            throw new Error("An image anchored by one cell needs its size (range.ext: width and height) and its picture");
        }
        xmlStream.openNode(this.tag, { editAs: range.editAs || "oneCell" });

        this.map["xdr:from"].render(xmlStream, range.tl);
        this.map["xdr:ext"].render(xmlStream, range.ext);
        this.map["xdr:pic"].render(xmlStream, picture);
        this.map["xdr:clientData"].render(xmlStream);

        xmlStream.closeNode();
    }

    override parseClose(name: string): boolean {
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                this.parser = undefined;
            }
            return true;
        }
        switch (name) {
            case this.tag: {
                // the children were reset when the element opened (null when absent): an anchor without its corner or its
                // size (CT_OneCellAnchor requires them) has no model; one without picture (a chart, a shape) has no picture
                const tl = this.map["xdr:from"].model;
                const ext = this.map["xdr:ext"].model;
                const picture = this.map["xdr:pic"].model;
                this.model = tl && ext ? { range: { editAs: this.parsedEditAs, tl, ext }, ...(picture ? { picture } : {}) } : null;
                return false;
            }
            default:
                // could be some unrecognised tags
                return true;
        }
    }

    /** Sets the `medium` of the anchor: the workbook image of its picture. */
    override reconcile(model: DrawingAnchorModel, options: DrawingReconcileOptions): void {
        model.medium = this.reconcilePicture(model.picture, options);
    }
}

export default OneCellAnchorXform;
