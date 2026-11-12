import type XmlStream from "../../../utils/xml-stream";
import type { DrawingAnchorModel, DrawingReconcileOptions } from "../../../types";
import BaseCellAnchorXform from "./base-cell-anchor-xform";
import StaticXform from "../static-xform";

import CellPositionXform from "./cell-position-xform";
import PicXform, { type PicPrepareOptions } from "./pic-xform";

/**
 * A picture anchored by two cells, `<xdr:twoCellAnchor editAs>` (CT_TwoCellAnchor order: from, to, pic, clientData):
 * its range is `{ editAs, tl, br }`. Parsing gives no model (`null`) for an anchor without `<xdr:from>` or `<xdr:to>`.
 */
class TwoCellAnchorXform extends BaseCellAnchorXform {
    declare map: {
        readonly "xdr:from": CellPositionXform;
        readonly "xdr:to": CellPositionXform;
        readonly "xdr:pic": PicXform;
        readonly "xdr:clientData": StaticXform;
    };

    constructor() {
        super();

        this.map = {
            "xdr:from": new CellPositionXform({ tag: "xdr:from" }),
            "xdr:to": new CellPositionXform({ tag: "xdr:to" }),
            "xdr:pic": new PicXform(),
            "xdr:clientData": new StaticXform({ tag: "xdr:clientData" }),
        };
    }

    get tag(): string {
        return "xdr:twoCellAnchor";
    }

    /** Prepares the picture (its index and creation id). */
    override prepare(model: DrawingAnchorModel, options: PicPrepareOptions): void {
        // the anchors of a drawing are built with their picture (worksheet-xform prepare)
        if (model.picture) {
            this.map["xdr:pic"].prepare(model.picture, options);
        }
    }

    /** @throws Error for an anchor without bottom-right corner (`range.br`) or picture */
    override render(xmlStream: XmlStream, model: DrawingAnchorModel): void {
        const { range, picture } = model;
        // DrawingXform chose the anchor type by the bottom-right corner
        if (!range.br || !picture) {
            throw new Error("An image anchored by two cells needs its bottom-right corner (range.br) and its picture");
        }
        xmlStream.openNode(this.tag, { editAs: range.editAs || "oneCell" });

        this.map["xdr:from"].render(xmlStream, range.tl);
        this.map["xdr:to"].render(xmlStream, range.br);
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
                // the children were reset when the element opened (null when absent): an anchor without its corners
                // (CT_TwoCellAnchor requires them) has no model; one without picture (a chart, a shape) has no picture
                const tl = this.map["xdr:from"].model;
                const br = this.map["xdr:to"].model;
                const picture = this.map["xdr:pic"].model;
                this.model = tl && br ? { range: { editAs: this.parsedEditAs, tl, br }, ...(picture ? { picture } : {}) } : null;
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

export default TwoCellAnchorXform;
