import type XmlStream from "../../../utils/xml-stream";
import type { AnchorModel, DrawingAnchorModel, DrawingReconcileOptions } from "../../../types";
import BaseCellAnchorXform from "./base-cell-anchor-xform";
import StaticXform from "../static-xform";

import CellPositionXform from "./cell-position-xform";
import PicXform, { type PicPrepareOptions } from "./pic-xform";

/**
 * A picture anchored by two cells, `<xdr:twoCellAnchor editAs>` (CT_TwoCellAnchor order: from, to, pic, clientData):
 * its range is `{ editAs, tl, br }`.
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
        this.map["xdr:pic"].prepare(model.picture!, options);
    }

    override render(xmlStream: XmlStream, model: DrawingAnchorModel): void {
        xmlStream.openNode(this.tag, { editAs: model.range.editAs || "oneCell" });

        this.map["xdr:from"].render(xmlStream, model.range.tl);
        // a two-cell anchor has its bottom-right corner (DrawingXform.prepare chose the anchor type by it)
        this.map["xdr:to"].render(xmlStream, model.range.br!);
        this.map["xdr:pic"].render(xmlStream, model.picture!);
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
            case this.tag:
                // the children were reset when the element opened: an element absent from the anchor leaves null in
                // its place (unvalidated, see DrawingAnchorModel)
                this.model = {
                    range: {
                        editAs: this.parsedEditAs,
                        tl: this.map["xdr:from"].model as AnchorModel,
                        br: this.map["xdr:to"].model as AnchorModel,
                    },
                    picture: this.map["xdr:pic"].model as DrawingAnchorModel["picture"],
                };
                return false;
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
