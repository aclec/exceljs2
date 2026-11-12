import type { SaxTag } from "../../../utils/parse-sax";
import colCache from "../../../utils/col-cache";
import XmlStream, { type XmlAttributes } from "../../../utils/xml-stream";
import type { DrawingAnchorModel, DrawingModel, DrawingReconcileOptions, ImageRangeModel } from "../../../types";

import BaseXform, { ownEntry } from "../base-xform";
import TwoCellAnchorXform from "./two-cell-anchor-xform";
import OneCellAnchorXform from "./one-cell-anchor-xform";

/** The element of an anchor. */
type AnchorType = NonNullable<DrawingAnchorModel["anchorType"]>;

// a range with a bottom-right corner is a two-cell anchor
function getAnchorType(model: { readonly range: ImageRangeModel | string }): AnchorType {
    if (typeof model.range === "string") {
        // a decoded reference has no `br`: a one-cell anchor (decode throws for an invalid reference)
        colCache.decode(model.range);
        return "xdr:oneCellAnchor";
    }
    return model.range.br ? "xdr:twoCellAnchor" : "xdr:oneCellAnchor";
}

/**
 * The drawing part of a worksheet, `xl/drawings/drawingN.xml`: its pictures, each with a two-cell or a one-cell anchor.
 * Parsing another drawing (e.g. a chart's `<c:userShapes>`) gives no model.
 */
class DrawingXform extends BaseXform<DrawingModel> {
    /** The namespaces of `<xdr:wsDr>`. */
    static readonly DRAWING_ATTRIBUTES: XmlAttributes = {
        "xmlns:xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    };

    declare map: { readonly "xdr:twoCellAnchor": TwoCellAnchorXform; readonly "xdr:oneCellAnchor": OneCellAnchorXform };

    constructor() {
        super();

        this.map = {
            "xdr:twoCellAnchor": new TwoCellAnchorXform(),
            "xdr:oneCellAnchor": new OneCellAnchorXform(),
        };
    }

    /** Sets the `anchorType` of each anchor and prepares its picture (index and creation id). */
    override prepare(model: DrawingModel): void {
        // drawings are named drawingN with N unique in the workbook
        const match = /(\d+)$/.exec(model.name || "");
        const drawingNumber = match ? parseInt(match[1]!, 10) : 0;
        model.anchors.forEach((item, index) => {
            const anchorType = getAnchorType(item);
            item.anchorType = anchorType;
            const anchor = this.map[anchorType];
            anchor.prepare(item, { index, drawingNumber });
        });
    }

    get tag(): string {
        return "xdr:wsDr";
    }

    /** @param model - prepared: each anchor has its `anchorType` */
    override render(xmlStream: XmlStream, model: DrawingModel): void {
        xmlStream.openXml(XmlStream.StdDocAttributes);
        xmlStream.openNode(this.tag, DrawingXform.DRAWING_ATTRIBUTES);

        for (const item of model.anchors) {
            const anchor = this.map[item.anchorType!];
            anchor.render(xmlStream, item);
        }

        xmlStream.closeNode();
    }

    /** @returns `true`: every element of the part is accepted (unknown ones are skipped) */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        switch (node.name) {
            case this.tag:
                this.reset();
                this.model = {
                    anchors: [],
                };
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

    override parseText(text: string): void {
        if (this.parser) {
            this.parser.parseText(text);
        }
    }

    override parseClose(name: string): boolean {
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                // an anchor outside <xdr:wsDr> (no model) is ignored
                if (this.model) {
                    const anchor =
                        this.parser === this.map["xdr:twoCellAnchor"] ? this.map["xdr:twoCellAnchor"] : this.map["xdr:oneCellAnchor"];
                    // the anchor has its model: its element was closed
                    this.model.anchors.push(anchor.model!);
                }
                this.parser = undefined;
            }
            return true;
        }
        switch (name) {
            case this.tag:
                return false;
            default:
                // could be some unrecognised tags
                return true;
        }
    }

    /** Resolves the medium of each picture (see BaseCellAnchorXform.reconcilePicture). */
    override reconcile(model: DrawingModel, options: DrawingReconcileOptions): void {
        for (const anchor of model.anchors) {
            if (anchor.range.br) {
                this.map["xdr:twoCellAnchor"].reconcile(anchor, options);
            } else {
                this.map["xdr:oneCellAnchor"].reconcile(anchor, options);
            }
        }
    }
}

export default DrawingXform;
