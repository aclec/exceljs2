import type { SaxTag } from "../../../utils/parse-sax";
import colCache from "../../../utils/col-cache";
import XmlStream, { type XmlAttributes } from "../../../utils/xml-stream";
import type { DrawingAnchorModel, DrawingModel, DrawingReconcileOptions, ImageRangeModel } from "../../../types";

import BaseXform, { ownEntry } from "../base-xform";
import TwoCellAnchorXform from "./two-cell-anchor-xform";
import OneCellAnchorXform from "./one-cell-anchor-xform";
import { ANCHOR_ELEMENTS, ChartAnchorCapture, renderChartAnchor } from "./chart-anchor";

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

// the namespaces declared by the attributes of an element (prefix → URI; "" for the default namespace)
function namespaceDeclarations(node: SaxTag): Map<string, string> {
    const scope = new Map<string, string>();
    for (const [name, value] of Object.entries(node.attributes)) {
        if (name === "xmlns") {
            scope.set("", value);
        } else if (name.startsWith("xmlns:") && !/[\s<>&"'/=:]/.test(name.slice(6))) {
            scope.set(name.slice(6), value);
        }
    }
    return scope;
}

// the namespaces of the <xdr:wsDr> written (DrawingXform.DRAWING_ATTRIBUTES)
const WRITTEN_SCOPE: ReadonlyMap<string, string> = new Map([
    ["xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"],
    ["a", "http://schemas.openxmlformats.org/drawingml/2006/main"],
]);

/**
 * The drawing part of a worksheet, `xl/drawings/drawingN.xml`: its pictures, each with a two-cell or a one-cell anchor,
 * and its charts (`chartFrames`, carried opaque: see chart-anchor.ts), written after the pictures. Parsing another
 * drawing (e.g. a chart's `<c:userShapes>`) gives no model.
 */
class DrawingXform extends BaseXform<DrawingModel> {
    /** The namespaces of `<xdr:wsDr>`. */
    static readonly DRAWING_ATTRIBUTES: XmlAttributes = {
        "xmlns:xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    };

    declare map: { readonly "xdr:twoCellAnchor": TwoCellAnchorXform; readonly "xdr:oneCellAnchor": OneCellAnchorXform };

    // parse: the depth of the element being parsed (1 inside <xdr:wsDr>), the namespaces declared on <xdr:wsDr>, and the
    // capture of the anchor being parsed, when it may hold a chart
    private depth = 0;
    private rootScope: ReadonlyMap<string, string> = new Map();
    private chartCapture: ChartAnchorCapture | undefined;

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
        const match = /\d+$/.exec(model.name || "");
        const drawingNumber = match ? parseInt(match[0], 10) : 0;
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

    /** @param model - prepared: each anchor has its `anchorType` (else it is chosen by its range) */
    override render(xmlStream: XmlStream, model: DrawingModel): void {
        xmlStream.openXml(XmlStream.StdDocAttributes);
        xmlStream.openNode(this.tag, DrawingXform.DRAWING_ATTRIBUTES);

        for (const item of model.anchors) {
            const anchor = this.map[item.anchorType ?? getAnchorType(item)];
            anchor.render(xmlStream, item);
        }
        // the pictures have the shape ids 1 to anchors.length (pic-xform prepare)
        let shapeId = model.anchors.length;
        const nextShapeId = (): number => ++shapeId;
        for (const chart of model.chartAnchors ?? []) {
            renderChartAnchor(xmlStream, chart.anchor, chart.ids, nextShapeId);
        }

        xmlStream.closeNode();
    }

    /** @returns `true`: every element of the part is accepted (unknown ones are skipped) */
    override parseOpen(node: SaxTag): boolean {
        if (this.chartCapture) {
            this.chartCapture.open(node);
        } else if (this.depth === 1 && this.model && ANCHOR_ELEMENTS.has(node.name)) {
            this.chartCapture = new ChartAnchorCapture(this.rootScope, WRITTEN_SCOPE);
            this.chartCapture.open(node);
        }
        this.depth++;
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
                this.depth = 1;
                this.chartCapture = undefined;
                this.rootScope = namespaceDeclarations(node);
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
        if (this.chartCapture) {
            this.chartCapture.text(text);
        }
        if (this.parser) {
            this.parser.parseText(text);
        }
    }

    override parseClose(name: string): boolean {
        this.depth--;
        // the capture of a child of <xdr:wsDr> (an anchor, or <mc:AlternateContent>), complete when that child ends
        const capture = this.chartCapture;
        const captured = capture ? capture.close() : false;
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                // an anchor outside <xdr:wsDr> (no model), or invalid (no model of the anchor), is ignored
                const anchor =
                    this.parser === this.map["xdr:twoCellAnchor"]
                        ? this.map["xdr:twoCellAnchor"].model
                        : this.map["xdr:oneCellAnchor"].model;
                if (this.model && anchor) {
                    this.model.anchors.push(anchor);
                }
                this.parser = undefined;
                if (anchor && anchor.picture && capture) {
                    // a picture (e.g. in a group with a chart) is read as an image, as it always was: not carried twice
                    capture.discard();
                }
            }
        }
        if (capture && captured) {
            this.chartCapture = undefined;
            const chartFrame = capture.result;
            if (chartFrame && this.model) {
                (this.model.chartFrames ??= []).push(chartFrame);
            }
        }
        if (this.parser) {
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
