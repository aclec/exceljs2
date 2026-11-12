/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import XmlStream from "../../../utils/xml-stream";

import BaseXform from "../base-xform";
import VmlShapeXform from "./vml-shape-xform";

// This class is (currently) single purposed to insert the triangle
// drawing icons on commented cells
class VmlNotesXform extends BaseXform {
    constructor() {
        super();
        this.map = {
            "v:shape": new VmlShapeXform(),
        };
    }

    get tag() {
        return "xml";
    }

    render(xmlStream, model) {
        xmlStream.openXml(XmlStream.StdDocAttributes);
        xmlStream.openNode(this.tag, VmlNotesXform.DRAWING_ATTRIBUTES);

        // shape ids are unique across the workbook: model.shapeBlock is the first of the blocks of this worksheet
        const firstBlock = model.shapeBlock || 1;
        xmlStream.openNode("o:shapelayout", { "v:ext": "edit" });
        xmlStream.leafNode("o:idmap", { "v:ext": "edit", data: VmlNotesXform.idmap(firstBlock, model.comments.length) });
        xmlStream.closeNode();

        xmlStream.openNode("v:shapetype", {
            id: "_x0000_t202",
            coordsize: "21600,21600",
            "o:spt": 202,
            path: "m,l,21600r21600,l21600,xe",
        });
        xmlStream.leafNode("v:stroke", { joinstyle: "miter" });
        xmlStream.leafNode("v:path", { gradientshapeok: "t", "o:connecttype": "rect" });
        xmlStream.closeNode();

        model.comments.forEach((item, index) => {
            this.map["v:shape"].render(xmlStream, item, VmlNotesXform.shapeIndex(firstBlock, index));
        });

        xmlStream.closeNode();
    }

    parseOpen(node) {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        switch (node.name) {
            case this.tag:
                this.reset();
                this.model = {
                    comments: [],
                };
                break;
            default:
                this.parser = this.map[node.name];
                if (this.parser) {
                    this.parser.parseOpen(node);
                }
                break;
        }
        return true;
    }

    parseText(text) {
        if (this.parser) {
            this.parser.parseText(text);
        }
    }

    parseClose(name) {
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                this.model.comments.push(this.parser.model);
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

    reconcile(model, options) {
        model.anchors.forEach((anchor) => {
            if (anchor.br) {
                this.map["xdr:twoCellAnchor"].reconcile(anchor, options);
            } else {
                this.map["xdr:oneCellAnchor"].reconcile(anchor, options);
            }
        });
    }
}

VmlNotesXform.DRAWING_ATTRIBUTES = {
    "xmlns:v": "urn:schemas-microsoft-com:vml",
    "xmlns:o": "urn:schemas-microsoft-com:office:office",
    "xmlns:x": "urn:schemas-microsoft-com:office:excel",
};

// VML shape ids must be unique across the workbook. They come in blocks of 1024 (block n holds ids n * 1024 to n * 1024 + 1023)
// that the drawing claims in <o:idmap data="n,m,...">; as Excel does, the shapes of block n are numbered from n * 1024 + 1, so
// a block holds 1023 notes. Each worksheet claims contiguous blocks, as many as its notes need.
VmlNotesXform.SHAPES_PER_BLOCK = 1023;

// the number of blocks of a worksheet with count notes
VmlNotesXform.blockCount = (count) => Math.max(1, Math.ceil(count / VmlNotesXform.SHAPES_PER_BLOCK));

// the idmap data of a worksheet whose count notes use the blocks from firstBlock
VmlNotesXform.idmap = (firstBlock, count) => Array.from({ length: VmlNotesXform.blockCount(count) }, (_, i) => firstBlock + i).join(",");

// the index for VmlShapeXform (which numbers shapes from 1025) of the note at index
VmlNotesXform.shapeIndex = (firstBlock, index) => {
    const block = firstBlock + Math.floor(index / VmlNotesXform.SHAPES_PER_BLOCK);
    return (block - 1) * 1024 + (index % VmlNotesXform.SHAPES_PER_BLOCK);
};

export default VmlNotesXform;
