/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import BaseXform from "../base-xform";
import IntegerXform from "../simple/integer-xform";
import colCache from "../../../utils/col-cache";

// the cell of an anchor is in the worksheet grid (0-based): a huge index from the file would make readers of
// the anchor (Anchor col/row) work up to it
function clamp(value, max) {
    return typeof value === "number" ? Math.min(Math.max(value, 0), max) : value;
}

class CellPositionXform extends BaseXform {
    constructor(options) {
        super();

        this.tag = options.tag;
        this.map = {
            "xdr:col": new IntegerXform({ tag: "xdr:col", zero: true }),
            "xdr:colOff": new IntegerXform({ tag: "xdr:colOff", zero: true }),
            "xdr:row": new IntegerXform({ tag: "xdr:row", zero: true }),
            "xdr:rowOff": new IntegerXform({ tag: "xdr:rowOff", zero: true }),
        };
    }

    render(xmlStream, model) {
        xmlStream.openNode(this.tag);

        this.map["xdr:col"].render(xmlStream, model.nativeCol);
        this.map["xdr:colOff"].render(xmlStream, model.nativeColOff);

        this.map["xdr:row"].render(xmlStream, model.nativeRow);
        this.map["xdr:rowOff"].render(xmlStream, model.nativeRowOff);

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
                this.parser = undefined;
            }
            return true;
        }
        switch (name) {
            case this.tag:
                this.model = {
                    nativeCol: clamp(this.map["xdr:col"].model, colCache.MAX_COLS - 1),
                    nativeColOff: this.map["xdr:colOff"].model,
                    nativeRow: clamp(this.map["xdr:row"].model, colCache.MAX_ROWS - 1),
                    nativeRowOff: this.map["xdr:rowOff"].model,
                };
                return false;
            default:
                // not quite sure how we get here!
                return true;
        }
    }
}

export default CellPositionXform;
