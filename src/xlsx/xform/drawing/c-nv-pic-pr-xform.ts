/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import BaseXform from "../base-xform";

class CNvPicPrXform extends BaseXform {
    get tag() {
        return "xdr:cNvPicPr";
    }

    render(xmlStream) {
        xmlStream.openNode(this.tag);
        xmlStream.leafNode("a:picLocks", {
            noChangeAspect: "1",
        });
        xmlStream.closeNode();
    }

    parseOpen(node) {
        switch (node.name) {
            case this.tag:
                return true;
            default:
                return true;
        }
    }

    parseText() {}

    parseClose(name) {
        switch (name) {
            case this.tag:
                return false;
            default:
                // unprocessed internal nodes
                return true;
        }
    }
}

export default CNvPicPrXform;
