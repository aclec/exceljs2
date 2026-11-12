/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import BaseXform from "../base-xform";

class MergeCellXform extends BaseXform {
    get tag() {
        return "mergeCell";
    }

    render(xmlStream, model) {
        xmlStream.leafNode("mergeCell", { ref: model });
    }

    parseOpen(node) {
        if (node.name === "mergeCell") {
            this.model = node.attributes.ref;
            return true;
        }
        return false;
    }

    parseText() {}

    parseClose() {
        return false;
    }
}

export default MergeCellXform;
