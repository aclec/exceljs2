/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import BaseXform from "../base-xform";

class RelationshipXform extends BaseXform {
    render(xmlStream, model) {
        xmlStream.leafNode("Relationship", model);
    }

    parseOpen(node) {
        switch (node.name) {
            case "Relationship":
                this.model = node.attributes;
                return true;
            default:
                return false;
        }
    }

    parseText() {}

    parseClose() {
        return false;
    }
}

export default RelationshipXform;
