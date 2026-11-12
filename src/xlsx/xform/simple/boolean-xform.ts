/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import BaseXform from "../base-xform";
import utils from "../../../utils/utils";

class BooleanXform extends BaseXform {
    constructor(options) {
        super();

        this.tag = options.tag;
        this.attr = options.attr;
    }

    render(xmlStream, model) {
        if (model) {
            xmlStream.openNode(this.tag);
            xmlStream.closeNode();
        }
    }

    parseOpen(node) {
        if (node.name === this.tag) {
            // CT_BooleanProperty: <b/> is true, <b val="0"/> or <b val="false"/> is false
            this.model = utils.parseBoolean(node.attributes[this.attr], true);
        }
    }

    parseText() {}

    parseClose() {
        return false;
    }
}

export default BooleanXform;
