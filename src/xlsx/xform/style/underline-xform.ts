/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import BaseXform from "../base-xform";

class UnderlineXform extends BaseXform {
    constructor(model) {
        super();

        this.model = model;
    }

    get tag() {
        return "u";
    }

    render(xmlStream, model) {
        model = model || this.model;

        if (model === true) {
            xmlStream.leafNode("u");
        } else {
            const attr = UnderlineXform.Attributes[model];
            if (attr) {
                xmlStream.leafNode("u", attr);
            }
        }
    }

    parseOpen(node) {
        if (node.name === "u") {
            this.model = node.attributes.val || true;
        }
    }

    parseText() {}

    parseClose() {
        return false;
    }
}

UnderlineXform.Attributes = {
    single: {},
    double: { val: "double" },
    singleAccounting: { val: "singleAccounting" },
    doubleAccounting: { val: "doubleAccounting" },
};

export default UnderlineXform;
