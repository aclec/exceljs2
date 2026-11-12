/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import BaseXform from "../../base-xform";

class SqrefExtXform extends BaseXform {
    get tag() {
        return "xm:sqref";
    }

    render(xmlStream, model) {
        xmlStream.leafNode(this.tag, null, model);
    }

    parseOpen() {
        this.model = "";
    }

    parseText(text) {
        this.model += text;
    }

    parseClose(name) {
        return name !== this.tag;
    }
}

export default SqrefExtXform;
