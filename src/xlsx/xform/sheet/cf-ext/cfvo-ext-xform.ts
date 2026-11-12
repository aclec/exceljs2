/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import CompositeXform from "../../composite-xform";

import CfvoXform from "../cf/cfvo-xform";
import FExtXform from "./f-ext-xform";

class CfvoExtXform extends CompositeXform {
    constructor() {
        super();

        this.map = {
            "xm:f": (this.fExtXform = new FExtXform()),
        };
    }

    get tag() {
        return "x14:cfvo";
    }

    render(xmlStream, model) {
        xmlStream.openNode(this.tag, {
            type: model.type,
        });
        if (model.value !== undefined) {
            this.fExtXform.render(xmlStream, model.value);
        }
        xmlStream.closeNode();
    }

    createNewModel(node) {
        return {
            type: node.attributes.type,
        };
    }

    onParserClose(name, parser) {
        switch (name) {
            case "xm:f":
                this.model.value = parser.model ? CfvoXform.toValue(parser.model) : 0;
                break;
        }
    }
}

export default CfvoExtXform;
