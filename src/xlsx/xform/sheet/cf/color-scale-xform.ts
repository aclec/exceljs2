/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import CompositeXform from "../../composite-xform";

import ColorXform from "../../style/color-xform";
import CfvoXform from "./cfvo-xform";

class ColorScaleXform extends CompositeXform {
    constructor() {
        super();

        this.map = {
            cfvo: (this.cfvoXform = new CfvoXform()),
            color: (this.colorXform = new ColorXform()),
        };
    }

    get tag() {
        return "colorScale";
    }

    render(xmlStream, model) {
        xmlStream.openNode(this.tag);

        model.cfvo.forEach((cfvo) => {
            this.cfvoXform.render(xmlStream, cfvo);
        });
        model.color.forEach((color) => {
            this.colorXform.render(xmlStream, color);
        });

        xmlStream.closeNode();
    }

    createNewModel(node) {
        return {
            cfvo: [],
            color: [],
        };
    }

    onParserClose(name, parser) {
        this.model[name].push(parser.model);
    }
}

export default ColorScaleXform;
