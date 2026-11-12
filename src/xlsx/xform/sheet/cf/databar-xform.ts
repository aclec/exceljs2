/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import BaseXform from "../../base-xform";
import CompositeXform from "../../composite-xform";

import ColorXform from "../../style/color-xform";
import CfvoXform from "./cfvo-xform";

class DatabarXform extends CompositeXform {
    constructor() {
        super();

        this.map = {
            cfvo: (this.cfvoXform = new CfvoXform()),
            color: (this.colorXform = new ColorXform()),
        };
    }

    get tag() {
        return "dataBar";
    }

    render(xmlStream, model) {
        xmlStream.openNode(this.tag, { showValue: BaseXform.toBoolAttribute(model.showValue, true) });

        // CT_DataBar requires two cfvo and a color: fall back to Excel's defaults
        const cfvoList = model.cfvo && model.cfvo.length ? model.cfvo : DatabarXform.DEFAULT_CFVO;
        cfvoList.forEach((cfvo) => {
            // the x14 types autoMin / autoMax (written in the extension) are min / max for Excel 2007
            const type = DatabarXform.MAIN_CFVO_TYPE[cfvo.type];
            this.cfvoXform.render(xmlStream, type ? { ...cfvo, type } : cfvo);
        });
        this.colorXform.render(xmlStream, model.color || DatabarXform.DEFAULT_COLOR);

        xmlStream.closeNode();
    }

    createNewModel({ attributes }) {
        return {
            cfvo: [],
            showValue: BaseXform.toBoolValue(attributes.showValue),
        };
    }

    onParserClose(name, parser) {
        switch (name) {
            case "cfvo":
                this.model.cfvo.push(parser.model);
                break;
            case "color":
                this.model.color = parser.model;
                break;
        }
    }
}

DatabarXform.DEFAULT_CFVO = [{ type: "min" }, { type: "max" }];
DatabarXform.DEFAULT_COLOR = { argb: "FF638EC6" };
DatabarXform.MAIN_CFVO_TYPE = { autoMin: "min", autoMax: "max" };

export default DatabarXform;
