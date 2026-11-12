/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import BaseXform from "../../base-xform";

class CfIconExtXform extends BaseXform {
    get tag() {
        return "x14:cfIcon";
    }

    render(xmlStream, model) {
        xmlStream.leafNode(this.tag, {
            iconSet: model.iconSet,
            iconId: model.iconId,
        });
    }

    parseOpen({ attributes }) {
        this.model = {
            iconSet: attributes.iconSet,
            iconId: BaseXform.toIntValue(attributes.iconId),
        };
    }

    parseClose(name) {
        return name !== this.tag;
    }
}

export default CfIconExtXform;
