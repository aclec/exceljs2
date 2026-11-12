/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import BaseXform from "../../base-xform";

class CfvoXform extends BaseXform {
    get tag() {
        return "cfvo";
    }

    render(xmlStream, model) {
        xmlStream.leafNode(this.tag, {
            type: model.type,
            val: model.value,
        });
    }

    parseOpen(node) {
        this.model = {
            type: node.attributes.type,
            value: CfvoXform.toValue(node.attributes.val),
        };
    }

    parseClose(name) {
        return name !== this.tag;
    }
}

// a number literal ("50", "-1.5", "1E3"); each digit run can only be matched one way, so testing a long
// non-numeric value ("111...1x") is linear instead of quadratic
const NUMBER = /^[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:E[-+]?\d+)?$/i;

// the value of a threshold is a number, or a formula or cell reference ("$A$1", "AVERAGE(A:A)")
// kept as written
CfvoXform.toValue = (text) => {
    if (text === undefined) {
        return undefined;
    }
    return NUMBER.test(text) ? parseFloat(text) : text;
};

export default CfvoXform;
