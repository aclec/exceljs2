/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import CompositeXform from "../../composite-xform";

import CfRuleExtXform from "./cf-rule-ext-xform";
import ConditionalFormattingExtXform from "./conditional-formatting-ext-xform";

class ConditionalFormattingsExtXform extends CompositeXform {
    constructor() {
        super();

        this.map = {
            "x14:conditionalFormatting": (this.cfXform = new ConditionalFormattingExtXform()),
        };
    }

    get tag() {
        return "x14:conditionalFormattings";
    }

    hasContent(model) {
        // not cached on the model: it is the worksheet's own list, which may change between writes
        return model.some((cf) => cf.rules.some(CfRuleExtXform.isExt));
    }

    prepare(model, options) {
        model.forEach((cf) => {
            this.cfXform.prepare(cf, options);
        });
    }

    render(xmlStream, model) {
        if (this.hasContent(model)) {
            xmlStream.openNode(this.tag);
            model.forEach((cf) => this.cfXform.render(xmlStream, cf));
            xmlStream.closeNode();
        }
    }

    createNewModel() {
        return [];
    }

    onParserClose(name, parser) {
        // model is array of conditional formatting objects
        this.model.push(parser.model);
    }
}

export default ConditionalFormattingsExtXform;
