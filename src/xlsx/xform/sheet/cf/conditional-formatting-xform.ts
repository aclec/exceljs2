import type { SaxTag } from "../../../../utils/parse-sax";
import type XmlStream from "../../../../utils/xml-stream";
import type { ConditionalFormattingModel } from "../../../../types";
import CompositeXform from "../../composite-xform";

import CfRuleXform from "./cf-rule-xform";

/**
 * `<conditionalFormatting sqref="A1:B10">`: the Excel 2007 rules of a range (see CfRuleXform.isPrimitive). Nothing is
 * rendered when the range has none.
 */
class ConditionalFormattingXform extends CompositeXform<ConditionalFormattingModel> {
    declare map: { readonly cfRule: CfRuleXform };

    constructor() {
        super();

        this.map = {
            cfRule: new CfRuleXform(),
        };
    }

    get tag(): string {
        return "conditionalFormatting";
    }

    /** Writes the Excel 2007 rules of the range; each of them gets the range as `ref` (mutates the rules). */
    override render(xmlStream: XmlStream, model: ConditionalFormattingModel): void {
        // if there are no primitive rules, exit now
        if (!model.rules.some(CfRuleXform.isPrimitive)) {
            return;
        }

        xmlStream.openNode(this.tag, { sqref: model.ref });

        for (const rule of model.rules) {
            if (CfRuleXform.isPrimitive(rule)) {
                rule.ref = model.ref;
                this.map.cfRule.render(xmlStream, rule);
            }
        }

        xmlStream.closeNode();
    }

    override createNewModel({ attributes }: SaxTag): ConditionalFormattingModel {
        return {
            // a <conditionalFormatting> without sqref (malformed) has an undefined ref
            ref: attributes.sqref as string,
            rules: [],
        };
    }

    override onParserClose(): void {
        // the element was opened before its children (createNewModel); the rule model is set by its start tag
        this.model!.rules.push(this.map.cfRule.model!);
    }
}

export default ConditionalFormattingXform;
