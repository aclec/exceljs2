import type { SaxTag } from "../../../../utils/parse-sax";
import type XmlStream from "../../../../utils/xml-stream";
import type { ConditionalFormattingModel } from "../../../../types";
import CompositeXform from "../../composite-xform";

import CfRuleXform from "./cf-rule-xform";

/** The model of a parsed `<conditionalFormatting>`: `ref` is absent when it has no `sqref` (malformed). */
export type ParsedConditionalFormattingModel = Omit<ConditionalFormattingModel, "ref"> & { ref?: string };

/**
 * `<conditionalFormatting sqref="A1:B10">`: the Excel 2007 rules of a range (see CfRuleXform.isPrimitive). Nothing is
 * rendered when the range has none. Parsing gives the rules read (see CfRuleXform: those of a type that the model does
 * not have are not read).
 */
class ConditionalFormattingXform extends CompositeXform<ParsedConditionalFormattingModel | ConditionalFormattingModel> {
    /** The parsed element (the xform renders a ConditionalFormattingModel). */
    declare model: ParsedConditionalFormattingModel | null | undefined;

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

    override createNewModel({ attributes }: SaxTag): ParsedConditionalFormattingModel {
        return {
            // a <conditionalFormatting> without sqref (malformed) has an undefined ref
            ref: attributes.sqref,
            rules: [],
        };
    }

    override onParserClose(): void {
        // the element was opened before its children (createNewModel); the rule model is set by its start tag (null for
        // a rule that is not read)
        const { model } = this;
        const rule = this.map.cfRule.model;
        if (model && rule) {
            model.rules.push(rule);
        }
    }
}

export default ConditionalFormattingXform;
