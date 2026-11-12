import type XmlStream from "../../../../utils/xml-stream";
import type { ConditionalFormattingModel } from "../../../../types";
import CompositeXform from "../../composite-xform";

import SqRefExtXform from "./sqref-ext-xform";
import CfRuleExtXform from "./cf-rule-ext-xform";

/** The model of a parsed `<x14:conditionalFormatting>`: `ref` is absent when it has no `<xm:sqref>`. */
export type ConditionalFormattingExtModel = Omit<ConditionalFormattingModel, "ref"> & { ref?: string };

/**
 * `<x14:conditionalFormatting>`: the x14 rules of a range (see CfRuleExtXform.isExt), then its `<xm:sqref>`. Nothing is
 * rendered when the range has no x14 rule.
 */
class ConditionalFormattingExtXform extends CompositeXform<ConditionalFormattingExtModel> {
    /** The xform of the range. */
    readonly sqRef = new SqRefExtXform();
    /** The xform of the rules. */
    readonly cfRule = new CfRuleExtXform();

    declare map: { readonly "xm:sqref": SqRefExtXform; readonly "x14:cfRule": CfRuleExtXform };

    constructor() {
        super();

        this.map = {
            "xm:sqref": this.sqRef,
            "x14:cfRule": this.cfRule,
        };
    }

    get tag(): string {
        return "x14:conditionalFormatting";
    }

    /** Gives the x14 rules their `x14Id` (mutates the rules). */
    override prepare(model: ConditionalFormattingExtModel): void {
        for (const rule of model.rules) {
            this.cfRule.prepare(rule);
        }
    }

    /** Writes the x14 rules of the range; each of them gets the range as `ref` (mutates the rules). */
    override render(xmlStream: XmlStream, model: ConditionalFormattingExtModel): void {
        if (!model.rules.some(CfRuleExtXform.isExt)) {
            return;
        }

        xmlStream.openNode(this.tag, {
            "xmlns:xm": "http://schemas.microsoft.com/office/excel/2006/main",
        });

        for (const rule of model.rules.filter(CfRuleExtXform.isExt)) {
            // text and time period rules without formulae get one for the top left cell of the range
            rule.ref = model.ref;
            this.cfRule.render(xmlStream, rule);
        }

        // for some odd reason, Excel needs the <xm:sqref> node to be after the rules
        this.sqRef.render(xmlStream, model.ref);

        xmlStream.closeNode();
    }

    override createNewModel(): ConditionalFormattingExtModel {
        return {
            rules: [],
        };
    }

    override onParserClose(name: string): void {
        // the element was opened before its children (createNewModel); a child model is set by its start tag (null for a
        // rule that is not read)
        const { model } = this;
        if (!model) {
            return;
        }
        switch (name) {
            case "xm:sqref":
                model.ref = this.sqRef.model ?? undefined;
                break;

            case "x14:cfRule": {
                const rule = this.cfRule.model;
                if (rule) {
                    model.rules.push(rule);
                }
                break;
            }
        }
    }
}

export default ConditionalFormattingExtXform;
