import type XmlStream from "../../../../utils/xml-stream";
import CompositeXform from "../../composite-xform";

import CfRuleExtXform from "./cf-rule-ext-xform";
import ConditionalFormattingExtXform, { type ConditionalFormattingExtModel } from "./conditional-formatting-ext-xform";

/**
 * `<x14:conditionalFormattings>` of the worksheet extension: the conditional formattings that have x14 rules. Nothing is
 * rendered when none has.
 */
class ConditionalFormattingsExtXform extends CompositeXform<ConditionalFormattingExtModel[]> {
    /** The xform of one conditional formatting. */
    readonly cfXform = new ConditionalFormattingExtXform();

    declare map: { readonly "x14:conditionalFormatting": ConditionalFormattingExtXform };

    constructor() {
        super();

        this.map = {
            "x14:conditionalFormatting": this.cfXform,
        };
    }

    get tag(): string {
        return "x14:conditionalFormattings";
    }

    /** @returns whether a rule of the conditional formattings is written in the extension */
    hasContent(model: readonly ConditionalFormattingExtModel[]): boolean {
        // not cached on the model: it is the worksheet's own list, which may change between writes
        return model.some((cf) => cf.rules.some(CfRuleExtXform.isExt));
    }

    /** Gives the x14 rules their `x14Id` (mutates the rules). */
    override prepare(model: readonly ConditionalFormattingExtModel[]): void {
        for (const cf of model) {
            this.cfXform.prepare(cf);
        }
    }

    override render(xmlStream: XmlStream, model: readonly ConditionalFormattingExtModel[]): void {
        if (this.hasContent(model)) {
            xmlStream.openNode(this.tag);
            for (const cf of model) {
                this.cfXform.render(xmlStream, cf);
            }
            xmlStream.closeNode();
        }
    }

    override createNewModel(): ConditionalFormattingExtModel[] {
        return [];
    }

    override onParserClose(): void {
        // model is array of conditional formatting objects; the element was opened before its children (createNewModel)
        const list = this.model;
        const cf = this.cfXform.model;
        if (list && cf) {
            list.push(cf);
        }
    }
}

export default ConditionalFormattingsExtXform;
