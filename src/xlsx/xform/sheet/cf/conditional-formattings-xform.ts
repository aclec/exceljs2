import type { SaxTag } from "../../../../utils/parse-sax";
import type XmlStream from "../../../../utils/xml-stream";
import type { CellPrepareOptions, ConditionalFormattingModel, WorksheetReconcileOptions } from "../../../../types";
import BaseXform from "../../base-xform";

import ConditionalFormattingXform from "./conditional-formatting-xform";

/**
 * The `<conditionalFormatting>` elements of a worksheet (they have no list element): the model is their array. The
 * worksheet xform delegates each `<conditionalFormatting>` to this xform, which appends it to the model (when it has a
 * range and a rule read); {@link reset} starts a new list.
 */
class ConditionalFormattingsXform extends BaseXform<ConditionalFormattingModel[]> {
    /** The xform of one `<conditionalFormatting>`. */
    readonly cfXform = new ConditionalFormattingXform();

    get tag(): string {
        return "conditionalFormatting";
    }

    /** Starts a new list (the model is `[]`, never `null`). */
    override reset(): void {
        this.model = [];
    }

    /**
     * Gives each rule without a priority one (after the highest one given), and registers the styles of the rules as
     * dxfs (`dxfId`). Mutates the rules.
     */
    override prepare(model: readonly ConditionalFormattingModel[], options: Pick<CellPrepareOptions, "styles">): void {
        // ensure each rule has a priority value, after the highest one given
        let nextPriority = 0;
        for (const cf of model) {
            for (const rule of cf.rules) {
                nextPriority = Math.max(nextPriority, rule.priority || 0);
            }
        }
        nextPriority++;
        for (const cf of model) {
            for (const rule of cf.rules) {
                if (!rule.priority) {
                    rule.priority = nextPriority++;
                }

                if (rule.style) {
                    rule.dxfId = options.styles.addDxfStyle(rule.style);
                }
            }
        }
    }

    override render(xmlStream: XmlStream, model: readonly ConditionalFormattingModel[]): void {
        for (const cf of model) {
            this.cfXform.render(xmlStream, cf);
        }
    }

    /** @returns `true` for `<conditionalFormatting>` and its descendants */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }

        switch (node.name) {
            case "conditionalFormatting":
                this.parser = this.cfXform;
                this.parser.parseOpen(node);
                return true;

            default:
                return false;
        }
    }

    override parseText(text: string): void {
        if (this.parser) {
            this.parser.parseText(text);
        }
    }

    /** @returns `false` at the end of a `<conditionalFormatting>` (appended to the model) */
    override parseClose(name: string): boolean {
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                // reset() created the list; the model of the element is set by its start tag. A <conditionalFormatting>
                // without sqref (malformed: its rules apply to no cell) or without a rule read (see CfRuleXform) is not
                // read
                const list = this.model;
                const cf = this.cfXform.model;
                if (list && cf && cf.ref !== undefined && cf.rules.length) {
                    list.push({ ref: cf.ref, rules: cf.rules });
                }
                this.parser = undefined;
                return false;
            }
            return true;
        }
        return false;
    }

    /** Turns the `dxfId` of the rules into their `style` (none when the workbook has no styles part). */
    override reconcile(model: readonly ConditionalFormattingModel[], options: Pick<WorksheetReconcileOptions, "styles">): void {
        for (const cf of model) {
            for (const rule of cf.rules) {
                if (rule.dxfId !== undefined) {
                    rule.style = options.styles?.getDxfStyle(rule.dxfId);
                    delete rule.dxfId;
                }
            }
        }
    }
}

export default ConditionalFormattingsXform;
