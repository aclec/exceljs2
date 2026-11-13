import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { RichTextRun } from "../../../types";
import TextXform from "./text-xform";
import RichTextXform from "./rich-text-xform";

import BaseXform, { ownEntry } from "../base-xform";

// <rPh sb="0" eb="1">
//   <t>(its pronounciation in KATAKANA)</t>
// </rPh>

/**
 * A phonetic run (`<rPh>`): the reading of the base text characters `sb`..`eb`, as text or rich text runs. Parsing gives
 * `NaN` for a missing or non-numeric `sb` / `eb`.
 */
export interface PhoneticTextModel {
    sb: number;
    eb: number;
    text?: string;
    richText?: RichTextRun[];
}

/** A phonetic run of a shared or inline string (`<rPh>`). Strings parse it but do not keep it. */
class PhoneticTextXform extends BaseXform<PhoneticTextModel> {
    declare map: { readonly r: RichTextXform; readonly t: TextXform };

    constructor() {
        super();

        this.map = {
            r: new RichTextXform(),
            t: new TextXform(),
        };
    }

    /** The element: `rPh`. */
    get tag(): string {
        return "rPh";
    }

    override render(xmlStream: XmlStream, model: PhoneticTextModel): void {
        xmlStream.openNode(this.tag, {
            sb: model.sb || 0,
            eb: model.eb || 0,
        });
        if (Object.hasOwn(model, "richText") && model.richText) {
            const { r } = this.map;
            for (const text of model.richText) {
                r.render(xmlStream, text);
            }
        } else if (model.text !== undefined) {
            this.map.t.render(xmlStream, model.text);
        } else {
            throw new Error("A phonetic run (<rPh>) needs its text or its rich text");
        }
        xmlStream.closeNode();
    }

    /** @returns `true` for `<rPh>` and its runs */
    override parseOpen(node: SaxTag): boolean {
        const { name } = node;
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        if (name === this.tag) {
            this.model = {
                sb: BaseXform.toIntValue(node.attributes.sb, NaN),
                eb: BaseXform.toIntValue(node.attributes.eb, NaN),
            };
            return true;
        }
        this.parser = ownEntry(this.map, name);
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        return false;
    }

    override parseText(text: string): void {
        if (this.parser) {
            this.parser.parseText(text);
        }
    }

    /** @returns `false` at the end of `<rPh>` */
    override parseClose(name: string): boolean {
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                // a run closed (a closed run has its model). An element nested in a run that ends it early (e.g.
                // <t>a<x/></t>) is ignored, as it always was; a run outside <rPh> (no model) too
                const { model } = this;
                const run = this.map.r.model;
                const text = this.map.t.model;
                if (!model) {
                    // outside <rPh>
                } else if (name === "r" && this.parser === this.map.r && run) {
                    let rt = model.richText;
                    if (!rt) {
                        rt = model.richText = [];
                    }
                    rt.push(run);
                } else if (name === "t" && this.parser === this.map.t && typeof text === "string") {
                    model.text = text;
                }
                this.parser = undefined;
            }
            return true;
        }
        switch (name) {
            case this.tag:
                return false;
            default:
                return true;
        }
    }
}

export default PhoneticTextXform;
