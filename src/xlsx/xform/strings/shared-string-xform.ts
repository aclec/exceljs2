import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { SharedStringValue } from "../../../types";
import TextXform from "./text-xform";
import RichTextXform from "./rich-text-xform";
import PhoneticTextXform from "./phonetic-text-xform";

import BaseXform, { ownEntry } from "../base-xform";

// <si>
//   <r></r><r></r>...
// </si>
// <si>
//   <t></t>
// </si>

/**
 * A shared string, `<si>` (ECMA-376 CT_Rst: `<t>?<r>*<rPh>*`): plain text or rich text runs. Phonetic runs are parsed and
 * dropped. Parsing: an empty `<si/>` (or one holding only phonetic runs) is `""`; a leading `<t>` followed by runs is
 * the first, unformatted run. Rendering skips the runs without text (an `<si>` left without runs gets an empty `<t>`).
 */
class SharedStringXform extends BaseXform<SharedStringValue> {
    declare map: { readonly r: RichTextXform; readonly t: TextXform; readonly rPh: PhoneticTextXform };

    /** @param model - the initial model */
    constructor(model?: SharedStringValue) {
        super();

        this.model = model;

        this.map = {
            r: new RichTextXform(),
            t: new TextXform(),
            rPh: new PhoneticTextXform(),
        };
    }

    /** The element: `si`. */
    get tag(): string {
        return "si";
    }

    /** @param model - `null` / `undefined` render an empty `<si/>` */
    override render(xmlStream: XmlStream, model: SharedStringValue | null | undefined): void {
        xmlStream.openNode(this.tag);
        if (typeof model === "string") {
            this.map.t.render(xmlStream, model);
        } else if (model) {
            // a rich text object without runs (or whose runs have no text) is an empty string
            const runs =
                Object.hasOwn(model, "richText") && model.richText ? model.richText.filter((run) => !RichTextXform.isEmpty(run)) : [];
            if (runs.length) {
                const { r } = this.map;
                for (const run of runs) {
                    r.render(xmlStream, run);
                }
            } else {
                this.map.t.render(xmlStream, "");
            }
        }
        xmlStream.closeNode();
    }

    /** @returns `true` for `<si>` and its content */
    override parseOpen(node: SaxTag): boolean {
        const { name } = node;
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        if (name === this.tag) {
            // nothing read yet: completed by the runs, or "" at the end
            this.model = undefined;
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

    /** @returns `false` at the end of `<si>` */
    override parseClose(name: string): boolean {
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                // a child closed (its model is complete)
                const run = this.map.r.model;
                const text = this.map.t.model;
                if (name === "r" && this.parser === this.map.r && run) {
                    const { model } = this;
                    if (typeof model === "string") {
                        // CT_Rst is <t>?<r>*: a leading <t> is an unformatted first run
                        this.model = { richText: [{ text: model }, run] };
                    } else if (model) {
                        model.richText.push(run);
                    } else {
                        this.model = { richText: [run] };
                    }
                } else if (name === "t" && this.parser === this.map.t && typeof text === "string") {
                    this.model = text;
                }
                this.parser = undefined;
            }
            return true;
        }
        switch (name) {
            case this.tag:
                if (this.model === undefined || this.model === null) {
                    // an empty <si/> (or one holding only phonetic runs) is the empty string
                    this.model = "";
                }
                return false;
            default:
                return true;
        }
    }
}

export default SharedStringXform;
