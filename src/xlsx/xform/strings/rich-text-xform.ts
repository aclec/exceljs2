import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { FontModel, RichTextRun } from "../../../types";
import TextXform from "./text-xform";
import FontXform from "../style/font-xform";

import BaseXform from "../base-xform";

// <r>
//   <rPr>
//     <sz val="11"/>
//     <color theme="1" tint="5"/>
//     <rFont val="Calibri"/>
//     <family val="2"/>
//     <scheme val="minor"/>
//   </rPr>
//   <t xml:space="preserve"> is </t>
// </r>

/**
 * A run of rich text, `<r>`: its font (`<rPr>`, a font whose name element is `rFont`) and its text. A run without text
 * is not rendered (Excel treats it as a corrupt record). A parsed `<r>` without `<t>` has no `text`.
 */
class RichTextXform extends BaseXform<RichTextRun> {
    /** The options of the run font xform. */
    static readonly FONT_OPTIONS = {
        tagName: "rPr",
        fontNameTag: "rFont",
    };

    private _textXform?: TextXform;
    private _fontXform?: FontXform;

    /** @param model - the run rendered when {@link render} gets none */
    constructor(model?: RichTextRun) {
        super();

        this.model = model;
    }

    /** The element: `r`. */
    get tag(): string {
        return "r";
    }

    /** The xform of the text (created on first use). */
    get textXform(): TextXform {
        return this._textXform || (this._textXform = new TextXform());
    }

    /** The xform of the font (created on first use). */
    get fontXform(): FontXform {
        return this._fontXform || (this._fontXform = new FontXform(RichTextXform.FONT_OPTIONS));
    }

    /** Excel treats a run without text as a corrupt record. */
    static isEmpty(model: Partial<RichTextRun>): boolean {
        return model.text === undefined || model.text === null || model.text === "";
    }

    /** Writes the run (nothing when it has no text). */
    override render(xmlStream: XmlStream, model?: RichTextRun): void {
        const run = model || this.model!;
        if (RichTextXform.isEmpty(run)) {
            return;
        }

        xmlStream.openNode("r");
        if (run.font) {
            this.fontXform.render(xmlStream, run.font);
        }
        this.textXform.render(xmlStream, run.text);
        xmlStream.closeNode();
    }

    /** @returns `true` for `<r>` and its content */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        switch (node.name) {
            case "r":
                // its text and font are added when their elements close (a run without <t> has no text)
                this.model = {} as RichTextRun;
                return true;
            case "t":
                this.parser = this.textXform;
                this.parser.parseOpen(node);
                return true;
            case "rPr":
                this.parser = this.fontXform;
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

    /** @returns `false` at the end of `<r>` */
    override parseClose(name: string): boolean {
        switch (name) {
            case "r":
                return false;
            case "t":
                // the text, completed by the text xform's parseClose (a <t> misplaced inside <rPr> is ignored)
                if (this.parser && this.parser === this._textXform) {
                    this.textXform.parseClose();
                    this.model!.text = this.textXform.model!;
                }
                this.parser = undefined;
                return true;
            case "rPr":
                // the font (an <rPr> misplaced inside <t> is ignored). FontXform is not typed yet: its model is unknown
                if (this.parser && this.parser === this._fontXform) {
                    this.model!.font = this.fontXform.model as FontModel;
                }
                this.parser = undefined;
                return true;
            default:
                if (this.parser) {
                    this.parser.parseClose(name);
                }
                return true;
        }
    }
}

export default RichTextXform;
