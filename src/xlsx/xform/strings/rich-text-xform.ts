import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { RichTextRun } from "../../../types";
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
 * is not rendered (Excel treats it as a corrupt record). A parsed `<r>` without `<t>` (CT_RElt requires it) has the
 * text `""`.
 */
class RichTextXform extends BaseXform<RichTextRun> {
    /** The options of the run font xform. */
    static readonly FONT_OPTIONS = {
        tagName: "rPr",
        fontNameTag: "rFont",
    };

    private _textXform?: TextXform;
    private _fontXform?: FontXform;

    /** While parsing: the parts of the run read so far (in the order of their elements). */
    private run: Partial<RichTextRun> = {};

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
    override render(xmlStream: XmlStream, model: RichTextRun): void {
        const run = model;
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
                // its text and font are added when their elements close
                this.run = {};
                this.model = { text: "" };
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
                    const text = this.textXform.model;
                    if (typeof text === "string") {
                        this.run.text = text;
                        this.updateModel();
                    }
                }
                this.parser = undefined;
                return true;
            case "rPr": {
                // the font (an <rPr> misplaced inside <t> is ignored)
                const font = this.fontXform.model;
                if (this.parser && this.parser === this._fontXform && font) {
                    this.run.font = font;
                    this.updateModel();
                }
                this.parser = undefined;
                return true;
            }
            default:
                if (this.parser) {
                    this.parser.parseClose(name);
                }
                return true;
        }
    }

    // the model: the parts read so far (in the order of their elements), and the text "" until <t> is read. The callers
    // that delegate <r> read it without passing its end tag
    private updateModel(): void {
        const { run } = this;
        this.model = { ...run, text: run.text ?? "" };
    }
}

export default RichTextXform;
