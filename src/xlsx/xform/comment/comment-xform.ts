import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { NoteModel, RichTextRun } from "../../../types";
import RichTextXform from "../strings/rich-text-xform";
import BaseXform from "../base-xform";

/**
  <comment ref="B1" authorId="0">
    <text>
      <r>
        <rPr>
          <b/>
          <sz val="9"/>
          <rFont val="宋体"/>
          <charset val="134"/>
        </rPr>
        <t>51422:</t>
      </r>
      <r>
        <rPr>
          <sz val="9"/>
          <rFont val="宋体"/>
          <charset val="134"/>
        </rPr>
        <t xml:space="preserve">&#10;test</t>
      </r>
    </text>
  </comment>
 */

/**
 * A note of the comments part, `<comment ref authorId><text><r>...</r></text></comment>`: its rich text runs. Rendered
 * with `authorId="0"` (the only author written). Parsed as `{ type: "note", note: { texts }, ...the attributes of
 * <comment> }`: `type` and `note` are never taken from the attributes.
 */
class CommentXform extends BaseXform<NoteModel> {
    private _richTextXform?: RichTextXform;

    /** While parsing: the runs of the note (the `texts` of its model). */
    private texts: RichTextRun[] = [];

    /** The element of the items parsed by this xform: `r` (it is delegated `<comment>` by its parent). */
    get tag(): string {
        return "r";
    }

    /** The xform of the runs (created on first use). */
    get richTextXform(): RichTextXform {
        if (!this._richTextXform) {
            this._richTextXform = new RichTextXform();
        }
        return this._richTextXform;
    }

    /** Writes the note. */
    override render(xmlStream: XmlStream, note: NoteModel): void {
        xmlStream.openNode("comment", {
            ref: note.ref,
            authorId: 0,
        });
        xmlStream.openNode("text");
        if (note.note && note.note.texts) {
            for (const text of note.note.texts) {
                this.richTextXform.render(xmlStream, text);
            }
        }
        xmlStream.closeNode();
        xmlStream.closeNode();
    }

    /** @returns `true` for `<comment>` and its runs, `false` for other elements (`<text>`) */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        switch (node.name) {
            case "comment": {
                // the attributes (ref, authorId...) first: an attribute named "type" or "note" must not replace them
                const texts: RichTextRun[] = [];
                this.texts = texts;
                this.model = {
                    ...node.attributes,
                    type: "note",
                    note: {
                        texts,
                    },
                };
                return true;
            }
            case "r":
                this.parser = this.richTextXform;
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

    override parseClose(name: string): boolean {
        switch (name) {
            case "comment":
                return false;
            case "r": {
                // a run of a <comment> (texts is the array of its model)
                const run = this.richTextXform.model;
                if (this.parser && run) {
                    this.texts.push(run);
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
}

export default CommentXform;
