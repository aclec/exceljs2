import type { SaxTag } from "../../../utils/parse-sax";
import XmlStream, { type XmlAttributes } from "../../../utils/xml-stream";
import type { NoteModel } from "../../../types";
import BaseXform from "../base-xform";

import CommentXform from "./comment-xform";

/** The model of a comments part: its notes (rendered with their `ref`; parsed with the attributes of `<comment>`). */
export interface CommentsModel {
    comments: NoteModel[];
}

/**
 * A comments part, `xl/commentsN.xml`: one author (`Author`) and the notes of a worksheet. Parsing starts at
 * `<commentList>`.
 */
class CommentsXform extends BaseXform<CommentsModel> {
    /** The namespace of `<comments>`. */
    static readonly COMMENTS_ATTRIBUTES: XmlAttributes = {
        xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    };

    declare map: { readonly comment: CommentXform };

    constructor() {
        super();
        this.map = {
            comment: new CommentXform(),
        };
    }

    /** Writes the part. */
    override render(xmlStream: XmlStream, model: { readonly comments: readonly NoteModel[] }): void {
        const { comments } = model;
        xmlStream.openXml(XmlStream.StdDocAttributes);
        xmlStream.openNode("comments", CommentsXform.COMMENTS_ATTRIBUTES);

        // authors
        // TODO: support authors properly
        xmlStream.openNode("authors");
        xmlStream.leafNode("author", undefined, "Author");
        xmlStream.closeNode();

        // comments
        xmlStream.openNode("commentList");
        for (const comment of comments) {
            this.map.comment.render(xmlStream, comment);
        }
        xmlStream.closeNode();
        xmlStream.closeNode();
    }

    /** @returns `true` for `<commentList>` and the notes, `false` for other elements (skipped) */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        switch (node.name) {
            case "commentList":
                this.model = {
                    comments: [],
                };
                return true;
            case "comment":
                this.parser = this.map.comment;
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
            case "commentList":
                return false;
            case "comment": {
                // a <comment> outside <commentList> (no model) is ignored
                const note = this.map.comment.model;
                if (this.parser && this.model && note) {
                    this.model.comments.push(note);
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

export default CommentsXform;
