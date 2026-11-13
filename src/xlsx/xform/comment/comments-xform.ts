import type { SaxTag } from "../../../utils/parse-sax";
import XmlStream, { type XmlAttributes } from "../../../utils/xml-stream";
import type { NoteModel } from "../../../types";
import BaseXform from "../base-xform";
import { decodeXstring, encodeXstring } from "../strings/text-xform";

import CommentXform from "./comment-xform";

/** The model of a comments part: its notes (rendered with their `ref`; parsed with the attributes of `<comment>`). */
export interface CommentsModel {
    comments: NoteModel[];
}

// a valid authorId (xsd:unsignedInt, without sign): its digits
const AUTHOR_ID = /^[0-9]+$/;

/**
 * A comments part, `xl/commentsN.xml`: the authors and the notes of a worksheet.
 *
 * Rendering writes the distinct authors of the notes, in the order of their first note, and points the `authorId` of
 * each `<comment>` to its author: a note without author gets the placeholder {@link CommentsXform.DEFAULT_AUTHOR}.
 * Parsing reads `<authors>` (which comes before `<commentList>`) and sets the `author` of each note from its `authorId`:
 * a note whose `authorId` is not the index of an author, or whose author is the placeholder, has no author. Parsing ends
 * with `</commentList>`.
 */
class CommentsXform extends BaseXform<CommentsModel> {
    /** The namespace of `<comments>`. */
    static readonly COMMENTS_ATTRIBUTES: XmlAttributes = {
        xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    };

    /** The author written for the notes without author (and read as no author). */
    static readonly DEFAULT_AUTHOR = "Author";

    declare map: { readonly comment: CommentXform };

    /** While parsing: the authors read so far, by index. */
    private _authors: string[] = [];
    /** While parsing an `<author>`: its chunks of text (`undefined` outside `<author>`). */
    private _authorText: string[] | undefined;
    /** While parsing: true inside `<authors>`. */
    private _inAuthors = false;

    constructor() {
        super();
        this.map = {
            comment: new CommentXform(),
        };
    }

    /** Writes the part. */
    override render(xmlStream: XmlStream, model: { readonly comments: readonly NoteModel[] }): void {
        const { comments } = model;
        // the distinct authors, in the order of their first note (a Map: the names are user data)
        const authorIds = new Map<string, number>();
        const commentAuthorIds = comments.map((comment) => {
            const author = CommentsXform.authorOf(comment);
            let authorId = authorIds.get(author);
            if (authorId === undefined) {
                authorId = authorIds.size;
                authorIds.set(author, authorId);
            }
            return authorId;
        });

        xmlStream.openXml(XmlStream.StdDocAttributes);
        xmlStream.openNode("comments", CommentsXform.COMMENTS_ATTRIBUTES);

        xmlStream.openNode("authors");
        if (authorIds.size === 0) {
            // a part without notes (not written by the workbook writers) still names an author
            authorIds.set(CommentsXform.DEFAULT_AUTHOR, 0);
        }
        for (const author of authorIds.keys()) {
            xmlStream.leafNode("author", undefined, encodeXstring(author));
        }
        xmlStream.closeNode();

        xmlStream.openNode("commentList");
        comments.forEach((comment, index) => {
            this.map.comment.render(xmlStream, comment, commentAuthorIds[index] ?? 0);
        });
        xmlStream.closeNode();
        xmlStream.closeNode();
    }

    /**
     * The author written for a note: its `author`, or {@link CommentsXform.DEFAULT_AUTHOR}.
     *
     * @throws Error for an author that is not a string
     */
    static authorOf(comment: NoteModel): string {
        const { author } = comment.note;
        if (author === undefined) {
            return CommentsXform.DEFAULT_AUTHOR;
        }
        if (typeof author !== "string") {
            throw new Error(`The author of a note must be a string (note of ${String(comment.ref)})`);
        }
        return author;
    }

    /**
     * The author of a parsed note: the author at the index `authorId` (digits only), if any, unless it is
     * {@link CommentsXform.DEFAULT_AUTHOR}.
     */
    private _authorOf(authorId: string | undefined): string | undefined {
        if (authorId === undefined || !AUTHOR_ID.test(authorId)) {
            return undefined;
        }
        const author = this._authors[Number(authorId)];
        return author === CommentsXform.DEFAULT_AUTHOR ? undefined : author;
    }

    /** Forgets the last parse, and its authors. */
    override reset(): void {
        super.reset();
        this._authors = [];
        this._authorText = undefined;
        this._inAuthors = false;
    }

    /** @returns `true` for `<authors>`, `<commentList>` and their content, `false` for other elements (skipped) */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        switch (node.name) {
            case "authors":
                this._inAuthors = true;
                this._authors = [];
                return true;
            case "author":
                if (this._inAuthors) {
                    this._authorText = [];
                }
                return true;
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
        } else if (this._authorText) {
            this._authorText.push(text);
        }
    }

    override parseClose(name: string): boolean {
        switch (name) {
            case "commentList":
                return false;
            case "authors":
            case "author":
                if (this.parser) {
                    this.parser.parseClose(name);
                } else if (name === "authors") {
                    this._inAuthors = false;
                } else if (this._authorText) {
                    this._authors.push(decodeXstring(this._authorText.join("")));
                    this._authorText = undefined;
                }
                return true;
            case "comment": {
                // a <comment> outside <commentList> (no model) is ignored
                const note = this.map.comment.model;
                if (this.parser && this.model && note) {
                    const author = this._authorOf(note.authorId);
                    if (author !== undefined) {
                        note.note.author = author;
                    }
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
